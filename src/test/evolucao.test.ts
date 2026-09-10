import { describe, it, expect } from "vitest";
import { analyzeWorkbook, type AnalysisResult } from "../lib/excel-analyzer";
import { montarArquivo, type AbaFixture, type ClienteFixture } from "./relatorio-fixture";

/**
 * The analysis has to be comparative: every month measured against the one
 * before it, and the whole window read as one movement. These tests pin the
 * arithmetic of that model for 3 months and for 6.
 */

/** One client buying a single product worth `valor` (qty 1). */
const cli = (nome: string, valor: number): ClienteFixture => ({
  codigo: nome,
  nome,
  itens: [{ codigo: "5001", produto: "PRODUTO A", qtd: 1, preco: valor }],
});

const abas = (meses: Record<string, ClienteFixture[]>): Record<string, AbaFixture> =>
  Object.fromEntries(Object.entries(meses).map(([m, clientes]) => [m, { clientes }]));

//        junho          julho              agosto
//  ALFA   100            120 (+20)          90 (-30)
//  BETA   200            —   (perdido)      210 (retornado)
//  GAMA   —              50  (novo)         50 (mantido)
//  DELTA  —              —                  70 (novo)
//  total  300            170                420
const tresMeses = () =>
  analyzeWorkbook(
    montarArquivo(
      abas({
        junho: [cli("ALFA", 100), cli("BETA", 200)],
        julho: [cli("ALFA", 120), cli("GAMA", 50)],
        agosto: [cli("ALFA", 90), cli("BETA", 210), cli("GAMA", 50), cli("DELTA", 70)],
      }),
    ),
  );

/** Client names of a movement group, in the order the analysis returns them. */
const nomes = (g: { clientes: { nome: string }[] }) => g.clientes.map((c) => c.nome);

describe("transições mês a mês", () => {
  it("gera uma transição para cada par de meses consecutivos", () => {
    const r = tresMeses();
    expect(r.transicoes.map((t) => `${t.de}→${t.para}`)).toEqual([
      "junho→julho",
      "julho→agosto",
    ]);
  });

  it("classifica novo, mantido, perdido e retornado olhando todo o histórico", () => {
    const [t1, t2] = tresMeses().transicoes;

    // junho → julho
    expect(nomes(t1.mantidos)).toEqual(["ALFA"]);
    expect(nomes(t1.novos)).toEqual(["GAMA"]);
    expect(nomes(t1.perdidos)).toEqual(["BETA"]);
    expect(nomes(t1.retornados)).toEqual([]);

    // julho → agosto: BETA já tinha comprado em junho, logo retornou, não é novo
    expect(nomes(t2.mantidos)).toEqual(["ALFA", "GAMA"]);
    expect(nomes(t2.retornados)).toEqual(["BETA"]);
    expect(nomes(t2.novos)).toEqual(["DELTA"]);
    expect(nomes(t2.perdidos)).toEqual([]);
  });

  it("mede faturamento, clientes e ticket médio de cada transição", () => {
    const [t1, t2] = tresMeses().transicoes;

    expect(t1.faturamentoDe).toBe(300);
    expect(t1.faturamentoPara).toBe(170);
    expect(t1.diferencaFaturamento).toBe(-130);
    expect(t1.variacaoFaturamento).toBeCloseTo((-130 / 300) * 100, 4);
    expect(t1.clientesDe).toBe(2);
    expect(t1.clientesPara).toBe(2);
    expect(t1.variacaoClientes).toBe(0);
    expect(t1.ticketMedioDe).toBe(150);
    expect(t1.ticketMedioPara).toBe(85);
    expect(t1.taxaRetencao).toBe(50); // 1 de 2 clientes mantidos

    expect(t2.faturamentoPara).toBe(420);
    expect(t2.diferencaFaturamento).toBe(250);
    expect(t2.taxaRetencao).toBe(100); // ninguém saiu
  });

  it("atribui a receita de cada grupo ao mês em que ela existiu", () => {
    const [t1, t2] = tresMeses().transicoes;

    expect(t1.novos.faturamento).toBe(50); // GAMA em julho
    expect(t1.perdidos.faturamento).toBe(200); // BETA, medido em junho
    expect(t1.mantidos.faturamento).toBe(120); // ALFA em julho

    expect(t2.retornados.faturamento).toBe(210); // BETA em agosto
    expect(t2.novos.faturamento).toBe(70); // DELTA em agosto
    expect(t2.mantidos.faturamento).toBe(140); // ALFA 90 + GAMA 50
  });

  it("explica a variação do mês pela ponte de receita", () => {
    for (const t of tresMeses().transicoes) {
      const { novos, retornados, expansao, contracao, perdidos, total } = t.ponte;
      // as cinco parcelas explicam integralmente a diferença do mês
      expect(novos + retornados + expansao + contracao + perdidos).toBeCloseTo(total, 4);
      expect(total).toBeCloseTo(t.diferencaFaturamento, 4);
    }

    const [t1, t2] = tresMeses().transicoes;
    expect(t1.ponte).toEqual({
      novos: 50,
      retornados: 0,
      expansao: 20, // ALFA 100 → 120
      contracao: 0,
      perdidos: -200, // BETA
      total: -130,
    });
    expect(t2.ponte).toEqual({
      novos: 70,
      retornados: 210,
      expansao: 0,
      contracao: -30, // ALFA 120 → 90
      perdidos: 0,
      total: 250,
    });
  });
});

describe("resumo mês a mês", () => {
  it("traz a movimentação de clientes em cada mês, sem variação no primeiro", () => {
    const [junho, julho, agosto] = tresMeses().evolucaoPeriodos;

    expect(junho.variacaoFaturamento).toBeNull();
    expect(junho.novos).toBeNull();
    expect(junho.mantidos).toBeNull();

    expect(julho.novos).toBe(1);
    expect(julho.mantidos).toBe(1);
    expect(julho.perdidos).toBe(1);
    expect(julho.retornados).toBe(0);
    expect(julho.faturamentoNovos).toBe(50);
    expect(julho.faturamentoPerdidos).toBe(200);
    expect(julho.taxaRetencao).toBe(50);

    expect(agosto.novos).toBe(1);
    expect(agosto.retornados).toBe(1);
    expect(agosto.mantidos).toBe(2);
    expect(agosto.entraram).toBe(2); // novos + retornados
    expect(agosto.perdidos).toBe(0);
  });

  it("mantém o resumo mês a mês igual às transições", () => {
    const r = tresMeses();
    r.evolucaoPeriodos.slice(1).forEach((p, i) => {
      const t = r.transicoes[i];
      expect(p.faturamento).toBe(t.faturamentoPara);
      expect(p.variacaoFaturamento).toBe(t.variacaoFaturamento);
      expect(p.novos).toBe(t.novos.total);
      expect(p.perdidos).toBe(t.perdidos.total);
    });
  });
});

describe("evolução consolidada", () => {
  it("lê a janela inteira como um movimento só", () => {
    const { evolucaoConsolidada: e } = tresMeses();

    expect(e.meses).toBe(3);
    expect(e.primeiro).toBe("junho");
    expect(e.ultimo).toBe("agosto");
    expect(e.faturamentoInicial).toBe(300);
    expect(e.faturamentoFinal).toBe(420);
    expect(e.variacaoTotal).toBeCloseTo(40, 4);
    expect(e.tendencia).toBe("crescimento");
    expect(e.melhorMes).toEqual({ nome: "agosto", faturamento: 420 });
    expect(e.piorMes).toEqual({ nome: "julho", faturamento: 170 });
    expect(e.mesesEmAlta).toBe(1);
    expect(e.mesesEmQueda).toBe(1);

    expect(e.clientesInicial).toBe(2);
    expect(e.clientesFinal).toBe(4);
    expect(e.clientesDistintos).toBe(4);
    expect(e.clientesRecorrentes).toBe(1); // só ALFA comprou nos três meses
    expect(e.novosNoPeriodo).toBe(2); // GAMA e DELTA
  });

  it("soma o total da janela sem substituir a leitura comparativa", () => {
    const { evolucaoConsolidada: e, evolucaoPeriodos } = tresMeses();
    expect(e.faturamentoTotal).toBe(300 + 170 + 420);
    expect(e.faturamentoMedio).toBeCloseTo((300 + 170 + 420) / 3, 4);
    // o total é contexto: os meses continuam existindo separadamente
    expect(evolucaoPeriodos.map((p) => p.faturamento)).toEqual([300, 170, 420]);
  });

  it("marca estabilidade quando a variação da janela é pequena", () => {
    const r = analyzeWorkbook(
      montarArquivo(abas({ junho: [cli("A", 100)], julho: [cli("A", 98)], agosto: [cli("A", 102)] })),
    );
    expect(r.evolucaoConsolidada.tendencia).toBe("estavel");
  });
});

describe("janela de 6 meses", () => {
  const seisMeses = () =>
    analyzeWorkbook(
      montarArquivo(
        abas({
          janeiro: [cli("A", 100), cli("B", 100)],
          fevereiro: [cli("A", 110), cli("B", 90)],
          marco: [cli("A", 120), cli("C", 60)],
          abril: [cli("A", 130), cli("C", 70), cli("D", 40)],
          maio: [cli("A", 140), cli("D", 50)],
          junho: [cli("A", 150), cli("B", 120), cli("D", 60)],
        }),
      ),
    );

    it("aplica a mesma lógica automaticamente às seis abas", () => {
    const r = seisMeses();

    expect(r.periodos).toHaveLength(6);
    expect(r.transicoes).toHaveLength(5);
    expect(r.evolucaoPeriodos.map((p) => p.nome)).toEqual([
      "janeiro", "fevereiro", "marco", "abril", "maio", "junho",
    ]);
    expect(r.evolucaoConsolidada.meses).toBe(6);
    expect(r.evolucaoConsolidada.primeiro).toBe("janeiro");
    expect(r.evolucaoConsolidada.ultimo).toBe("junho");

    // B sumiu em março e voltou em junho: retornado, nunca novo
    const ultima = r.transicoes[4];
    expect(nomes(ultima.retornados)).toEqual(["B"]);
    expect(nomes(ultima.novos)).toEqual([]);
    expect(nomes(ultima.mantidos)).toEqual(["A", "D"]);
    expect(nomes(ultima.perdidos)).toEqual([]);

    // C comprou em março e abril e parou: perdido na transição abril → maio
    expect(nomes(r.transicoes[3].perdidos)).toEqual(["C"]);

    // cada mês continua com o seu próprio faturamento
    expect(r.evolucaoPeriodos.map((p) => p.faturamento)).toEqual([
      200, 200, 180, 240, 190, 330,
    ]);
  });

  it("mantém a ponte de receita fechada em todas as transições", () => {
    for (const t of seisMeses().transicoes) {
      const p = t.ponte;
      expect(p.novos + p.retornados + p.expansao + p.contracao + p.perdidos).toBeCloseTo(
        t.diferencaFaturamento,
        4,
      );
    }
  });

  it("acompanha a tendência de cada cliente ao longo da janela", () => {
    const porNome = Object.fromEntries(
      seisMeses().classificacaoCarteira.map((c) => [c.nome, c]),
    );

    expect(porNome.A.tendencia).toBe("crescimento"); // 100 → 150
    expect(porNome.A.totalPeriodosPresente).toBe(6);
    expect(Object.keys(porNome.A.faturamentoPorPeriodo)).toHaveLength(6);
    expect(porNome.C.faturamentoPorPeriodo).toEqual({
      janeiro: 0, fevereiro: 0, marco: 60, abril: 70, maio: 0, junho: 0,
    });
    // medida entre a primeira e a última compra do próprio cliente
    expect(porNome.C.variacaoPeriodo).toBeCloseTo(16.6667, 3);
  });

  it("acompanha a tendência de cada produto ao longo da janela", () => {
    const r = analyzeWorkbook(
      montarArquivo({
        junho: { clientes: [{ codigo: "1", nome: "A", itens: [
          { codigo: "5001", produto: "SOBE", qtd: 1, preco: 100 },
          { codigo: "5002", produto: "DESCE", qtd: 1, preco: 100 },
        ] }] },
        julho: { clientes: [{ codigo: "1", nome: "A", itens: [
          { codigo: "5001", produto: "SOBE", qtd: 1, preco: 150 },
          { codigo: "5002", produto: "DESCE", qtd: 1, preco: 60 },
        ] }] },
        agosto: { clientes: [{ codigo: "1", nome: "A", itens: [
          { codigo: "5001", produto: "SOBE", qtd: 1, preco: 200 },
        ] }] },
      }),
    );

    const porProduto = Object.fromEntries(r.comparativoProdutos.map((p) => [p.produto, p]));
    expect(porProduto.SOBE.tendencia).toBe("crescimento");
    expect(porProduto.SOBE.variacaoPeriodo).toBeCloseTo(100, 4);
    expect(porProduto.SOBE.mesesComVenda).toBe(3);
    expect(porProduto.DESCE.tendencia).toBe("queda");
    expect(porProduto.DESCE.mesesComVenda).toBe(2);
    expect(porProduto.DESCE.variacaoPeriodo).toBe(-100);
  });
});

describe("coerência entre resumo e dados dos meses", () => {
  it("o faturamento de cada mês é a soma dos clientes daquele mês", () => {
    const r: AnalysisResult = tresMeses();
    r.periodos.forEach((p, i) => {
      const soma = [...p.clientes.values()].reduce((s, c) => s + c.totalGeral, 0);
      expect(r.evolucaoPeriodos[i].faturamento).toBeCloseTo(soma, 4);
      expect(r.evolucaoPeriodos[i].clientes).toBe(p.clientes.size);
    });
  });
});

describe("tendência com um único mês", () => {
  it("não inventa direção para quem aparece uma vez só", () => {
    const r = analyzeWorkbook(
      montarArquivo(
        abas({
          junho: [cli("ANTIGO", 100)],
          julho: [cli("ANTIGO", 100)],
          agosto: [cli("ANTIGO", 100), cli("ESTREANTE", 50)],
        }),
      ),
    );

    const estreante = r.classificacaoCarteira.find((c) => c.nome === "ESTREANTE")!;
    expect(estreante.totalPeriodosPresente).toBe(1);
    expect(estreante.variacaoPeriodo).toBeNull();
  });
});
