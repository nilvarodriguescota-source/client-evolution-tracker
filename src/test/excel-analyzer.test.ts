import { describe, it, expect } from "vitest";
import * as XLSX from "xlsx";
import { analyzeWorkbook, ordenarPeriodosCronologicamente } from "../lib/excel-analyzer";

/**
 * Builds a workbook in the "Vendas de produto no período" report layout,
 * one sheet per period.
 *
 * @param periods sheet name -> client names present in that period
 */
function buildWorkbook(
  periods: Record<string, string[]>,
  bookType: XLSX.BookType = "xlsx",
  /** sheet name -> product sold in it (defaults to PRODUTO A) */
  produtoPorAba: Record<string, string> = {},
): ArrayBuffer {
  const wb = XLSX.utils.book_new();

  for (const [sheetName, clientes] of Object.entries(periods)) {
    const produto = produtoPorAba[sheetName] ?? "PRODUTO A";
    const rows: (string | number | null)[][] = [];
    clientes.forEach((nome, i) => {
      rows.push([`Cliente: ${100 + i} - ${nome}`, null, null, null, null, null]);
      rows.push(["Código", "Descrição", "Unidade", "Qtd", "Preço Médio", "Venda Bruta"]);
      // qty 2 x price 10 = 20
      rows.push(["5001", produto, "UN", 2, 10, 20]);
      rows.push([null, null, null, 2, null, 20]);
      rows.push([]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), sheetName);
  }

  const out = XLSX.write(wb, { type: "array", bookType });
  return out instanceof ArrayBuffer ? out : new Uint8Array(out).buffer;
}

describe("analyzeWorkbook — janela de análise da carteira", () => {
  it("lê todas as abas de um arquivo .xls legado (regressão: ArrayBuffer)", () => {
    // Legacy .xls is a CFB container. Passing a raw ArrayBuffer to SheetJS used
    // to yield a single bogus "Sheet1", collapsing the analysis to one period.
    const buf = buildWorkbook(
      { junho: ["CLIENTE A"], julho: ["CLIENTE A"], agosto: ["CLIENTE A"] },
      "biff8",
    );
    const r = analyzeWorkbook(buf);

    expect(r.periodos.map((p) => p.nome)).toEqual(["junho", "julho", "agosto"]);
  });

  it("lê todos os períodos do arquivo, sem teto de abas", () => {
    const nomes = ["m1", "m2", "m3", "m4", "m5", "m6", "m7", "m8"];
    const periods = Object.fromEntries(nomes.map((n) => [n, ["CLIENTE A"]]));
    const r = analyzeWorkbook(buildWorkbook(periods));

    // every tab of the file is a period — nothing is dropped for being old
    expect(r.periodos.map((p) => p.nome)).toEqual(nomes);
    expect(r.leitura.abasLidas).toEqual(nomes);
    expect(r.leitura.abasIgnoradas).toEqual([]);
    // the 2-period comparison is untouched
    expect(r.periodoAnterior.nome).toBe("m7");
    expect(r.periodoAtual.nome).toBe("m8");
  });

  it("usa todas as abas quando há poucas", () => {
    const r = analyzeWorkbook(
      buildWorkbook({ junho: ["CLIENTE A"], julho: ["CLIENTE A"], agosto: ["CLIENTE A"] }),
    );
    expect(r.periodos).toHaveLength(3);
  });

  it("classifica a carteira pela distância do último período de compra", () => {
    // 6 períodos m1..m6; a distância é contada a partir de m6 (período atual):
    // m6=0 ativo · m5=1 reduzida · m4=2 e m3=3 risco · m2=4 e m1=5 inativo
    const r = analyzeWorkbook(
      buildWorkbook({
        m1: ["INATIVO_5"],
        m2: ["INATIVO_4"],
        m3: ["RISCO_3"],
        m4: ["RISCO_2"],
        m5: ["REDUZIDA"],
        m6: ["ATIVO"],
      }),
    );

    const byName = Object.fromEntries(r.classificacaoCarteira.map((c) => [c.nome, c]));

    expect(byName["ATIVO"].status).toBe("ativo");
    expect(byName["REDUZIDA"].status).toBe("frequencia_reduzida");
    expect(byName["RISCO_2"].status).toBe("risco_abandono");
    expect(byName["RISCO_3"].status).toBe("risco_abandono");
    expect(byName["INATIVO_4"].status).toBe("inativo");
    expect(byName["INATIVO_5"].status).toBe("inativo");
  });

  it("não marca como perdido o cliente sazonal que voltou a comprar", () => {
    // Present in the oldest period, absent for months, back in the current one.
    // The 2-period view calls it "novo"; the window view must call it active
    // and show the full history.
    const r = analyzeWorkbook(
      buildWorkbook({
        m1: ["SAZONAL"],
        m2: ["OUTRO"],
        m3: ["OUTRO"],
        m4: ["OUTRO"],
        m5: ["OUTRO"],
        m6: ["SAZONAL", "OUTRO"],
      }),
    );

    const sazonal = r.classificacaoCarteira.find((c) => c.nome === "SAZONAL")!;
    expect(sazonal.status).toBe("ativo");
    expect(sazonal.periodosPresente).toEqual(["m1", "m6"]);
    expect(sazonal.totalPeriodosPresente).toBe(2);
  });

  it("mantém na carteira clientes ausentes das duas últimas abas", () => {
    // ANTIGO appears only in m1 — invisible to the 2-period comparison entirely.
    const r = analyzeWorkbook(
      buildWorkbook({
        m1: ["ANTIGO", "ATUAL"],
        m2: ["ATUAL"],
        m3: ["ATUAL"],
      }),
    );

    expect(r.clientesNovos).toHaveLength(0);
    expect(r.clientesPerdidos).toHaveLength(0);
    // but the portfolio view still knows about ANTIGO
    expect(r.classificacaoCarteira.map((c) => c.nome).sort()).toEqual(["ANTIGO", "ATUAL"]);
    expect(r.classificacaoCarteira.find((c) => c.nome === "ANTIGO")!.status).toBe(
      "risco_abandono",
    );
  });

  it("rejeita planilha com menos de 2 abas informando as abas encontradas", () => {
    expect(() => analyzeWorkbook(buildWorkbook({ unica: ["CLIENTE A"] }))).toThrow(
      /pelo menos 2 abas.*unica/s,
    );
  });
});

describe("ordenarPeriodosCronologicamente", () => {
  it("ordena abas fora de ordem pelo mês", () => {
    expect(ordenarPeriodosCronologicamente(["agosto", "junho", "julho"])).toEqual([
      "junho",
      "julho",
      "agosto",
    ]);
  });

  it("preserva abas já cronológicas", () => {
    expect(ordenarPeriodosCronologicamente(["junho", "julho", "agosto"])).toEqual([
      "junho",
      "julho",
      "agosto",
    ]);
  });

  it("atravessa a virada de ano sem ano explícito", () => {
    expect(
      ordenarPeriodosCronologicamente(["janeiro", "novembro", "dezembro"]),
    ).toEqual(["novembro", "dezembro", "janeiro"]);
  });

  it("usa o ano quando informado", () => {
    expect(
      ordenarPeriodosCronologicamente(["janeiro/2026", "dezembro/2025", "novembro/2025"]),
    ).toEqual(["novembro/2025", "dezembro/2025", "janeiro/2026"]);
  });

  it("aceita abreviações, acentos e formato numérico", () => {
    expect(ordenarPeriodosCronologicamente(["mai", "MARÇO", "Abril"])).toEqual([
      "MARÇO",
      "Abril",
      "mai",
    ]);
    expect(ordenarPeriodosCronologicamente(["2025-08", "2025-06", "2025-07"])).toEqual([
      "2025-06",
      "2025-07",
      "2025-08",
    ]);
  });

  it("ordena os meses mesmo com uma aba que não nomeia mês no meio", () => {
    // "Resumo" keeps its slot; junho and agosto are ordered around it instead of
    // one unrecognised tab disabling the ordering for every month.
    expect(ordenarPeriodosCronologicamente(["agosto", "Resumo", "junho"])).toEqual([
      "junho",
      "Resumo",
      "agosto",
    ]);
  });
});

describe("carteira — semântica de janela completa", () => {
  it("separa clientes novos de clientes que retornaram", () => {
    const r = analyzeWorkbook(
      buildWorkbook({
        junho: ["ANTIGO", "SAZONAL"],
        julho: ["ANTIGO"],
        agosto: ["ANTIGO", "SAZONAL", "NOVO"],
      }),
    );

    // month over month, both SAZONAL and NOVO look new
    expect(r.clientesNovos.map((c) => c.nome).sort()).toEqual(["NOVO", "SAZONAL"]);
    // over the window, only NOVO never bought before
    expect(r.carteira.novos.map((c) => c.nome)).toEqual(["NOVO"]);
    expect(r.carteira.retornados.map((c) => c.nome)).toEqual(["SAZONAL"]);
  });

  it("só considera perdido quem está 2+ meses sem comprar", () => {
    const r = analyzeWorkbook(
      buildWorkbook({
        junho: ["SUMIU_CEDO", "FALTOU_UM_MES", "ATIVO"],
        julho: ["FALTOU_UM_MES", "ATIVO"],
        agosto: ["ATIVO"],
      }),
    );

    // month over month lumps both absences together
    expect(r.clientesPerdidos.map((c) => c.nome)).toEqual(["FALTOU_UM_MES"]);
    // over the window they are told apart
    expect(r.carteira.perdidos.map((c) => c.nome)).toEqual(["SUMIU_CEDO"]);
    expect(r.carteira.frequenciaReduzida.map((c) => c.nome)).toEqual(["FALTOU_UM_MES"]);
  });

  it("soma a receita de novos e de perdidos coerentemente", () => {
    const r = analyzeWorkbook(
      buildWorkbook({
        junho: ["PERDIDO"],
        julho: ["ATIVO"],
        agosto: ["ATIVO", "NOVO"],
      }),
    );
    // each client in the fixture bills 20 in its period
    expect(r.carteira.faturamentoNovos).toBe(20);
    expect(r.carteira.faturamentoPerdidos).toBe(20);
    expect(r.carteira.ativos.map((c) => c.nome).sort()).toEqual(["ATIVO", "NOVO"]);
  });

  it("a janela é montada pelos meses mais recentes, não pelas últimas abas", () => {
    const r = analyzeWorkbook(
      buildWorkbook({
        agosto: ["A"],
        marco: ["A"],
        junho: ["A"],
        julho: ["A"],
      }),
    );
    expect(r.periodos.map((p) => p.nome)).toEqual(["marco", "junho", "julho", "agosto"]);
    expect(r.periodoAtual.nome).toBe("agosto");
    expect(r.periodoAnterior.nome).toBe("julho");
  });
});

describe("evolucaoPeriodos — comparativo mês a mês", () => {
  it("cobre todos os meses da janela, com o primeiro sem variação", () => {
    const r = analyzeWorkbook(
      buildWorkbook({
        junho: ["A", "B"],
        julho: ["A", "B", "C"],
        agosto: ["A"],
      }),
    );

    expect(r.evolucaoPeriodos.map((p) => p.nome)).toEqual(["junho", "julho", "agosto"]);

    const [jun, jul, ago] = r.evolucaoPeriodos;

    // each client in the fixture bills 20
    expect(jun.faturamento).toBe(40);
    expect(jul.faturamento).toBe(60);
    expect(ago.faturamento).toBe(20);

    expect(jun.variacaoFaturamento).toBeNull();
    expect(jun.entraram).toBeNull();
    expect(jun.sairam).toBeNull();

    expect(jul.variacaoFaturamento).toBe(50);
    expect(jul.entraram).toBe(1); // C
    expect(jul.sairam).toBe(0);

    expect(ago.variacaoFaturamento).toBeCloseTo(-66.6667, 3);
    expect(ago.entraram).toBe(0);
    expect(ago.sairam).toBe(2); // B, C
  });

  it("calcula ticket médio e variação de clientes por mês", () => {
    const r = analyzeWorkbook(
      buildWorkbook({ junho: ["A", "B"], julho: ["A", "B", "C", "D"] }),
    );
    const [jun, jul] = r.evolucaoPeriodos;

    expect(jun.clientes).toBe(2);
    expect(jun.ticketMedio).toBe(20);
    expect(jun.variacaoClientes).toBeNull();

    expect(jul.clientes).toBe(4);
    expect(jul.ticketMedio).toBe(20);
    expect(jul.variacaoClientes).toBe(100);
  });

  it("expõe o faturamento de cada produto em todos os meses", () => {
    const r = analyzeWorkbook(
      buildWorkbook({ junho: ["A"], julho: ["A"], agosto: ["A"] }),
    );
    const prod = r.comparativoProdutos[0];

    expect(Object.keys(prod.valorPorPeriodo)).toEqual(["junho", "julho", "agosto"]);
    expect(prod.valorPorPeriodo).toEqual({ junho: 20, julho: 20, agosto: 20 });
    expect(prod.qtdPorPeriodo).toEqual({ junho: 2, julho: 2, agosto: 2 });
    // the pre-existing anterior/atual fields keep their meaning
    expect(prod.valorAnterior).toBe(20);
    expect(prod.valorAtual).toBe(20);
  });

  it("zera o produto nos meses em que não houve venda", () => {
    const r = analyzeWorkbook(
      buildWorkbook({ junho: ["A"], julho: ["A"], agosto: ["A"] }, "xlsx", {
        junho: "PRODUTO B",
      }),
    );

    const a = r.comparativoProdutos.find((p) => p.produto === "PRODUTO A")!;
    expect(a.valorPorPeriodo.junho).toBe(0);
    expect(a.valorPorPeriodo.agosto).toBe(20);

    // sold only in junho — it must still be in the comparison, not disappear
    // because the last two months don't carry it
    const b = r.comparativoProdutos.find((p) => p.produto === "PRODUTO B")!;
    expect(b).toBeDefined();
    expect(b.valorPorPeriodo).toEqual({ junho: 20, julho: 0, agosto: 0 });
    expect(b.valorAtual).toBe(0);
  });
});
