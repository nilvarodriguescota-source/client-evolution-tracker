import { describe, it, expect, vi } from "vitest";
import { render, screen, within, fireEvent } from "@testing-library/react";
import { cloneElement, type ReactElement } from "react";
import { AnalysisResults } from "../components/AnalysisResults";
import { analyzeWorkbook } from "../lib/excel-analyzer";
import { lerCabecalhoCliente, separarCamposCliente } from "../lib/sheet-discovery";
import { montarArquivo, type AbaFixture, type ClienteFixture } from "./relatorio-fixture";

vi.mock("recharts", async () => {
  const actual = await vi.importActual<typeof import("recharts")>("recharts");
  return {
    ...actual,
    ResponsiveContainer: ({ children }: { children: ReactElement }) =>
      cloneElement(children, { width: 800, height: 400 }),
  };
});

/**
 * The report prints the client line as
 *   "Cliente: 1045 - BRASA - Data da última compra: 08/09/2026 - Fone: … - Celular: …"
 * and the last purchase date changes from one month to the next.
 */
const cli = (
  codigo: string,
  nome: string,
  valor: number,
  dataUltimaCompra: string,
  contato?: { fone?: string; celular?: string },
): ClienteFixture => ({
  codigo,
  nome,
  dataUltimaCompra,
  fone: contato?.fone,
  celular: contato?.celular,
  itens: [{ codigo: "5001", produto: "PRODUTO A", qtd: 1, preco: valor }],
});

const BRASA = (v: number, d: string) =>
  cli("1045", "BRASA", v, d, { fone: "(11) 3333-3333", celular: "(11) 99999-9999" });
const ALFREDO = (v: number, d: string) => cli("1046", "ALFREDO HAMBÚRGUER ARTESANAL", v, d);
const AMERICAN = (v: number, d: string) => cli("1047", "AMERICAN EMBER", v, d);
const ANGELICO = (v: number, d: string) =>
  cli("1048", "ANGELICO HORTIFRUTI", v, d, { celular: "(11) 98888-7777" });
const ARAMADO = (v: number, d: string) => cli("1049", "ARAMADO REST", v, d);

const abas = (meses: Record<string, ClienteFixture[]>): Record<string, AbaFixture> =>
  Object.fromEntries(
    Object.entries(meses).map(([m, clientes]) => [m, { clientes, intervalo: undefined }]),
  );

//        junho                julho                     agosto
//  BRASA    1000               1100                      1200
//  ALFREDO  500                — (saiu)                  600 (retorno)
//  AMERICAN —                  300 (novo)                350
//  ANGELICO —                  —                         800 (novo)
//  ARAMADO  200                — (saiu)                  — (segue fora)
const ARQUIVO = () =>
  montarArquivo(
    abas({
      junho: [BRASA(1000, "30/06/2026"), ALFREDO(500, "28/06/2026"), ARAMADO(200, "03/06/2026")],
      julho: [BRASA(1100, "31/07/2026"), AMERICAN(300, "26/07/2026")],
      agosto: [
        BRASA(1200, "08/09/2026"),
        ALFREDO(600, "01/09/2026"),
        AMERICAN(350, "26/08/2026"),
        ANGELICO(800, "26/08/2026"),
      ],
    }),
  );

function abrirAba(nome: RegExp) {
  const trigger = screen.getByRole("tab", { name: nome });
  fireEvent.mouseDown(trigger);
  fireEvent.click(trigger);
}

describe("cabeçalho do cliente com data e contato", () => {
  it("separa nome, data da última compra e telefones em campos próprios", () => {
    const c = lerCabecalhoCliente([
      "Cliente: 1045 - BRASA - Data da última compra: 08/09/2026 - Fone: (11) 3333-3333 - Celular: (11) 99999-9999",
    ])!;

    expect(c.codigo).toBe("1045");
    expect(c.nome).toBe("BRASA");
    expect(c.dataUltimaCompra).toBe("08/09/2026");
    expect(c.fone).toBe("(11) 3333-3333");
    expect(c.celular).toBe("(11) 99999-9999");
  });

  it("aceita a grafia sem acento e mantém rótulos desconhecidos", () => {
    const c = lerCabecalhoCliente([
      "Cliente: 90 - ANGELICO HORTIFRUTI - Data da ultima compra: 26/08/2026 - Vendedor: MARIA",
    ])!;

    expect(c.nome).toBe("ANGELICO HORTIFRUTI");
    expect(c.dataUltimaCompra).toBe("26/08/2026");
    expect(c.campos.vendedor).toBe("MARIA");
  });

  it("não confunde traço do nome com início dos campos", () => {
    expect(separarCamposCliente("SUPER - MERCADO ZE - Fone: 3232-1010")).toEqual({
      nome: "SUPER - MERCADO ZE",
      campos: { fone: "3232-1010" },
    });
  });
});

describe("identidade do cliente entre os meses", () => {
  it("o mesmo cliente não vira vários porque a data da última compra mudou", () => {
    const r = analyzeWorkbook(ARQUIVO());

    // BRASA aparece nos três meses com datas diferentes: é um cliente só
    const brasa = r.classificacaoCarteira.filter((c) => c.nome === "BRASA");
    expect(brasa).toHaveLength(1);
    expect(brasa[0].totalPeriodosPresente).toBe(3);
    expect(r.evolucaoConsolidada.clientesRecorrentes).toBe(1);
    expect(r.evolucaoConsolidada.clientesDistintos).toBe(5);
  });

  it("guarda o nome limpo, sem data nem telefone grudados", () => {
    const r = analyzeWorkbook(ARQUIVO());
    for (const c of r.classificacaoCarteira) {
      expect(c.nome).not.toMatch(/Data da última compra|Fone|Celular/);
    }
    expect(r.classificacaoCarteira.map((c) => c.nome).sort()).toEqual([
      "ALFREDO HAMBÚRGUER ARTESANAL",
      "AMERICAN EMBER",
      "ANGELICO HORTIFRUTI",
      "ARAMADO REST",
      "BRASA",
    ]);
  });

  it("mantém a correspondência entre cliente, data, receita e classificação", () => {
    const [, t2] = analyzeWorkbook(ARQUIVO()).transicoes;

    // julho → agosto: ANGELICO é novo, ALFREDO retornou, ninguém saiu
    expect(t2.novos.clientes).toEqual([
      expect.objectContaining({
        nome: "ANGELICO HORTIFRUTI",
        codigoCliente: "1048",
        dataUltimaCompra: "26/08/2026",
        faturamento: 800,
        celular: "(11) 98888-7777",
      }),
    ]);
    expect(t2.retornados.clientes).toEqual([
      expect.objectContaining({
        nome: "ALFREDO HAMBÚRGUER ARTESANAL",
        dataUltimaCompra: "01/09/2026",
        faturamento: 600,
      }),
    ]);
    expect(t2.novos.faturamento).toBe(800);
    expect(t2.retornados.faturamento).toBe(600);
  });

  it("não duplica nem omite clientes nas listas de cada mês", () => {
    const r = analyzeWorkbook(ARQUIVO());

    for (const t of r.transicoes) {
      const todos = [
        ...t.novos.clientes,
        ...t.retornados.clientes,
        ...t.mantidos.clientes,
      ].map((c) => c.chave);
      // cada cliente do mês aparece em exatamente um grupo
      expect(new Set(todos).size).toBe(todos.length);
      expect(todos.length).toBe(t.clientesPara);
      // e a receita dos grupos soma o faturamento do mês
      expect(t.novos.faturamento + t.retornados.faturamento + t.mantidos.faturamento)
        .toBeCloseTo(t.faturamentoPara, 4);
    }
  });
});

describe("aba Clientes Novos", () => {
  const renderAnalise = () => {
    const result = analyzeWorkbook(ARQUIVO());
    return { result, ...render(<AnalysisResults result={result} />) };
  };

  it("apresenta um cliente por linha, com a data em coluna própria", () => {
    renderAnalise();
    abrirAba(/Clientes Novos/);

    const angelico = screen.getByRole("row", { name: /ANGELICO HORTIFRUTI/ });
    const celulas = within(angelico).getAllByRole("cell").map((c) => c.textContent);

    // Cliente | Data da última compra | Código | Receita | Classificação | Fone | Celular
    expect(celulas[0]).toBe("ANGELICO HORTIFRUTI");
    expect(celulas[1]).toBe("26/08/2026");
    expect(celulas[2]).toBe("1048");
    expect(celulas[3]).toContain("800,00");
    expect(celulas[4]).toBe("Novo");
  });

  it("não junta os clientes em um texto corrido", () => {
    renderAnalise();
    abrirAba(/Clientes Novos/);

    // nenhuma célula concentra mais de um cliente
    for (const celula of screen.getAllByRole("cell")) {
      const texto = celula.textContent ?? "";
      const clientes = ["BRASA", "ALFREDO", "AMERICAN", "ANGELICO", "ARAMADO"].filter((n) =>
        texto.includes(n),
      );
      expect(clientes.length).toBeLessThanOrEqual(1);
      expect(texto).not.toMatch(/Data da última compra:|Fone:|Celular:/);
    }
  });

  it("separa quem entrou em cada passagem de mês, com os totais do mês", () => {
    renderAnalise();
    abrirAba(/Clientes Novos/);

    // junho → julho: AMERICAN entrou como novo
    const american = screen.getByRole("row", { name: /AMERICAN EMBER/ });
    expect(within(american).getByText("26/07/2026")).toBeInTheDocument();
    expect(within(american).getByText("Novo")).toBeInTheDocument();

    // julho → agosto: ALFREDO volta como retorno, não como novo
    const alfredo = screen.getByRole("row", { name: /ALFREDO HAMBÚRGUER/ });
    expect(within(alfredo).getByText("Retorno")).toBeInTheDocument();
    expect(within(alfredo).getByText("01/09/2026")).toBeInTheDocument();

    // os totais de cada passagem continuam visíveis
    expect(screen.getAllByText("Receita dos novos").length).toBe(2);
    expect(screen.getAllByText("Receita dos retornos").length).toBe(2);
  });

  it("mostra fone e celular em colunas próprias quando o relatório traz", () => {
    renderAnalise();
    abrirAba(/Clientes Novos/);

    const angelico = screen.getByRole("row", { name: /ANGELICO HORTIFRUTI/ });
    expect(within(angelico).getByText("(11) 98888-7777")).toBeInTheDocument();
  });

  it("lista quem saiu, com data e receita, na aba Clientes Perdidos", () => {
    renderAnalise();
    abrirAba(/Clientes Perdidos/);

    // ARAMADO aparece na passagem em que saiu e, depois, na lista da janela:
    // aqui interessa a linha da passagem junho → julho.
    const [passagem] = screen.getAllByRole("table");
    const aramado = within(passagem).getByRole("row", { name: /ARAMADO REST/ });
    const celulas = within(aramado).getAllByRole("cell").map((c) => c.textContent);

    expect(celulas[0]).toBe("ARAMADO REST");
    expect(celulas[1]).toBe("03/06/2026");
    expect(celulas[3]).toContain("200,00");
  });

  it("não junta em texto corrido quem saiu", () => {
    renderAnalise();
    abrirAba(/Clientes Perdidos/);

    for (const celula of screen.getAllByRole("cell")) {
      expect(celula.textContent ?? "").not.toMatch(/Data da última compra:|Fone:|Celular:/);
    }
  });
});

describe("relatório sem código de cliente", () => {
  it("segue reconhecendo o mesmo cliente pelo nome ao longo dos meses", () => {
    // Alguns exports não imprimem o código: a identidade cai para o nome
    // normalizado, e a data da última compra continua fora dele.
    const semCodigo = (data: string, valor: number): AbaFixture => ({
      clientes: [
        {
          codigo: "",
          nome: "BAR DO MARQUINHOS",
          dataUltimaCompra: data,
          itens: [{ codigo: "5001", produto: "PRODUTO A", qtd: 1, preco: valor }],
        },
      ],
    });

    const r = analyzeWorkbook(
      montarArquivo({
        junho: semCodigo("29/06/2026", 100),
        julho: semCodigo("30/07/2026", 150),
        agosto: semCodigo("27/08/2026", 200),
      }),
    );

    expect(r.classificacaoCarteira).toHaveLength(1);
    const [c] = r.classificacaoCarteira;
    expect(c.nome).toBe("BAR DO MARQUINHOS");
    expect(c.totalPeriodosPresente).toBe(3);
    expect(c.dataUltimaCompra).toBe("27/08/2026");
    expect(r.transicoes.every((t) => t.mantidos.total === 1)).toBe(true);
  });
});
