import { describe, it, expect } from "vitest";
import { analyzeWorkbook, type PeriodData } from "../lib/excel-analyzer";
import * as XLSX from "xlsx";
import { abaRelatorio, montarArquivo, type AbaFixture, type Rows } from "./relatorio-fixture";

/**
 * Clients are keyed by identity (the report's client code), so the tests look
 * them up by the name a person would read on screen.
 */
function porNome(periodo: PeriodData, nome: string) {
  return [...periodo.clientes.values()].find((c) => c.nome === nome);
}

/** A file whose first tab is written by hand, followed by regular report tabs. */
function comAbaCrua(nome: string, rows: Rows, resto: Record<string, AbaFixture>): ArrayBuffer {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), nome);
  for (const [n, aba] of Object.entries(resto)) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(abaRelatorio(aba)), n);
  }
  const out = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  return out instanceof ArrayBuffer ? out : new Uint8Array(out).buffer;
}

/**
 * The scenario the reader used to get wrong: a three-month file (junho, julho,
 * agosto) where junho was loaded but never reached the numbers.
 *
 * Every assertion here fails if junho is dropped — the point is not that the
 * file opened, it is that junho's rows are inside the KPIs, the month-by-month
 * evolution, the product comparison, the portfolio and the validation.
 */

//  junho                                   julho                  agosto
//  ALFA  PARAFUSO 10 × 2,50 = 25           ALFA PARAFUSO 30       ALFA  PARAFUSO 50
//        PORCA    20 × 1,00 = 25 (errado)  GAMA PORCA    10       BETA  ARRUELA  12
//  BETA  ARRUELA   5 × 3,00 = 15                                  DELTA PARAFUSO  5
//        GRAXA     3 × 10,0 = 30
const junho: AbaFixture = {
  intervalo: "01/06/2025 a 30/06/2025",
  clientes: [
    {
      codigo: "1001",
      nome: "CLIENTE ALFA",
      itens: [
        { codigo: "5001", produto: "PARAFUSO", qtd: 10, preco: 2.5 },
        // 20 × 1,00 = 20, but the sheet reports 25 — an inconsistency in junho
        { codigo: "5002", produto: "PORCA", qtd: 20, preco: 1, total: 25 },
      ],
    },
    {
      codigo: "1002",
      nome: "CLIENTE BETA",
      itens: [
        { codigo: "5003", produto: "ARRUELA", qtd: 5, preco: 3 },
        // sold in junho and never again
        { codigo: "5004", produto: "GRAXA", qtd: 3, preco: 10 },
      ],
    },
  ],
};

const julho: AbaFixture = {
  intervalo: "01/07/2025 a 31/07/2025",
  clientes: [
    { codigo: "1001", nome: "CLIENTE ALFA", itens: [{ codigo: "5001", produto: "PARAFUSO", qtd: 12, preco: 2.5 }] },
    { codigo: "1003", nome: "CLIENTE GAMA", itens: [{ codigo: "5002", produto: "PORCA", qtd: 10, preco: 1 }] },
  ],
};

const agosto: AbaFixture = {
  intervalo: "01/08/2025 a 31/08/2025",
  clientes: [
    { codigo: "1001", nome: "CLIENTE ALFA", itens: [{ codigo: "5001", produto: "PARAFUSO", qtd: 20, preco: 2.5 }] },
    { codigo: "1002", nome: "CLIENTE BETA", itens: [{ codigo: "5003", produto: "ARRUELA", qtd: 4, preco: 3 }] },
    { codigo: "1004", nome: "CLIENTE DELTA", itens: [{ codigo: "5001", produto: "PARAFUSO", qtd: 2, preco: 2.5 }] },
  ],
};

const TOTAL_JUNHO = 25 + 25 + 15 + 30; // 95
const TOTAL_JULHO = 30 + 10; // 40
const TOTAL_AGOSTO = 50 + 12 + 5; // 67

const arquivo = () => montarArquivo({ junho, julho, agosto });

describe("cenário real · junho + julho + agosto", () => {
  it("lê as três abas, na ordem cronológica, com os dados de cada uma", () => {
    const r = analyzeWorkbook(arquivo());

    expect(r.periodos.map((p) => p.nome)).toEqual(["junho", "julho", "agosto"]);
    expect(r.leitura.abasLidas).toEqual(["junho", "julho", "agosto"]);
    expect(r.leitura.abasIgnoradas).toEqual([]);

    // each tab was really processed, not merely opened
    expect(r.periodos.map((p) => p.clientes.size)).toEqual([2, 2, 3]);
    expect(r.periodos.map((p) => p.produtosLidos)).toEqual([4, 2, 3]);
    expect(r.periodos.map((p) => p.totalFaturamento)).toEqual([
      TOTAL_JUNHO,
      TOTAL_JULHO,
      TOTAL_AGOSTO,
    ]);

    // junho's rows, down to the individual product
    const alfaJunho = porNome(r.periodos[0], "CLIENTE ALFA")!;
    expect(alfaJunho.codigoCliente).toBe("1001");
    expect(alfaJunho.produtos.map((p) => p.produto)).toEqual(["PARAFUSO", "PORCA"]);
    expect(alfaJunho.produtos[0].quantidade).toBe(10);
    expect(alfaJunho.produtos[0].valorTotal).toBe(25);
  });

  it("usa junho no comparativo mês a mês", () => {
    const { evolucaoPeriodos: ev } = analyzeWorkbook(arquivo());

    expect(ev.map((p) => p.nome)).toEqual(["junho", "julho", "agosto"]);
    expect(ev[0].faturamento).toBe(TOTAL_JUNHO);
    expect(ev[0].clientes).toBe(2);
    expect(ev[0].ticketMedio).toBe(TOTAL_JUNHO / 2);
    expect(ev[0].variacaoFaturamento).toBeNull(); // first month of the window

    // julho is measured against junho — not against itself
    expect(ev[1].variacaoFaturamento).toBeCloseTo(((TOTAL_JULHO - TOTAL_JUNHO) / TOTAL_JUNHO) * 100, 4);
    expect(ev[1].entraram).toBe(1); // GAMA
    expect(ev[1].sairam).toBe(1); // BETA
    expect(ev[2].variacaoFaturamento).toBeCloseTo(((TOTAL_AGOSTO - TOTAL_JULHO) / TOTAL_JULHO) * 100, 4);
  });

  it("usa junho no comparativo de produtos e nos gráficos", () => {
    const r = analyzeWorkbook(arquivo());
    const porNome = Object.fromEntries(r.comparativoProdutos.map((p) => [p.produto, p]));

    // every product of every month is in the comparison
    expect(Object.keys(porNome).sort()).toEqual(["ARRUELA", "GRAXA", "PARAFUSO", "PORCA"]);

    // GRAXA only ever sold in junho — it used to vanish from the analysis
    expect(porNome.GRAXA.valorPorPeriodo).toEqual({ junho: 30, julho: 0, agosto: 0 });
    expect(porNome.GRAXA.qtdPorPeriodo.junho).toBe(3);

    // the series that feed the charts carry the three months for every product
    for (const p of r.comparativoProdutos) {
      expect(Object.keys(p.valorPorPeriodo)).toEqual(["junho", "julho", "agosto"]);
      expect(Object.keys(p.qtdPorPeriodo)).toEqual(["junho", "julho", "agosto"]);
    }

    expect(porNome.PARAFUSO.valorPorPeriodo).toEqual({ junho: 25, julho: 30, agosto: 55 });
    expect(porNome.PORCA.valorPorPeriodo).toEqual({ junho: 25, julho: 10, agosto: 0 });
    expect(porNome.ARRUELA.valorPorPeriodo).toEqual({ junho: 15, julho: 0, agosto: 12 });
  });

  it("usa junho para separar cliente novo de cliente que voltou", () => {
    const r = analyzeWorkbook(arquivo());

    // month over month, BETA and DELTA both look new in agosto
    expect(r.clientesNovos.map((c) => c.nome).sort()).toEqual(["CLIENTE BETA", "CLIENTE DELTA"]);
    // reading junho shows BETA had already bought — only DELTA is really new
    expect(r.carteira.novos.map((c) => c.nome)).toEqual(["CLIENTE DELTA"]);
    expect(r.carteira.retornados.map((c) => c.nome)).toEqual(["CLIENTE BETA"]);
    expect(r.carteira.faturamentoNovos).toBe(5);

    // GAMA bought in julho only: one month without buying
    const gama = r.classificacaoCarteira.find((c) => c.nome === "CLIENTE GAMA")!;
    expect(gama.status).toBe("frequencia_reduzida");
    expect(gama.ultimoPeriodo).toBe("julho");

    // the portfolio spans the three months
    expect(r.carteira.periodos).toEqual(["junho", "julho", "agosto"]);
    expect(r.classificacaoCarteira.find((c) => c.nome === "CLIENTE ALFA")!.periodosPresente).toEqual([
      "junho",
      "julho",
      "agosto",
    ]);
    expect(r.classificacaoCarteira.find((c) => c.nome === "CLIENTE BETA")!.periodosPresente).toEqual([
      "junho",
      "agosto",
    ]);
  });

  it("valida junho junto com os demais meses", () => {
    const r = analyzeWorkbook(arquivo());

    // the wrong PORCA line lives in junho — the alert must name junho
    const alertasJunho = r.inconsistencias.filter((i) => i.startsWith("junho →"));
    expect(alertasJunho).toHaveLength(1);
    expect(alertasJunho[0]).toContain("CLIENTE ALFA");
    expect(alertasJunho[0]).toContain("PORCA");
    // traceable back to the exact cell of the junho tab
    expect(alertasJunho[0]).toMatch(/junho!F\d+/);

    // the validation report covers the three tabs, not just the two compared,
    // plus the conferência planilha × sistema
    const checksDosPeriodos = r.periodos.reduce((s, p) => s + p.validationReport.totalChecks, 0);
    expect(r.validationReport.totalChecks).toBe(
      checksDosPeriodos + r.conferencia.validationReport.totalChecks,
    );
    expect(r.validationReport.diagnostics.some((d) => d.cellRef.startsWith("junho!"))).toBe(true);
  });

  it("mantém as três abas em .xls legado, fora de ordem e com aba oculta", () => {
    const r = analyzeWorkbook(
      montarArquivo({ agosto, junho, julho }, "biff8", ["junho"]),
    );

    expect(r.periodos.map((p) => p.nome)).toEqual(["junho", "julho", "agosto"]);
    expect(r.periodos[0].oculta).toBe(true);
    expect(r.periodos[0].totalFaturamento).toBe(TOTAL_JUNHO);
  });

  it("lê junho mesmo quando a aba dele veio de um modelo diferente", () => {
    // The junho tab of a real file came from another version of the export:
    // indented by one column, uppercase client label and renamed columns.
    const junhoOutroModelo: AbaFixture = {
      ...junho,
      recuo: 1,
      cabecalhoDeProdutos: ["Cod.", "Produto", "Unid", "Quantidade", "Preço médio", "Valor total"],
    };

    const r = analyzeWorkbook(montarArquivo({ junho: junhoOutroModelo, julho, agosto }));

    expect(r.periodos.map((p) => p.nome)).toEqual(["junho", "julho", "agosto"]);
    expect(r.periodos[0].totalFaturamento).toBe(TOTAL_JUNHO);
    expect(porNome(r.periodos[0], "CLIENTE BETA")!.produtos.map((p) => p.produto)).toEqual([
      "ARRUELA",
      "GRAXA",
    ]);
  });

  it("identifica os três períodos pelo conteúdo quando as abas não têm nome de mês", () => {
    const r = analyzeWorkbook(montarArquivo({ Plan1: junho, Plan2: julho, Plan3: agosto }));

    expect(r.periodos.map((p) => p.nome)).toEqual(["junho/2025", "julho/2025", "agosto/2025"]);
    expect(r.periodos.map((p) => p.sheetName)).toEqual(["Plan1", "Plan2", "Plan3"]);
    expect(r.periodos.map((p) => p.origemPeriodo)).toEqual(["conteudo", "conteudo", "conteudo"]);
    expect(r.periodos[0].totalFaturamento).toBe(TOTAL_JUNHO);
  });

  it("um quarto mês entra sozinho, sem derrubar os anteriores", () => {
    const setembro: AbaFixture = {
      intervalo: "01/09/2025 a 30/09/2025",
      clientes: [
        { codigo: "1001", nome: "CLIENTE ALFA", itens: [{ codigo: "5001", produto: "PARAFUSO", qtd: 1, preco: 2.5 }] },
      ],
    };

    const r = analyzeWorkbook(montarArquivo({ junho, julho, agosto, setembro }));

    expect(r.periodos.map((p) => p.nome)).toEqual(["junho", "julho", "agosto", "setembro"]);
    expect(r.periodos[0].totalFaturamento).toBe(TOTAL_JUNHO);
    expect(r.periodoAtual.nome).toBe("setembro");
    expect(r.evolucaoPeriodos.map((p) => p.nome)).toEqual([
      "junho",
      "julho",
      "agosto",
      "setembro",
    ]);
    expect(r.comparativoProdutos[0].valorPorPeriodo).toHaveProperty("junho");
  });

  it("não perde produtos depois do rodapé de página de um relatório paginado", () => {
    // Paginated exports print "Gerado em ... / Página 1 de 2" in the middle of a
    // client and go on with that client's products on the next page.
    const r = analyzeWorkbook(
      comAbaCrua(
        "junho",
        [
          ["Vendas de produto no período"],
          ["Cliente: 1001 - CLIENTE ALFA - Data: 01/06/2025 a 30/06/2025"],
          ["Código", "Descrição", "UN", "Qtd. venda", "Preço médio venda", "Venda bruta"],
          ["5001", "PARAFUSO", "UN", 10, 2.5, 25],
          ["Gerado em 05/09/2025 14:32"],
          ["Página 1 de 2"],
          ["Código", "Descrição", "UN", "Qtd. venda", "Preço médio venda", "Venda bruta"],
          ["5002", "PORCA", "UN", 20, 1, 20],
          ["Totais por unidade:"],
          ["UN", null, null, 30, null, 45],
        ],
        { julho, agosto },
      ),
    );

    const alfa = porNome(r.periodos[0], "CLIENTE ALFA")!;
    expect(alfa.produtos.map((p) => p.produto)).toEqual(["PARAFUSO", "PORCA"]);
    expect(r.periodos[0].totalFaturamento).toBe(45);
    // the per-unit totals row is not read as a nameless product
    expect(r.periodos[0].produtosLidos).toBe(2);
  });

  it("soma os dois blocos de um cliente listado duas vezes na mesma aba", () => {
    const r = analyzeWorkbook(
      comAbaCrua(
        "junho",
        [
          ["Cliente: 1001 - CLIENTE ALFA"],
          ["Código", "Descrição", "UN", "Qtd. venda", "Preço médio venda", "Venda bruta"],
          ["5001", "PARAFUSO", "UN", 10, 2.5, 25],
          [null, null, null, 10, null, 25],
          [],
          ["Cliente: 1001 - CLIENTE ALFA"],
          ["Código", "Descrição", "UN", "Qtd. venda", "Preço médio venda", "Venda bruta"],
          ["5002", "PORCA", "UN", 20, 1, 20],
          [null, null, null, 20, null, 20],
        ],
        { julho, agosto },
      ),
    );

    const alfa = porNome(r.periodos[0], "CLIENTE ALFA")!;
    // neither block is dropped, and the reported total is the sum of both
    expect(alfa.produtos.map((p) => p.produto)).toEqual(["PARAFUSO", "PORCA"]);
    expect(alfa.totalGeral).toBe(45);
    expect(alfa.totalConsistente).toBe(true);
  });

  it("avisa quando uma aba de relatório não produziu nenhum cliente", () => {
    // A tab whose client lines the reader cannot recognise must not slip through
    // as a zeroed month: it becomes an error in the validation report.
    const r = analyzeWorkbook(
      comAbaCrua("junho", [
        ["Vendas de produto no período"],
        ["Razao social 1001 INDUSTRIA X"],
        ["Código", "Descrição", "UN", "Qtd. venda", "Preço médio venda", "Venda bruta"],
        ["5001", "PARAFUSO", "UN", 10, 2.5, 25],
      ], { julho, agosto }),
    );

    const erro = r.validationReport.diagnostics.find(
      (d) => d.field === "aba junho" && d.severity === "error",
    );
    expect(erro).toBeDefined();
    expect(erro!.impact).toContain("nenhum cliente");
  });
});
