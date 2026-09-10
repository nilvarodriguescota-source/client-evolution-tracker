import { describe, it, expect } from "vitest";
import * as XLSX from "xlsx";
import { analyzeWorkbook } from "../lib/excel-analyzer";

/**
 * Builds a workbook in the "Vendas de produto no período" report layout,
 * one sheet per period.
 *
 * @param periods sheet name -> client names present in that period
 */
function buildWorkbook(
  periods: Record<string, string[]>,
  bookType: XLSX.BookType = "xlsx",
): ArrayBuffer {
  const wb = XLSX.utils.book_new();

  for (const [sheetName, clientes] of Object.entries(periods)) {
    const rows: (string | number | null)[][] = [];
    clientes.forEach((nome, i) => {
      rows.push([`Cliente: ${100 + i} - ${nome}`, null, null, null, null, null]);
      rows.push(["Código", "Descrição", "Unidade", "Qtd", "Preço Médio", "Venda Bruta"]);
      // qty 2 x price 10 = 20
      rows.push(["5001", "PRODUTO A", "UN", 2, 10, 20]);
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

  it("considera até 6 períodos, mantendo os 2 últimos para a comparação", () => {
    const nomes = ["m1", "m2", "m3", "m4", "m5", "m6", "m7", "m8"];
    const periods = Object.fromEntries(nomes.map((n) => [n, ["CLIENTE A"]]));
    const r = analyzeWorkbook(buildWorkbook(periods));

    // only the 6 most recent sheets feed the portfolio window
    expect(r.periodos.map((p) => p.nome)).toEqual(["m3", "m4", "m5", "m6", "m7", "m8"]);
    // the 2-period comparison is untouched
    expect(r.periodoAnterior.nome).toBe("m7");
    expect(r.periodoAtual.nome).toBe("m8");
  });

  it("usa todas as abas quando há menos de 6", () => {
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
    // The 2-period view calls it "novo"; the 6-month view must call it active
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
