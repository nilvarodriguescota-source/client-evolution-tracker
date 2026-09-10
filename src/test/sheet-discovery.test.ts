import { describe, it, expect } from "vitest";
import * as XLSX from "xlsx";
import {
  descobrirAbas,
  lerCabecalhoCliente,
  lerCabecalhoProdutos,
  lerPeriodoDoConteudo,
  lerPeriodoDoNome,
  lerPastaDeTrabalho,
  ordenarPorPeriodo,
  COLUNA_AUSENTE,
} from "../lib/sheet-discovery";
import { montarArquivo, type AbaFixture } from "./relatorio-fixture";

const mes: AbaFixture = {
  clientes: [
    { codigo: "1", nome: "CLIENTE A", itens: [{ codigo: "5001", produto: "PARAFUSO", qtd: 2, preco: 10 }] },
  ],
};

const abrir = (buffer: ArrayBuffer) => descobrirAbas(lerPastaDeTrabalho(buffer));

describe("lerCabecalhoCliente", () => {
  it("aceita as variações de escrita do cabeçalho", () => {
    const casos: [string, string, string][] = [
      ["Cliente: 1001 - INDUSTRIA X", "1001", "INDUSTRIA X"],
      ["CLIENTE: 1001 - INDUSTRIA X", "1001", "INDUSTRIA X"],
      ["cliente:1001-INDUSTRIA X", "1001", "INDUSTRIA X"],
      ["Cliente: 1001 – INDUSTRIA X", "1001", "INDUSTRIA X"],
      ["Cliente: AB-10 - INDUSTRIA X", "AB-10", "INDUSTRIA X"],
      ["Cliente: INDUSTRIA X", "", "INDUSTRIA X"],
      ["Cliente: SUPER-MERCADO ZE", "", "SUPER-MERCADO ZE"],
      ["Cliente: 1001 - INDUSTRIA X - Data: 01/06/2025 a 30/06/2025", "1001", "INDUSTRIA X"],
    ];

    for (const [texto, codigo, nome] of casos) {
      expect(lerCabecalhoCliente([texto]), texto).toEqual({ codigo, nome, coluna: 0 });
    }
  });

  it("encontra o cabeçalho em qualquer coluna", () => {
    expect(lerCabecalhoCliente([null, null, "Cliente: 7 - BETA"])).toEqual({
      codigo: "7",
      nome: "BETA",
      coluna: 2,
    });
  });

  it("não confunde outras linhas com cabeçalho de cliente", () => {
    expect(lerCabecalhoCliente(["Clientes atendidos no mês"])).toBeNull();
    expect(lerCabecalhoCliente(["5001", "PARAFUSO", "UN", 2, 10, 20])).toBeNull();
  });
});

describe("lerCabecalhoProdutos", () => {
  it("mapeia as colunas pelo rótulo, em qualquer ordem", () => {
    expect(
      lerCabecalhoProdutos(["Descrição", "Código", "Qtd. venda", "Venda bruta", "Preço médio venda", "UN"]),
    ).toEqual({ descricao: 0, codigo: 1, quantidade: 2, vendaBruta: 3, precoMedio: 4, unidade: 5 });
  });

  it("aceita rótulos alternativos e colunas deslocadas", () => {
    expect(
      lerCabecalhoProdutos([null, "Cod.", "Produto", "Unid", "Quantidade", "Preço médio", "Valor total"]),
    ).toEqual({ codigo: 1, descricao: 2, unidade: 3, quantidade: 4, precoMedio: 5, vendaBruta: 6 });
  });

  it("marca como ausente a coluna que o relatório não traz", () => {
    const mapa = lerCabecalhoProdutos(["Código", "Descrição", "Qtd. venda", "Venda bruta"])!;
    expect(mapa.precoMedio).toBe(COLUNA_AUSENTE);
    expect(mapa.unidade).toBe(COLUNA_AUSENTE);
  });

  it("não trata linha de produto como cabeçalho", () => {
    expect(lerCabecalhoProdutos(["5001", "PARAFUSO", "UN", 2, 10, 20])).toBeNull();
  });
});

describe("lerPeriodoDoNome", () => {
  it("lê mês e ano de vários formatos de nome de aba", () => {
    expect(lerPeriodoDoNome("junho")).toEqual({ mes: 6, ano: undefined });
    expect(lerPeriodoDoNome("Junho 2025")).toEqual({ mes: 6, ano: 2025 });
    expect(lerPeriodoDoNome("JUN-25")).toEqual({ mes: 6, ano: 2025 });
    expect(lerPeriodoDoNome("06/2025")).toEqual({ mes: 6, ano: 2025 });
    expect(lerPeriodoDoNome("2025-06")).toEqual({ mes: 6, ano: 2025 });
    expect(lerPeriodoDoNome("MARÇO")).toEqual({ mes: 3, ano: undefined });
    expect(lerPeriodoDoNome("Plan1").mes).toBeUndefined();
  });
});

describe("lerPeriodoDoConteudo", () => {
  it("lê o período pelas datas impressas no relatório", () => {
    expect(
      lerPeriodoDoConteudo([
        ["Vendas de produto no período"],
        ["Cliente: 1 - A - Data: 01/06/2025 a 30/06/2025"],
        ["Gerado em 05/09/2025 14:32"],
      ]),
    ).toEqual({ mes: 6, ano: 2025 });
  });

  it("ignora a data de emissão do relatório", () => {
    expect(lerPeriodoDoConteudo([["Gerado em 05/09/2025"]])).toEqual({});
  });
});

describe("ordenarPorPeriodo", () => {
  it("mantém no lugar o item sem mês e ordena os demais", () => {
    const itens = [
      { index: 0, mes: 8 },
      { index: 1 },
      { index: 2, mes: 6 },
    ];
    expect(ordenarPorPeriodo(itens).map((i) => i.mes)).toEqual([6, undefined, 8]);
  });
});

describe("descobrirAbas", () => {
  it("lê todas as abas de relatório, sem teto", () => {
    const nomes = ["janeiro", "fevereiro", "marco", "abril", "maio", "junho", "julho", "agosto"];
    const d = abrir(montarArquivo(Object.fromEntries(nomes.map((n) => [n, mes]))));

    expect(d.abas.map((a) => a.sheetName)).toEqual(nomes);
    expect(d.ignoradas).toEqual([]);
  });

  it("descarta apenas o que não é relatório, dizendo o motivo", () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["Relatório de vendas"], ["Empresa X"]]), "Capa");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([[]]), "Vazia");
    const comCapa = XLSX.write(wb, { type: "array", bookType: "xlsx" });

    const base = lerPastaDeTrabalho(montarArquivo({ junho: mes, julho: mes, agosto: mes }));
    const extra = XLSX.read(new Uint8Array(comCapa), { type: "array" });
    for (const nome of extra.SheetNames) {
      base.SheetNames.unshift(nome);
      base.Sheets[nome] = extra.Sheets[nome];
    }

    const d = descobrirAbas(base);
    expect(d.abas.map((a) => a.sheetName)).toEqual(["junho", "julho", "agosto"]);
    expect(d.ignoradas.map((i) => i.sheetName).sort()).toEqual(["Capa", "Vazia"]);
    expect(d.ignoradas.find((i) => i.sheetName === "Vazia")!.motivo).toContain("vazia");
    expect(d.ignoradas.find((i) => i.sheetName === "Capa")!.motivo).toContain("relatório");
    expect(d.todasAsAbas).toHaveLength(5);
  });

  it("põe as abas em ordem cronológica e não pela posição no arquivo", () => {
    const d = abrir(montarArquivo({ agosto: mes, junho: mes, julho: mes }));
    expect(d.abas.map((a) => a.sheetName)).toEqual(["junho", "julho", "agosto"]);
    expect(d.abas.map((a) => a.origemPeriodo)).toEqual([
      "nome-da-aba",
      "nome-da-aba",
      "nome-da-aba",
    ]);
  });

  it("lê abas ocultas como qualquer outra", () => {
    const d = abrir(montarArquivo({ junho: mes, julho: mes, agosto: mes }, "xlsx", ["junho"]));
    expect(d.abas.map((a) => a.sheetName)).toEqual(["junho", "julho", "agosto"]);
    expect(d.abas[0].oculta).toBe(true);
  });

  it("detecta o layout de cada aba separadamente", () => {
    const d = abrir(
      montarArquivo({
        junho: { ...mes, recuo: 2 },
        julho: mes,
      }),
    );

    expect(d.abas[0].colunas.codigo).toBe(2);
    expect(d.abas[1].colunas.codigo).toBe(0);
  });
});
