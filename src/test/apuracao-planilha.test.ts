import { describe, it, expect } from "vitest";
import * as XLSX from "xlsx";
import { analyzeWorkbook, type AnalysisResult } from "../lib/excel-analyzer";
import { lerNumero } from "../lib/numero-br";

/**
 * Apuração do arquivo de vendas, de ponta a ponta, com SEIS abas mensais.
 *
 * O valor esperado de cada mês é calculado aqui, em centavos inteiros, direto
 * dos registros que vão para a planilha — nunca pelo mesmo código que o sistema
 * usa. O mesmo conteúdo é gravado em todos os formatos que o ERP produz
 * (.xlsx numérico, .xlsx com texto em formato brasileiro, .xls binário, .xls em
 * HTML, .xls em XML 2003) e o sistema precisa chegar exatamente aos mesmos
 * números em todos eles.
 */

// ---------------------------------------------------------------------------
// Base de dados do teste
// ---------------------------------------------------------------------------

interface Venda {
  codigo: string;
  produto: string;
  un: "UN" | "KG";
  qtd: number;
  /** preço médio como impresso (2 casas) */
  preco: number;
  /** venda bruta impressa, em centavos */
  centavos: number;
}

interface Comprador {
  codigo: string;
  nome: string;
  vendas: Venda[];
}

const MESES = ["janeiro", "fevereiro", "marco", "abril", "maio", "junho"];
const ROTULOS = ["janeiro 2026", "fevereiro 2026", "março 2026", "abril 2026", "maio 2026", "junho 2026"];

const v = (codigo: string, produto: string, qtd: number, preco: number, centavos: number, un: "UN" | "KG" = "UN"): Venda => ({
  codigo,
  produto,
  un,
  qtd,
  preco,
  centavos,
});

/**
 * ALFA       compra nos 6 meses (valores com milhar e centavos)
 * BETA       compra em jan–mar e some (perdido)
 * GAMA       começa em abril (novo)
 * DELTA      jan, fev, some em mar/abr, volta em mai/jun (retorno)
 * *MARIA     funcionária: jan–mar e para — nunca "perdida"
 * *JOÃO      funcionário: só junho — nunca "novo"
 * *CLIENTE TESTE  funcionário sem código: fev e mai
 */
function compradoresDoMes(m: number): Comprador[] {
  const lista: Comprador[] = [];
  const fator = m + 1; // valores mudam mês a mês

  lista.push({
    codigo: "1001",
    nome: "ALFA COMERCIO LTDA",
    vendas: [
      // 1.200 × 3,50 = 4.200,00 — o caso que o leitor antigo transformava em 4,2
      v("5001", "ARROZ TIPO 1 5KG", 1200, 3.5, 420000),
      // 7 × 1,43 = 10,01 — preço médio arredondado
      v("5002", "PARAFUSO 3/8", 7, 1.43, 1001),
      // 300 × 3,33 = 999,00 impresso como 1.000,00: diferença é só o arredondamento do preço médio
      v("5003", "OLEO DE SOJA 900ML", 300, 3.33, 100000),
      v("5010", "CARNE BOVINA", 12.5 * fator, 39.9, Math.round(12.5 * fator * 3990), "KG"),
    ],
  });
  if (m <= 2) {
    lista.push({
      codigo: "1002",
      nome: "BETA MERCEARIA",
      vendas: [v("5004", "FEIJAO CARIOCA 1KG", 50, 8.79, 43950), v("5001", "ARROZ TIPO 1 5KG", 10, 25.9, 25900)],
    });
  }
  if (m >= 3) {
    lista.push({
      codigo: "1003",
      nome: "GAMA SUPERMERCADOS",
      vendas: [v("5005", "ACUCAR CRISTAL 5KG", 1000, 12.34, 1234000)],
    });
  }
  if (m <= 1 || m >= 4) {
    lista.push({
      codigo: "1004",
      nome: "DELTA PADARIA",
      vendas: [v("5006", "FARINHA DE TRIGO 25KG", 3, 89.9, 26970)],
    });
  }
  if (m <= 2) {
    lista.push({
      codigo: "2001",
      nome: "*MARIA SILVA",
      vendas: [v("5007", "CAFE 500G", 2, 15.5, 3100)],
    });
  }
  if (m === 5) {
    lista.push({
      codigo: "2002",
      nome: "*JOÃO PEREIRA",
      vendas: [v("5007", "CAFE 500G", 4, 15.5, 6200), v("5008", "LEITE INTEGRAL 1L", 12, 4.99, 5988)],
    });
  }
  if (m === 1 || m === 4) {
    lista.push({
      codigo: "",
      nome: "*CLIENTE TESTE",
      vendas: [v("5009", "BOLACHA RECHEADA", 1, 3.75, 375)],
    });
  }
  return lista;
}

const esperadoCentavos = MESES.map((_, m) =>
  compradoresDoMes(m).reduce((s, c) => s + c.vendas.reduce((a, x) => a + x.centavos, 0), 0),
);
const esperadoFuncionariosCentavos = MESES.map((_, m) =>
  compradoresDoMes(m)
    .filter((c) => c.nome.startsWith("*"))
    .reduce((s, c) => s + c.vendas.reduce((a, x) => a + x.centavos, 0), 0),
);
const esperadoQuantidade = MESES.map(
  (_, m) => Math.round(compradoresDoMes(m).reduce((s, c) => s + c.vendas.reduce((a, x) => a + x.qtd, 0), 0) * 10000) / 10000,
);
const esperadoRegistros = MESES.map((_, m) =>
  compradoresDoMes(m).reduce((s, c) => s + c.vendas.length, 0),
);
const esperadoClientesCarteira = MESES.map(
  (_, m) => compradoresDoMes(m).filter((c) => !c.nome.startsWith("*")).length,
);

// ---------------------------------------------------------------------------
// Formatos de gravação
// ---------------------------------------------------------------------------

type Celula = string | number | null;

/** 4200 → "4.200,00" (como o ERP imprime em texto) */
function br(n: number, casas = 2): string {
  const [int, dec] = Math.abs(n).toFixed(casas).split(".");
  const comMilhar = int.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  return `${n < 0 ? "-" : ""}${comMilhar}${casas ? `,${dec}` : ""}`;
}
/** quantidade sem casas desnecessárias: 1200 → "1.200", 12,5 → "12,5" */
function brQtd(n: number): string {
  return Number.isInteger(n) ? br(n, 0) : br(n, 1);
}

type Formato = "numero" | "texto-br" | "moeda-br";

interface OpcoesAba {
  formato: Formato;
  /** colunas vazias antes do relatório */
  recuo?: number;
  /** rótulo alternativo da coluna de valor */
  rotuloValor?: string;
  /** imprime "Data: 01/mm/2026 a ..." em cada cliente (identificação por conteúdo) */
  intervalo?: boolean;
  /** sobrescreve o subtotal impresso do primeiro cliente (centavos) */
  subtotalErrado?: number;
  /** sobrescreve o total impresso da aba (centavos) */
  totalErrado?: number;
  /** linha com valor solta, fora de qualquer cliente */
  linhaSolta?: boolean;
}

function linhasDaAba(m: number, o: OpcoesAba): Celula[][] {
  const pad = (cells: Celula[]) => [...Array<null>(o.recuo ?? 0).fill(null), ...cells];
  const num = (n: number, casas = 2): Celula =>
    o.formato === "numero" ? n : o.formato === "moeda-br" && casas === 2 ? `R$ ${br(n, casas)}` : br(n, casas);
  const qtd = (n: number): Celula => (o.formato === "numero" ? n : brQtd(n));

  const mm = String(m + 1).padStart(2, "0");
  const rows: Celula[][] = [];
  rows.push(pad(["Vendas de produto no período"]));
  rows.push(pad([`Período: 01/${mm}/2026 a 28/${mm}/2026`]));
  rows.push([]);
  if (o.linhaSolta) rows.push(pad([null, null, null, null, null, num(123.45)]));

  const porUnidade = new Map<string, { qtd: number; centavos: number }>();
  compradoresDoMes(m).forEach((c, i) => {
    const id = c.codigo ? `${c.codigo} - ${c.nome}` : c.nome;
    const sufixo = o.intervalo ? ` - Data: 01/${mm}/2026 a 28/${mm}/2026` : "";
    rows.push(pad([`Cliente: ${id}${sufixo} - Data da última compra: 20/${mm}/2026`]));
    rows.push(pad(["Código", "Descrição", "UN", "Qtd. venda", "Preço médio venda", o.rotuloValor ?? "Venda bruta"]));
    let soma = 0;
    let q = 0;
    for (const x of c.vendas) {
      rows.push(pad([x.codigo, x.produto, x.un, qtd(x.qtd), num(x.preco), num(x.centavos / 100)]));
      soma += x.centavos;
      q += x.qtd;
      const u = porUnidade.get(x.un) ?? { qtd: 0, centavos: 0 };
      u.qtd += x.qtd;
      u.centavos += x.centavos;
      porUnidade.set(x.un, u);
    }
    const subtotal = i === 0 && o.subtotalErrado !== undefined ? o.subtotalErrado : soma;
    rows.push(pad([null, null, null, qtd(q), null, num(subtotal / 100)]));
    rows.push([]);
  });

  rows.push(pad(["Totais por unidade:"]));
  const unidades = [...porUnidade.entries()];
  unidades.forEach(([un, t], i) => {
    const valor = i === 0 && o.totalErrado !== undefined ? o.totalErrado : t.centavos;
    rows.push(pad([un, null, null, qtd(t.qtd), null, num(valor / 100)]));
  });
  rows.push(pad(["Gerado em 05/07/2026 14:32", null, null, null, null, null]));
  rows.push(pad(["Página 1 de 1"]));
  return rows;
}

function paraBuffer(out: unknown): ArrayBuffer {
  return out instanceof ArrayBuffer ? out : new Uint8Array(out as ArrayLike<number>).buffer;
}

function pasta(opcoes: (m: number) => OpcoesAba, bookType: XLSX.BookType = "xlsx", nomes = ROTULOS): ArrayBuffer {
  const wb = XLSX.utils.book_new();
  MESES.forEach((_, m) => {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(linhasDaAba(m, opcoes(m))), nomes[m]);
  });
  return paraBuffer(XLSX.write(wb, { type: "array", bookType }));
}

/** ".xls" gravado como HTML — várias tabelas, uma por mês, tudo em texto. */
function pastaHtml(): ArrayBuffer {
  const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;");
  const tabelas = MESES.map((_, m) => {
    const linhas = linhasDaAba(m, { formato: "texto-br", intervalo: true });
    return `<table>${linhas
      .map((r) => `<tr>${r.map((c) => `<td>${c === null ? "" : esc(String(c))}</td>`).join("")}</tr>`)
      .join("")}</table>`;
  });
  const html = `<html><head><meta charset="utf-8"></head><body>${tabelas.join("")}</body></html>`;
  return paraBuffer(new TextEncoder().encode(html));
}

/** ".xls" em XML 2003 (SpreadsheetML), com os números gravados como texto. */
function pastaXml2003(): ArrayBuffer {
  const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;");
  const abas = MESES.map((_, m) => {
    const linhas = linhasDaAba(m, { formato: "texto-br" });
    return `<Worksheet ss:Name="${ROTULOS[m]}"><Table>${linhas
      .map(
        (r) =>
          `<Row>${r
            .map((c) => (c === null ? "<Cell/>" : `<Cell><Data ss:Type="String">${esc(String(c))}</Data></Cell>`))
            .join("")}</Row>`,
      )
      .join("")}</Table></Worksheet>`;
  });
  const xml =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">` +
    abas.join("") +
    `</Workbook>`;
  return paraBuffer(new TextEncoder().encode(xml));
}

const centavos = (v: number) => Math.round(v * 100);

// ---------------------------------------------------------------------------
// Leitura de números
// ---------------------------------------------------------------------------

describe("leitura de números no formato brasileiro", () => {
  it.each([
    ["4.200,00", 4200],
    ["1.234.567,89", 1234567.89],
    ["123,45", 123.45],
    ["0,5", 0.5],
    ["R$ 1.234,56", 1234.56],
    ["R$1.234,56", 1234.56],
    ["1 234,56", 1234.56],
    ["1 234,56", 1234.56],
    ["-1.234,56", -1234.56],
    ["(1.234,56)", -1234.56],
    ["1.234,56-", -1234.56],
    ["1.200", 1200],
    ["12.345", 12345],
    ["1.5", 1.5],
    ["0.500", 0.5],
    ["1234.56", 1234.56],
    ["1,234.56", 1234.56],
    ["10", 10],
    [" 25 ", 25],
    [4200.5, 4200.5],
    [0, 0],
  ])("%s → %s", (entrada, esperado) => {
    expect(lerNumero(entrada)).toBeCloseTo(esperado, 10);
  });

  it.each([["PARAFUSO"], ["01/06/2025"], [""], ["1.23.4"], ["12%"], [null], [undefined], [true], ["UN"]])(
    "%s não é número",
    (entrada) => {
      expect(lerNumero(entrada)).toBeNull();
    },
  );
});

// ---------------------------------------------------------------------------
// Seis abas, todos os formatos
// ---------------------------------------------------------------------------

const formatos: [string, () => ArrayBuffer][] = [
  [".xlsx com números", () => pasta(() => ({ formato: "numero" }))],
  [".xlsx com texto brasileiro", () => pasta(() => ({ formato: "texto-br" }))],
  [".xlsx com R$ no texto", () => pasta(() => ({ formato: "moeda-br" }))],
  [".xls binário", () => pasta(() => ({ formato: "numero" }), "biff8")],
  [".xls binário com texto brasileiro", () => pasta(() => ({ formato: "texto-br" }), "biff8")],
  [".xls em HTML (abas sem nome de mês)", pastaHtml],
  [".xls em XML 2003", pastaXml2003],
  [
    "abas com formatações diferentes entre si",
    () =>
      pasta((m) => ({
        formato: (["numero", "texto-br", "moeda-br"] as Formato[])[m % 3],
        recuo: m % 2,
        rotuloValor: m === 3 ? "Valor total" : undefined,
      })),
  ],
];

describe.each(formatos)("apuração das 6 abas · %s", (_, gerar) => {
  const r = analyzeWorkbook(gerar());

  it("lê as seis abas, na ordem dos meses", () => {
    expect(r.periodos).toHaveLength(6);
    expect(r.periodos.map((p) => p.mes)).toEqual([1, 2, 3, 4, 5, 6]);
    expect(r.leitura.abasIgnoradas).toEqual([]);
  });

  it("apura o faturamento de cada mês exatamente como a soma dos registros da planilha", () => {
    expect(r.periodos.map((p) => centavos(p.totalFaturamento))).toEqual(esperadoCentavos);
    expect(r.evolucaoPeriodos.map((p) => centavos(p.faturamento))).toEqual(esperadoCentavos);
    expect(centavos(r.evolucaoConsolidada.faturamentoTotal)).toBe(
      esperadoCentavos.reduce((a, b) => a + b, 0),
    );
  });

  it("conta todos os registros de venda, sem duplicar nem omitir", () => {
    expect(r.periodos.map((p) => p.produtosLidos)).toEqual(esperadoRegistros);
    expect(r.periodos.map((p) => p.leituraAba.linhasDescartadas)).toEqual(MESES.map(() => []));
  });

  it("apura tudo só pela planilha, sem divergências nem cruzamentos", () => {
    const erros = r.validationReport.diagnostics.filter((d) => d.severity === "error");
    expect(erros.map((d) => `${d.cellRef} ${d.field}: ${d.impact}`)).toEqual([]);
    expect(r.validationReport.valid).toBe(true);
    expect(r.conferencia.abas.map((a) => centavos(a.faturamento))).toEqual(esperadoCentavos);
    expect(r.conferencia.abas.map((a) => a.registrosVenda)).toEqual(esperadoRegistros);
  });

  it("soma as unidades vendidas sem os somatórios por cliente nem os totais por unidade", () => {
    // subtotal rows repeat each client's units and the "Totais por unidade"
    // block repeats the tab's: none of them may be counted again.
    expect(r.periodos.map((p) => p.leituraAba.quantidadeTotal)).toEqual(esperadoQuantidade);
    expect(r.conferencia.abas.map((a) => a.quantidadeTotal)).toEqual(esperadoQuantidade);
    expect(r.evolucaoPeriodos.map((p) => p.quantidade)).toEqual(esperadoQuantidade);
    expect(r.conferencia.consolidado.quantidadeTotal).toBe(esperadoQuantidade.reduce((a, b) => a + b, 0));
  });

  it("classifica a movimentação de cada mês em alta, queda ou estável", () => {
    const esperado = esperadoCentavos.map((v, i) => {
      if (i === 0) return null;
      const ant = esperadoCentavos[i - 1];
      const variacao = ((v - ant) / ant) * 100;
      return Math.abs(variacao) < 5 ? "estavel" : variacao > 0 ? "crescimento" : "queda";
    });
    expect(r.evolucaoPeriodos.map((p) => p.movimentacao)).toEqual(esperado);
  });

  it("mantém os funcionários (*) no faturamento", () => {
    expect(r.periodos.map((p) => centavos(p.faturamentoFuncionarios))).toEqual(esperadoFuncionariosCentavos);
    expect(r.funcionarios.map((f) => f.nome)).toEqual(["*CLIENTE TESTE", "*JOÃO PEREIRA", "*MARIA SILVA"]);
    // their products are sales like any other
    const cafe = r.comparativoProdutos.find((p) => p.produto === "CAFE 500G")!;
    expect(centavos(cafe.valorPorPeriodo[r.periodos[0].nome])).toBe(3100);
    expect(centavos(cafe.valorPorPeriodo[r.periodos[5].nome])).toBe(6200);
    expect(r.comparativoProdutos.some((p) => p.produto === "LEITE INTEGRAL 1L")).toBe(true);
  });

  it("deixa os funcionários fora de clientes novos, perdidos e carteira", () => {
    const nomes = (lista: { nome: string }[]) => lista.map((c) => c.nome);
    const todosOsGrupos = r.transicoes.flatMap((t) => [
      ...nomes(t.novos.clientes),
      ...nomes(t.retornados.clientes),
      ...nomes(t.mantidos.clientes),
      ...nomes(t.perdidos.clientes),
    ]);
    const telas = [
      ...todosOsGrupos,
      ...nomes(r.classificacaoCarteira),
      ...nomes(r.carteira.novos),
      ...nomes(r.carteira.perdidos),
      ...nomes(r.carteira.ativos),
      ...nomes(r.carteira.frequenciaReduzida),
      ...nomes(r.clientesNovos),
      ...nomes(r.clientesPerdidos),
    ];
    expect(telas.filter((n) => n.startsWith("*"))).toEqual([]);

    expect(r.evolucaoPeriodos.map((p) => p.clientes)).toEqual(esperadoClientesCarteira);
    expect(r.evolucaoConsolidada.clientesDistintos).toBe(4);
  });

  it("identifica corretamente novos, perdidos e retornos da carteira", () => {
    const abril = r.transicoes[2]; // março → abril
    expect(abril.novos.clientes.map((c) => c.nome)).toEqual(["GAMA SUPERMERCADOS"]);
    expect(abril.perdidos.clientes.map((c) => c.nome)).toEqual(["BETA MERCEARIA"]);
    // *MARIA stops buying in abril and is not a lost client
    expect(abril.perdidos.total).toBe(1);

    const maio = r.transicoes[3];
    expect(maio.retornados.clientes.map((c) => c.nome)).toEqual(["DELTA PADARIA"]);
    expect(maio.novos.total).toBe(0);

    const junho = r.transicoes[4];
    // *JOÃO buys for the first time in junho and is not a new client
    expect(junho.novos.total).toBe(0);

    expect(r.carteira.perdidos.map((c) => c.nome)).toEqual(["BETA MERCEARIA"]);
  });

  it("fecha a ponte de receita de cada mês, com a parcela dos funcionários", () => {
    for (const t of r.transicoes) {
      const p = t.ponte;
      const soma = p.novos + p.retornados + p.expansao + p.contracao + p.perdidos + p.funcionarios;
      expect(centavos(soma)).toBe(centavos(t.diferencaFaturamento));
    }
    // março → abril: *MARIA (R$ 31,00) stopped buying
    expect(centavos(r.transicoes[2].ponte.funcionarios)).toBe(-3100);
  });

  it("totaliza cada cliente e cada produto pelos valores da planilha", () => {
    const alfa = r.classificacaoCarteira.find((c) => c.nome === "ALFA COMERCIO LTDA")!;
    MESES.forEach((_, m) => {
      const esperado = compradoresDoMes(m)[0].vendas.reduce((s, x) => s + x.centavos, 0);
      expect(centavos(alfa.faturamentoPorPeriodo[r.periodos[m].nome])).toBe(esperado);
    });

    const arroz = r.comparativoProdutos.find((p) => p.codigo === "5001")!;
    MESES.forEach((_, m) => {
      const esperado = compradoresDoMes(m)
        .flatMap((c) => c.vendas)
        .filter((x) => x.codigo === "5001")
        .reduce((s, x) => s + x.centavos, 0);
      expect(centavos(arroz.valorPorPeriodo[r.periodos[m].nome])).toBe(esperado);
    });
    expect(arroz.qtdPorPeriodo[r.periodos[0].nome]).toBe(1210);
  });

  it("calcula o ticket médio só com os clientes da carteira", () => {
    const jan = r.periodos[0];
    const semFuncionarios = esperadoCentavos[0] - esperadoFuncionariosCentavos[0];
    expect(centavos(r.evolucaoPeriodos[0].ticketMedio)).toBe(
      Math.round(semFuncionarios / esperadoClientesCarteira[0]),
    );
    expect(jan.clientesCarteira.size).toBe(3);
    expect(jan.funcionarios.size).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// Divergências entre planilha e sistema
// ---------------------------------------------------------------------------

describe("a planilha é a única fonte: nada é cruzado com totais impressos", () => {
  it("subtotal impresso diferente não altera o faturamento nem gera divergência", () => {
    const r = analyzeWorkbook(
      pasta((m) => ({ formato: "texto-br", subtotalErrado: m === 2 ? 999999 : undefined })),
    );
    expect(centavos(r.periodos[2].totalFaturamento)).toBe(esperadoCentavos[2]);
    expect(r.periodos[2].leituraAba.quantidadeTotal).toBe(esperadoQuantidade[2]);
    expect(r.validationReport.valid).toBe(true);
    expect(r.inconsistencias.filter((i) => /subtotal|impresso|confer/i.test(i))).toEqual([]);
  });

  it("total impresso na aba diferente é ignorado", () => {
    const r = analyzeWorkbook(pasta((m) => ({ formato: "numero", totalErrado: m === 5 ? 1 : undefined })));
    expect(centavos(r.periodos[5].totalFaturamento)).toBe(esperadoCentavos[5]);
    expect(r.validationReport.valid).toBe(true);
  });

  it("valor solto na aba, fora de qualquer cliente, é apontado e não entra no faturamento", () => {
    const r = analyzeWorkbook(pasta((m) => ({ formato: "numero", linhaSolta: m === 0 })));
    const jan = r.conferencia.abas[0];
    expect(jan.valorDescartado).toBe(123.45);
    expect(centavos(jan.faturamento)).toBe(esperadoCentavos[0]);
    expect(r.inconsistencias.some((i) => i.includes("fora de qualquer bloco de cliente"))).toBe(true);
  });

  it("o mesmo mês em duas abas é apontado como erro", () => {
    const nomes = [...ROTULOS];
    nomes[5] = "maio 2026 (2)";
    const r = analyzeWorkbook(pasta(() => ({ formato: "numero" }), "xlsx", nomes));
    expect(
      r.validationReport.diagnostics.some((d) => d.severity === "error" && d.impact.includes("mais de uma aba")),
    ).toBe(true);
    // both tabs are still visible as separate columns — none overwrites the other
    expect(new Set(r.periodos.map((p) => p.nome)).size).toBe(6);
  });
});

// ---------------------------------------------------------------------------
// Regressões de leitura
// ---------------------------------------------------------------------------

function abasCruas(abas: Record<string, Celula[][]>): ArrayBuffer {
  const wb = XLSX.utils.book_new();
  for (const [n, rows] of Object.entries(abas)) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), n);
  return paraBuffer(XLSX.write(wb, { type: "array", bookType: "xlsx" }));
}
const CAB = ["Código", "Descrição", "UN", "Qtd. venda", "Preço médio venda", "Venda bruta"];
const abaSimples: Celula[][] = [["Cliente: 1 - A"], CAB, ["10", "X", "UN", 1, 10, 10]];

describe("regressões da leitura das linhas", () => {
  it("lê 4.200,00 como quatro mil e duzentos em .xls salvo como HTML", () => {
    const r = analyzeWorkbook(pastaHtml());
    const alfa = [...r.periodos[0].clientes.values()].find((c) => c.nome === "ALFA COMERCIO LTDA")!;
    expect(alfa.produtos[0].quantidade).toBe(1200);
    expect(alfa.produtos[0].valorUnitario).toBe(3.5);
    expect(alfa.produtos[0].valorTotal).toBe(4200);
  });

  it('linha "Total:" com valor é somatório do cliente, não um produto', () => {
    const r = analyzeWorkbook(
      abasCruas({
        junho: [["Cliente: 1 - A"], CAB, ["10", "X", "UN", 2, 5, 10], ["Total:", null, null, 2, null, 10]],
        julho: abaSimples,
      }),
    );
    const a = [...r.periodos[0].clientes.values()][0];
    expect(a.produtos).toHaveLength(1);
    expect(a.totalGeral).toBe(10);
    expect(r.periodos[0].leituraAba.quantidadeTotal).toBe(2);
  });

  it("venda sem código de produto entra no faturamento e é sinalizada", () => {
    const r = analyzeWorkbook(
      abasCruas({
        junho: [["Cliente: 1 - A"], CAB, ["10", "X", "UN", 2, 5, 10], [null, "BRINDE AVULSO", "UN", 1, 7.5, 7.5], [null, null, null, 3, null, 17.5]],
        julho: abaSimples,
      }),
    );
    expect(r.periodos[0].totalFaturamento).toBe(17.5);
    expect(r.periodos[0].produtosLidos).toBe(2);
    expect(r.validationReport.diagnostics.some((d) => d.impact.includes("sem código"))).toBe(true);
    expect(r.periodos[0].leituraAba.quantidadeTotal).toBe(3);
  });

  it('cabeçalho de cliente com "Cliente:" numa célula e o nome na seguinte', () => {
    const r = analyzeWorkbook(
      abasCruas({
        junho: [["Cliente:", "1 - *ANA", "Fone: (11) 1111-1111"], CAB, ["10", "X", "UN", 1, 10, 10]],
        julho: abaSimples,
      }),
    );
    const ana = [...r.periodos[0].clientes.values()][0];
    expect(ana.codigoCliente).toBe("1");
    expect(ana.nome).toBe("*ANA");
    expect(ana.fone).toBe("(11) 1111-1111");
    expect(ana.funcionario).toBe(true);
  });

  it("produto com o mesmo código e descrição com espaços/caixa diferentes é um só", () => {
    const r = analyzeWorkbook(
      abasCruas({
        junho: [["Cliente: 1 - A"], CAB, ["10", "Arroz  Tipo 1", "UN", 1, 10, 10]],
        julho: [["Cliente: 1 - A"], CAB, ["10", "ARROZ TIPO 1", "UN", 2, 10, 20]],
      }),
    );
    expect(r.comparativoProdutos).toHaveLength(1);
    expect(r.comparativoProdutos[0].valorPorPeriodo).toEqual({ junho: 10, julho: 20 });
  });
});

describe("quantidade vendida", () => {
  const qtd = (r: AnalysisResult, i = 0) => r.periodos[i].leituraAba.quantidadeTotal;

  it('usa somente a coluna "Qtd. venda" de cada produto', () => {
    const cab = ["Código", "Descrição", "UN", "Qtd. venda", "Qtd. bonificação", "Preço médio venda", "Venda bruta"];
    const r = analyzeWorkbook(
      abasCruas({
        junho: [
          ["Cliente: 1 - A"],
          cab,
          ["10", "X", "UN", 10, 2, 5, 50],
          ["11", "Y", "UN", 3, null, 4, 12],
          [null, null, null, 13, 2, null, 62],
        ],
        julho: abaSimples,
      }),
    );
    expect(r.periodos[0].clientes.get("cod:1")!.produtos.map((p) => p.quantidade)).toEqual([10, 3]);
    expect(qtd(r)).toBe(13);
    expect(r.evolucaoPeriodos[0].quantidade).toBe(13);
    expect(r.conferencia.abas[0].quantidadeTotal).toBe(13);
    expect(r.periodos[0].totalFaturamento).toBe(62);
    expect(r.inconsistencias).toEqual([]);
  });

  it('prefere a coluna "Qtd. venda" mesmo quando outra coluna de quantidade vem antes', () => {
    const cab = ["Código", "Descrição", "Qtd. embalagem", "Qtde. caixa", "Qtd. venda", "Qtd. devolvida", "Preço médio", "Venda bruta"];
    const r = analyzeWorkbook(
      abasCruas({
        junho: [["Cliente: 1 - A"], cab, ["10", "X", 12, 2, 4, 1, 10, 40]],
        julho: abaSimples,
      }),
    );
    // never a money column, never another quantity column
    expect(qtd(r)).toBe(4);
    expect(r.periodos[0].totalFaturamento).toBe(40);
  });

  it.each([
    ["sem código nem descrição", [null, null, null, 5, null, 50]],
    ['com "Total do cliente:" em outra coluna', [null, null, "Total do cliente:", 5, null, 50]],
    ["com o nome do cliente na descrição", [null, "ALFA LTDA", null, 5, null, 50]],
    ["com o código do cliente", ["1", null, null, 5, null, 50]],
    ["repetindo a soma das vendas do bloco", [null, "TOTAL", null, 5, null, 50]],
    ["sem rótulo, repetindo a soma das vendas", [null, "  ", null, 5, null, 50]],
    ['com "Qtde. total:" no código', ["Qtde. total:", null, null, 5, null, 50]],
    ["com a unidade de medida na coluna de código", ["UN", null, null, 5, null, 50]],
    ["com a unidade de medida e só a quantidade", ["UN", null, null, 5, null, null]],
    ["só com a quantidade", [null, null, null, 5, null, null]],
    ["com a unidade de medida na coluna de descrição", [null, "UN", null, 5, null, 50]],
    ["com a unidade de medida em todas as colunas de texto", ["KG", "KG", "KG", 5, null, 50]],
    ["com um número na coluna de código", [5, null, null, 5, null, 50]],
    ["com código mas sem produto", ["10", null, null, 5, null, 50]],
    ["com código e nome do cliente", [null, "1 - ALFA LTDA", null, 5, null, 50]],
  ])("não conta de novo o somatório do cliente %s", (_, linhaSoma) => {
    const r = analyzeWorkbook(
      abasCruas({
        junho: [
          ["Cliente: 1 - ALFA LTDA"],
          CAB,
          ["10", "X", "UN", 2, 10, 20],
          ["11", "Y", "UN", 3, 10, 30],
          linhaSoma as Celula[],
          [],
          ["Totais por unidade:"],
          ["UN", null, null, 5, null, 50],
        ],
        julho: abaSimples,
      }),
    );
    expect(qtd(r)).toBe(5);
    expect(r.periodos[0].produtosLidos).toBe(2);
    expect(r.periodos[0].totalFaturamento).toBe(50);
    expect(r.evolucaoPeriodos[0].quantidade).toBe(5);
    expect(r.conferencia.abas[0].quantidadeTotal).toBe(5);
    // the client's own sum is not an error of the spreadsheet
    expect(r.inconsistencias).toEqual([]);
  });

  it("soma a Qtd. venda de cada produto de vários clientes, cada um com seu somatório abaixo", () => {
    // as the report prints it: products of the client, then the client's sum
    const r = analyzeWorkbook(
      abasCruas({
        junho: [
          ["Vendas de produto no período"],
          ["Cliente: 1001 - ALFA LTDA"],
          CAB,
          ["5001", "ARROZ 5KG", "UN", 120, 3.5, 420],
          ["5002", "FEIJAO 1KG", "UN", 50, 8, 400],
          ["5010", "CARNE BOVINA", "KG", 12.5, 40, 500],
          ["UN", null, null, 182.5, null, 1320],
          [],
          ["Cliente: 1002 - BETA MERCEARIA"],
          CAB,
          ["5004", "OLEO 900ML", "CX", 3, 100, 300],
          [null, null, "CX", 3, null, 300],
          [],
          ["Cliente: 2001 - *MARIA SILVA"],
          CAB,
          ["5007", "CAFE 500G", "UN", 2, 15.5, 31],
          [null, null, null, 2, null, 31],
          [],
          ["Totais por unidade:"],
          ["UN", null, null, 172, null, 851],
          ["KG", null, null, 12.5, null, 500],
          ["CX", null, null, 3, null, 300],
        ],
        julho: abaSimples,
      }),
    );
    // 120 + 50 + 12,5 + 3 + 2 — each product's Qtd. venda, once
    expect(qtd(r)).toBe(187.5);
    expect(r.evolucaoPeriodos[0].quantidade).toBe(187.5);
    expect(r.conferencia.abas[0].quantidadeTotal).toBe(187.5);
    expect(r.periodos[0].produtosLidos).toBe(5);
    expect(r.periodos[0].totalFaturamento).toBe(1651);
  });

  it("conta cada venda uma única vez, inclusive com o cliente em dois blocos", () => {
    const r = analyzeWorkbook(
      abasCruas({
        junho: [
          ["Cliente: 1 - A"],
          CAB,
          ["10", "X", "UN", 2, 10, 20],
          [null, null, null, 2, null, 20],
          ["Página 1 de 2"],
          ["Cliente: 1 - A"],
          CAB,
          ["11", "Y", "UN", 7, 1, 7],
          [null, null, null, 7, null, 7],
        ],
        julho: abaSimples,
      }),
    );
    expect(qtd(r)).toBe(9);
    expect(r.periodos[0].totalFaturamento).toBe(27);
  });
});

describe("layout real do relatório de vendas", () => {
  // As the ERP exports it: the product header once at the top, a "Cliente:" line
  // per client, its products, a sum line ("" code) only when the client bought
  // two or more items, then the tab's grand total and the per-unit totals —
  // printed without a label and shifted: unit | qty | qty | VALUE | 0 | 0.
  const aba = (fator: number): Celula[][] => [
    ["Código", "Descrição", "UN", "Qtd. venda", "Preço médio venda", "Venda bruta"],
    [null, null, null, null, null, null],
    ["Cliente: 22908 - *ANA MARIA - Data da última compra: 08/05/2026", null, null, null, null, null],
    [null, null, null, null, null, null],
    ["2275", "(UAI) PAO DE BATATA REQUEIJAO 20 UND", "PCT", 2 * fator, 83.18, 166.36 * fator],
    [null, null, null, null, null, null],
    ["Cliente: 1001 - PADARIA ALFA - Data da última compra: 15/05/2026 - Celular: (0xx48)99999-0000", null, null, null, null, null],
    [null, null, null, null, null, null],
    ["3046", "(NEVASCA) MINI CROISSANT CHOCOLATE 1KG", "PCT", 2 * fator, 43.62, 87.24 * fator],
    ["5531", "(O PADEIRO) PUDIM LEITE MOÇA 135G 12 UND", "CX", 1 * fator, 70, 70 * fator],
    ["1158", "(FERFUR) MINI BOLINHA QUEIJO 1,5KG", "KG", 3 * fator, 10, 30 * fator],
    ["", null, null, 6 * fator, null, 187.24 * fator],
    [null, null, null, null, null, null],
    ["", null, null, 8 * fator, null, 353.6 * fator],
    [null, null, null, null, null, null],
    ["PCT", 4 * fator, 4 * fator, 253.6 * fator, 0, 0],
    ["CX", 1 * fator, 1 * fator, 70 * fator, 0, 0],
    ["KG", 3 * fator, 3 * fator, 30 * fator, 0, 0],
    ["BANDEJ", 0, 0, 0, 0, 0],
    [null, null, null, null, null, null],
    ["Quantidade total de clientes listados: 2", null, null, null, null, null],
    ["Gerado em 13/11/26 09:46:15", null, null, null, null, null],
  ];

  it("soma só a Qtd. venda dos produtos: sem a soma do cliente, o total geral e os totais por unidade", () => {
    const r = analyzeWorkbook(abasCruas({ maio: aba(1), junho: aba(2) }));
    // 2 + 2 + 1 + 3 — the sum lines (6, 8) and the per-unit rows (whose third
    // number is money, R$ 253,60) are never counted as units
    expect(r.periodos.map((p) => p.leituraAba.quantidadeTotal)).toEqual([8, 16]);
    expect(r.evolucaoPeriodos.map((p) => p.quantidade)).toEqual([8, 16]);
    expect(r.conferencia.abas.map((a) => a.quantidadeTotal)).toEqual([8, 16]);
    expect(r.periodos.map((p) => p.totalFaturamento)).toEqual([353.6, 707.2]);
    expect(r.periodos.map((p) => p.produtosLidos)).toEqual([4, 4]);
    // 1 client + 1 employee = the 2 clients listed by the report
    expect(r.periodos[0].clientesCarteira.size).toBe(1);
    expect(r.periodos[0].funcionarios.size).toBe(1);
    expect(r.inconsistencias).toEqual([]);
    expect(r.validationReport.valid).toBe(true);
  });
});

describe("movimentação de vendas entre os meses", () => {
  it("aponta alta, queda e estabilidade a partir do faturamento da planilha", () => {
    const aba = (valor: number): Celula[][] => [["Cliente: 1 - A"], CAB, ["10", "X", "UN", 1, valor, valor]];
    const r = analyzeWorkbook(abasCruas({ marco: aba(100), abril: aba(150), maio: aba(100), junho: aba(102) }));
    expect(r.evolucaoPeriodos.map((p) => p.movimentacao)).toEqual([null, "crescimento", "queda", "estavel"]);
    expect(r.evolucaoPeriodos.map((p) => p.variacaoFaturamento)).toEqual([null, 50, -33.3333, 2]);
  });
});

describe("importação de novos arquivos", () => {
  it("cada arquivo importado gera uma apuração independente", () => {
    const primeiro = analyzeWorkbook(pasta(() => ({ formato: "numero" })));
    const segundo = analyzeWorkbook(
      abasCruas({ junho: abaSimples, julho: [["Cliente: 1 - A"], CAB, ["10", "X", "UN", 3, 10, 30]] }),
    );
    const terceiro = analyzeWorkbook(pasta(() => ({ formato: "numero" })));

    expect(segundo.periodos.map((p) => p.totalFaturamento)).toEqual([10, 30]);
    expect(segundo.funcionarios).toEqual([]);
    expect(terceiro.periodos.map((p) => centavos(p.totalFaturamento))).toEqual(
      primeiro.periodos.map((p) => centavos(p.totalFaturamento)),
    );
  });
});
