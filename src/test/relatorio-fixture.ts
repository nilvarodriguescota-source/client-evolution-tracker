import * as XLSX from "xlsx";

/**
 * Builds workbooks in the shape the ERP actually exports: a title, one block per
 * client ("Cliente: 1234 - NOME - Data: 01/06/2025 a 30/06/2025"), a product
 * header, product rows, a subtotal row and the footer lines — one tab per month.
 *
 * Used by the tests that prove every month of a real three-month file is read.
 */

export type Rows = (string | number | null)[][];

export interface ItemFixture {
  codigo: string;
  produto: string;
  qtd: number;
  preco: number;
  /** Gross value as printed in the sheet; defaults to qtd × preço */
  total?: number;
}

export interface ClienteFixture {
  codigo: string;
  nome: string;
  itens: ItemFixture[];
  /** Subtotal as printed; defaults to the sum of the items */
  subtotal?: number;
  /** Appended to the client line as "- Data da última compra: …" */
  dataUltimaCompra?: string;
  /** Appended as "- Fone: …" */
  fone?: string;
  /** Appended as "- Celular: …" */
  celular?: string;
}

export interface AbaFixture {
  clientes: ClienteFixture[];
  /** Date range printed in each client header, e.g. "01/06/2025 a 30/06/2025" */
  intervalo?: string;
  /** Leave the period out of the tab name to force detection by content */
  cabecalhoDeProdutos?: string[];
  /** Extra blank columns before the report (indented exports) */
  recuo?: number;
  /** Overrides the printed "Totais por unidade" value (to simulate a divergence) */
  totalImpresso?: number;
}

const CABECALHO_PADRAO = [
  "Código",
  "Descrição",
  "UN",
  "Qtd. venda",
  "Preço médio venda",
  "Venda bruta",
];

export function abaRelatorio(aba: AbaFixture): Rows {
  const recuo = aba.recuo ?? 0;
  const pad = (cells: (string | number | null)[]): (string | number | null)[] => [
    ...Array<null>(recuo).fill(null),
    ...cells,
  ];

  const rows: Rows = [];
  rows.push(pad(["Vendas de produto no período"]));
  rows.push([]);
  // What the tab adds up to, per unit — printed at the end, as the ERP does.
  const porUnidade = new Map<string, { qtd: number; valor: number }>();

  for (const cliente of aba.clientes) {
    // The exports append labelled fields to the client line, in this order.
    const partes = [aba.intervalo ? `Data: ${aba.intervalo}` : ""];
    if (cliente.dataUltimaCompra) partes.push(`Data da última compra: ${cliente.dataUltimaCompra}`);
    if (cliente.fone) partes.push(`Fone: ${cliente.fone}`);
    if (cliente.celular) partes.push(`Celular: ${cliente.celular}`);
    const sufixo = partes.filter(Boolean).map((p) => ` - ${p}`).join("");

    rows.push(pad([`Cliente: ${cliente.codigo} - ${cliente.nome}${sufixo}`]));
    rows.push(pad(aba.cabecalhoDeProdutos ?? CABECALHO_PADRAO));

    let soma = 0;
    let quantidade = 0;
    for (const item of cliente.itens) {
      const total = item.total ?? Number((item.qtd * item.preco).toFixed(2));
      soma += total;
      quantidade += item.qtd;
      rows.push(pad([item.codigo, item.produto, "UN", item.qtd, item.preco, total]));
      const u = porUnidade.get("UN") ?? { qtd: 0, valor: 0 };
      u.qtd += item.qtd;
      u.valor += total;
      porUnidade.set("UN", u);
    }

    rows.push(pad([null, null, null, quantidade, null, cliente.subtotal ?? Number(soma.toFixed(2))]));
    rows.push([]);
  }

  rows.push(pad(["Totais por unidade:"]));
  if (porUnidade.size === 0) rows.push(pad(["UN", null, null, 0, null, 0]));
  for (const [un, t] of porUnidade) {
    rows.push(pad([un, null, null, Number(t.qtd.toFixed(4)), null, aba.totalImpresso ?? Number(t.valor.toFixed(2))]));
  }
  rows.push(pad(["Gerado em 05/09/2025 14:32"]));
  return rows;
}

export function montarArquivo(
  abas: Record<string, AbaFixture>,
  bookType: XLSX.BookType = "xlsx",
  ocultas: string[] = [],
): ArrayBuffer {
  const wb = XLSX.utils.book_new();
  for (const [nome, aba] of Object.entries(abas)) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(abaRelatorio(aba)), nome);
  }
  if (ocultas.length) {
    wb.Workbook = {
      Sheets: wb.SheetNames.map((name) => ({ name, Hidden: ocultas.includes(name) ? 1 : 0 })),
    };
  }
  const out = XLSX.write(wb, { type: "array", bookType });
  return out instanceof ArrayBuffer ? out : new Uint8Array(out).buffer;
}
