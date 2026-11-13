import {
  descobrirAbas,
  lerPastaDeTrabalho,
  lerCabecalhoCliente,
  lerCabecalhoProdutos,
  lerPeriodoDoNome,
  normalizarTexto,
  ordenarPorPeriodo,
  COLUNA_AUSENTE,
  type AbaDescoberta,
  type AbaIgnorada,
  type ColumnMap,
  type SheetRows,
} from "./sheet-discovery";
import { lerNumero, ehNumero, paraNumero, somarValores } from "./numero-br";
import { resumirPlanilha, type Conferencia } from "./conferencia";
import {
  ExcelValidator,
  preciseMultiply,
  preciseRound,
  preciseSum,
  toDisplay,
  colIndexToLetter,
  type ValidationReport,
  type ValidationDiagnostic,
  type CellOrigin,
  type TracedValue,
} from "./excel-validator";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface ProductDetail {
  produto: string;
  codigo: string;
  unidade: string;
  valorUnitario: number;
  quantidade: number;
  valorTotal: number;
  validado: boolean;
  /** Internal calculated value with full precision */
  valorCalculado: number;
  /** Difference between reported and calculated */
  divergencia: number;
  /** Traceability: origin cell references for each field */
  origens: {
    codigo: CellOrigin;
    produto: CellOrigin;
    unidade: CellOrigin;
    quantidade: CellOrigin;
    valorUnitario: CellOrigin;
    valorTotal: CellOrigin;
  };
}

export interface ClientData {
  /**
   * Identity across months: the client code when the report carries one, the
   * normalised name otherwise. Never the raw header text — the exports append
   * the client's last purchase date to it, which changes every month.
   */
  chave: string;
  /** Client name alone, for display */
  nome: string;
  codigoCliente: string;
  /** Last purchase date as printed by the report, e.g. "08/09/2026" */
  dataUltimaCompra?: string;
  fone?: string;
  celular?: string;
  /** Every labelled field the client line carries, by normalised label */
  campos: Record<string, string>;
  produtos: ProductDetail[];
  totalGeral: number;
  /** Sum recalculated from products with full precision */
  totalRecalculado: number;
  /** Whether totalGeral matches sum of products */
  totalConsistente: boolean;
  /** Origin of client header row */
  origemLinha: number;
  /**
   * Identification starts with "*": an employee. Counts in revenue, products and
   * sales figures; never in new/lost clients, portfolio or retention.
   */
  funcionario: boolean;
}

/** What one tab of the spreadsheet holds, read straight from its sale rows. */
export interface LeituraAba {
  /** Sale rows read (one per product line of a client) */
  registrosVenda: number;
  /** Units sold: Σ of the "Qtd. venda" column over the sale rows only */
  quantidadeTotal: number;
  produtosDistintos: number;
  /** Σ "Venda bruta" of the sale rows — the revenue of the tab */
  faturamentoRegistros: number;
  /** Rows carrying a value that could not be placed as a sale or a total */
  linhasDescartadas: { linha: number; motivo: string; valor: number }[];
}

export interface PeriodData {
  /** Label used across the analysis (tab name, or the period read from content) */
  nome: string;
  /** Tab name exactly as it appears in the file */
  sheetName: string;
  /**
   * Every buyer of the month, employees included, keyed by ClientData.chave.
   * The revenue and product figures are built from this map.
   */
  clientes: Map<string, ClientData>;
  /** Buyers of the month that are clients of the portfolio (no employees) */
  clientesCarteira: Map<string, ClientData>;
  /** Employees ("*" in the identification) that bought in the month */
  funcionarios: Map<string, ClientData>;
  /** Revenue of the month: Σ of every sale row of the tab, employees included */
  totalFaturamento: number;
  /** Revenue of the portfolio clients only */
  faturamentoCarteira: number;
  /** Revenue of the employees' purchases */
  faturamentoFuncionarios: number;
  /** What the tab printed and where every value went */
  leituraAba: LeituraAba;
  /** Validation report for this period */
  validationReport: ValidationReport;
  /** How the period was identified */
  origemPeriodo: AbaDescoberta["origemPeriodo"];
  mes?: number;
  ano?: number;
  /** Tab is hidden in the file but was read all the same */
  oculta: boolean;
  /** Evidence that the tab was really processed, not merely opened */
  linhasLidas: number;
  produtosLidos: number;
}

export interface ProductComparison {
  /** Identity of the product: code + normalised description */
  chave: string;
  produto: string;
  codigo: string;
  qtdAnterior: number;
  qtdAtual: number;
  variacaoQtd: number;
  valorAnterior: number;
  valorAtual: number;
  variacaoValor: number;
  /** Revenue per period name, across the whole analysis window */
  valorPorPeriodo: Record<string, number>;
  /** Quantity per period name, across the whole analysis window */
  qtdPorPeriodo: Record<string, number>;
  /** Months of the window with any sale of this product */
  mesesComVenda: number;
  /** First month of the window → last, in % */
  variacaoPeriodo: number | null;
  /** Where the product is heading across the whole window */
  tendencia: Tendencia;
  /** Revenue in the whole window — context for ranking, not the analysis itself */
  valorTotalPeriodo: number;
}

/** One client inside a movement group, with the fields the report carries. */
export interface ClienteMovimento {
  /** Identity across months */
  chave: string;
  nome: string;
  codigoCliente?: string;
  /** Last purchase date as printed by the report in the month of this movement */
  dataUltimaCompra?: string;
  fone?: string;
  celular?: string;
  /**
   * Revenue of this client in the month the movement is measured in: the newer
   * month for novos/retornados/mantidos, the older one for perdidos.
   */
  faturamento: number;
}

/** A set of clients moving between two consecutive months. */
export interface GrupoClientes {
  /** One entry per client, sorted by name — never a single joined string */
  clientes: ClienteMovimento[];
  total: number;
  /** Sum of the clients' revenue in the month the group is measured in */
  faturamento: number;
}

/**
 * What happened between two consecutive months. One transition per pair, so a
 * six-month file yields five of them: m1→m2, m2→m3, … m5→m6.
 */
export interface TransicaoPeriodo {
  de: string;
  para: string;
  faturamentoDe: number;
  faturamentoPara: number;
  /** % change; null when the earlier month billed nothing */
  variacaoFaturamento: number | null;
  /** Absolute change in currency */
  diferencaFaturamento: number;
  clientesDe: number;
  clientesPara: number;
  variacaoClientes: number | null;
  ticketMedioDe: number;
  ticketMedioPara: number;
  variacaoTicketMedio: number | null;
  /** Bought in `para` and in no earlier month of the window */
  novos: GrupoClientes;
  /** Bought in `para`, skipped `de`, but had bought earlier in the window */
  retornados: GrupoClientes;
  /** Bought in both months */
  mantidos: GrupoClientes;
  /** Bought in `de` and not in `para` */
  perdidos: GrupoClientes;
  /**
   * Where the revenue change came from. The five parts add up to
   * `diferencaFaturamento`, so the month-over-month move is fully explained.
   */
  ponte: {
    novos: number;
    retornados: number;
    /** Kept clients that bought more (positive) */
    expansao: number;
    /** Kept clients that bought less (negative) */
    contracao: number;
    /** Revenue the clients that stopped buying used to bring (negative) */
    perdidos: number;
    /**
     * Change in the employees' purchases. They are revenue but not clients, so
     * they get their own part and the bridge still closes on the month's change.
     */
    funcionarios: number;
    total: number;
  };
  /** mantidos ÷ clientes do mês anterior, em % */
  taxaRetencao: number | null;
}

/** One month of the analysis window, summarised for the month-by-month comparison. */
export interface PeriodoResumo {
  nome: string;
  faturamento: number;
  clientes: number;
  /** faturamento ÷ nº de clientes */
  ticketMedio: number;
  /** % vs the previous month in the window; null on the first month */
  variacaoFaturamento: number | null;
  variacaoClientes: number | null;
  /** Clients billing in this month that did not bill in the previous one */
  entraram: number | null;
  /** Clients billing in the previous month that did not bill in this one */
  sairam: number | null;
  /** First purchase of the whole window happened in this month */
  novos: number | null;
  /** Came back after skipping the previous month */
  retornados: number | null;
  /** Bought in this month and in the previous one */
  mantidos: number | null;
  /** Bought in the previous month and not in this one */
  perdidos: number | null;
  /** Revenue the novos brought in this month */
  faturamentoNovos: number | null;
  /** Revenue the perdidos had in the previous month */
  faturamentoPerdidos: number | null;
  taxaRetencao: number | null;
  /** Units sold in the month: Σ "Qtd. venda" of the product rows */
  quantidade: number;
  /** % change of the units sold vs the previous month; null on the first month */
  variacaoQuantidade: number | null;
  /**
   * Sales movement against the previous month, read from the revenue of the
   * spreadsheet: "crescimento" (alta), "queda" or "estavel" (within ±5%).
   * Null on the first month, which has nothing before it.
   */
  movimentacao: Tendencia | null;
  /** Employees that bought in the month — outside the client count */
  funcionarios: number;
  /** Their purchases — inside `faturamento` */
  faturamentoFuncionarios: number;
}

/** An employee ("*") with their purchases, kept apart from the client analysis. */
export interface FuncionarioResumo {
  chave: string;
  nome: string;
  codigoCliente?: string;
  faturamentoPorPeriodo: Record<string, number>;
  faturamentoTotal: number;
  periodosPresente: string[];
}

export type Tendencia = "crescimento" | "queda" | "estavel";

/** The whole window read as one movement, for the executive summary. */
export interface EvolucaoConsolidada {
  meses: number;
  primeiro: string;
  ultimo: string;
  faturamentoInicial: number;
  faturamentoFinal: number;
  /** Sum of every month — context only; the analysis stays comparative */
  faturamentoTotal: number;
  faturamentoMedio: number;
  /** First month → last month, in % */
  variacaoTotal: number | null;
  /** Average of the month-over-month changes */
  variacaoMediaMensal: number | null;
  clientesInicial: number;
  clientesFinal: number;
  variacaoClientes: number | null;
  ticketMedioInicial: number;
  ticketMedioFinal: number;
  variacaoTicketMedio: number | null;
  melhorMes: { nome: string; faturamento: number };
  piorMes: { nome: string; faturamento: number };
  /** Months that grew / shrank against the month before */
  mesesEmAlta: number;
  mesesEmQueda: number;
  tendencia: Tendencia;
  /** Distinct clients seen in any month */
  clientesDistintos: number;
  /** Bought in every single month of the window */
  clientesRecorrentes: number;
  /** First purchase happened after the first month */
  novosNoPeriodo: number;
  /** Two or more months without buying by the end of the window */
  perdidosNoPeriodo: number;
  /** Average of the monthly retention rates */
  taxaRetencaoMedia: number | null;
}

export type StatusCarteira =
  | "ativo"
  | "frequencia_reduzida"
  | "risco_abandono"
  | "inativo";

export interface ClienteCarteira {
  /** Identity across months */
  chave: string;
  nome: string;
  codigoCliente?: string;
  /** Last purchase date as printed in the last month the client bought */
  dataUltimaCompra?: string;
  fone?: string;
  celular?: string;
  status: StatusCarteira;
  ultimoPeriodo: string;
  totalPeriodosPresente: number;
  periodosPresente: string[];
  faturamentoUltimoPeriodo: number;
  /** Revenue in every month of the window, zero where the client did not buy */
  faturamentoPorPeriodo: Record<string, number>;
  /** First month with a purchase → last month with a purchase */
  variacaoPeriodo: number | null;
  /** Where this client is heading across the window */
  tendencia: Tendencia;
}

/**
 * Client portfolio read over the whole analysis window — every period found in
 * the file — as opposed to the strict month-over-month comparison kept in
 * AnalysisResult.
 */
export interface CarteiraResumo {
  /** Period names in the window, oldest first */
  periodos: string[];
  /** Clients that bought in the current period */
  ativos: ClientData[];
  /** Bought now and in no earlier period of the window — genuinely new */
  novos: ClientData[];
  /** Bought now, absent from the previous period, but present earlier — came back */
  retornados: ClientData[];
  /** Absent from the current period only (last purchase one period ago) */
  frequenciaReduzida: ClienteCarteira[];
  /** No purchase for two or more periods — effectively out of the portfolio */
  perdidos: ClienteCarteira[];
  /** Revenue of `novos` in the current period */
  faturamentoNovos: number;
  /** Revenue of `perdidos`, measured in each one's last active period */
  faturamentoPerdidos: number;
}

/** What the reader found in the file — so a missing tab is visible, not silent. */
export interface LeituraPlanilha {
  /** Every tab name in the file, in file order */
  abasNoArquivo: string[];
  /** Tabs read as periods, oldest first */
  abasLidas: string[];
  /** Tabs left out, each with the reason why */
  abasIgnoradas: AbaIgnorada[];
}

export interface AnalysisResult {
  periodoAnterior: PeriodData;
  periodoAtual: PeriodData;
  /** Every period found in the file, oldest first — no cap */
  periodos: PeriodData[];
  /** Tab-by-tab account of what was read and what was left out */
  leitura: LeituraPlanilha;
  clientesNovos: ClientData[];
  clientesPerdidos: ClientData[];
  clientesMantidos: string[];
  faturamentoNovos: number;
  faturamentoPerdidos: number;
  variacaoFaturamento: number;
  comparativoProdutos: ProductComparison[];
  inconsistencias: string[];
  /** Portfolio classification for every client found in any period of the file */
  classificacaoCarteira: ClienteCarteira[];
  /** Portfolio figures computed over the whole window (drives the KPI cards) */
  carteira: CarteiraResumo;
  /** Month-by-month comparison across the whole window, oldest first */
  evolucaoPeriodos: PeriodoResumo[];
  /** One entry per pair of consecutive months: m1→m2, m2→m3, … */
  transicoes: TransicaoPeriodo[];
  /** The whole window read as a single movement */
  evolucaoConsolidada: EvolucaoConsolidada;
  /** Employees found in the file, with their purchases (counted in revenue) */
  funcionarios: FuncionarioResumo[];
  /** What the spreadsheet holds, tab by tab and for the whole window */
  conferencia: Conferencia;
  /** Combined validation report from all layers */
  validationReport: ValidationReport;
}

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

/** Tolerance for math checks (in currency units) */
const MATH_TOLERANCE = 0.01;
/**
 * The "Preço médio" column is an average printed with a few decimals: qty ×
 * printed price can be off by up to half a cent per unit without anything being
 * wrong with the sale. The check tolerates exactly that rounding, nothing more.
 */
const ARREDONDAMENTO_PRECO_POR_UNIDADE = 0.005;
/** "Totais por unidade:" — the block of per-unit totals that closes the tab. */
const RE_TOTAIS = /^totais?\b/;

/**
 * Lines that close the current client block: whatever comes after them
 * ("Totais por unidade" and its per-unit rows) is no longer that client's
 * products.
 */
const RE_FIM_DE_BLOCO =
  /^(totais?\b|total\b|sub-?\s?total|soma\b|somatorio|qtde?\.?\s*total|quantidade total|valor total geral)/;

/**
 * Page furniture. Skipped without closing the block — a paginated export prints
 * these in the middle of a client and then goes on with that client's products.
 */
const RE_LINHA_IGNORADA =
  /^(gerado em|emitido|impresso|pagina\b|página\b|folha\b|filtros?\b|relatorio|relatório|periodo\b|período\b|vendas de produto)/;

// ---------------------------------------------------------------------------
// Row helpers
// ---------------------------------------------------------------------------

/**
 * Identity of a client across months: the report's own client code when there is
 * one, the normalised name otherwise. Keeping the raw header text out of this is
 * what stops "BRASA - Data da última compra: 08/09/2026" from becoming a
 * different client every month.
 */
function chaveDoCliente(codigo: string, nome: string): string {
  const cod = codigo.trim();
  if (cod) return `cod:${cod}`;
  return `nome:${normalizarTexto(nome)}`;
}

/**
 * Identity of a product: its code plus its description, normalised. Spacing,
 * casing and accents in the description no longer split one product in two,
 * and two products sharing a description but not a code stay apart.
 */
export function chaveDoProduto(codigo: string, descricao: string): string {
  return `${codigo.trim()}|${normalizarTexto(descricao)}`;
}

/** Cell at `col`, or "" when the column is absent from this sheet. */
function celula(row: SheetRows[number], col: number): unknown {
  return col === COLUNA_AUSENTE ? null : row[col];
}

/** Trimmed text of the cell at `col`. */
function texto(row: SheetRows[number], col: number): string {
  return String(celula(row, col) ?? "").trim();
}

/** A label that is a total, a subtotal or page furniture — never a product. */
function ehRotulo(valor: string): boolean {
  const t = normalizarTexto(valor);
  return RE_FIM_DE_BLOCO.test(t) || RE_LINHA_IGNORADA.test(t);
}

/** First cell of the row that holds text — used to spot summary lines. */
function primeiroTexto(row: SheetRows[number]): string {
  for (const cell of row) {
    const texto = String(cell ?? "").trim();
    if (texto && !ehNumero(cell)) return normalizarTexto(texto);
  }
  return "";
}

// ---------------------------------------------------------------------------
// Period Parser with full validation
// ---------------------------------------------------------------------------

/** One block of a client: the values and units of its sale rows so far. */
interface BlocoCliente {
  valores: number[];
  qtds: number[];
  linha: number;
}

/**
 * A line that sums the client instead of selling a product:
 *   - it names the client itself (its code, or its name in the description), or
 *   - it has no product code and repeats exactly the units and the value of the
 *     sales already listed in the block (two or more of them).
 * Such a line is skipped, so the client's units and value are counted once.
 */
function ehSomaDoCliente(
  cliente: ClientData,
  blocos: BlocoCliente[],
  linha: {
    codigo: string;
    descricao: string;
    codigoValido: boolean;
    qtd: number | null;
    venda: number | null;
  },
): boolean {
  const bloco = blocos[blocos.length - 1];
  if (!bloco || bloco.valores.length === 0) return false;
  if (linha.qtd === null && linha.venda === null) return false;

  const nomeDoCliente =
    linha.descricao !== "" && normalizarTexto(linha.descricao) === normalizarTexto(cliente.nome);
  const codigoDoCliente = linha.codigo !== "" && linha.codigo === cliente.codigoCliente && !linha.descricao;
  if (nomeDoCliente || codigoDoCliente) return true;

  if (linha.codigoValido || bloco.valores.length < 2) return false;
  const mesmaQtd = linha.qtd === null || Math.abs(linha.qtd - somarValores(bloco.qtds)) < 0.0001;
  const mesmoValor = linha.venda === null || Math.abs(linha.venda - somarValores(bloco.valores)) < 0.005;
  return mesmaQtd && mesmoValor;
}

/**
 * Reads one report tab into a period.
 *
 * The tab is read through the layout discovered for it, not through fixed column
 * indices, so an export that indents the report or renames a column still gets
 * read instead of coming back empty.
 *
 * Every row that carries a value is either a sale row of a client, a total
 * (of the client, of a unit or of the tab — never counted), or a discarded row
 * that is reported. Nothing is dropped silently.
 *
 * The revenue of a client is the sum of its sale rows ("Venda bruta"); the
 * units sold are the sum of the "Qtd. venda" column of those same rows.
 */
function parsePeriod(aba: AbaDescoberta): PeriodData {
  const { rows, nome } = aba;

  const validator = new ExcelValidator(nome);
  const clientes = new Map<string, ClientData>();
  let currentClient: ClientData | null = null;
  /**
   * Blocks of each client — a client listed on two report pages has two blocks.
   * Each keeps the values and units of its sale rows, so a line that repeats
   * their sum is recognised as the client's total and not counted again.
   */
  const blocos = new Map<ClientData, BlocoCliente[]>();
  /** Refreshed every time a product header row shows up in the tab. */
  let colunas: ColumnMap = aba.colunas;
  let produtosLidos = 0;
  /** Inside the "Totais por unidade" block that closes the report */
  let emTotais = false;
  const linhasDescartadas: LeituraAba["linhasDescartadas"] = [];

  // Track client entries for duplicate detection
  const clientEntries: { key: string; label: string; row: number }[] = [];

  const descartar = (excelRow: number, motivo: string, valor: number) => {
    linhasDescartadas.push({ linha: excelRow, motivo, valor });
    validator.checkCrossField(
      false,
      excelRow,
      colunas.vendaBruta,
      "linha",
      "Linha com valor reconhecida como venda ou total",
      `Linha ${excelRow} da aba ${nome}: ${motivo} — valor ${valor} fora do faturamento.`,
    );
  };

  for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
    const row = rows[rowIdx];
    const excelRow = rowIdx + 1; // 1-based for Excel reference

    if (!row || row.length === 0) continue;

    // ---- Product header: re-reads the column layout for the rows below --------
    const cabecalho = lerCabecalhoProdutos(row);
    if (cabecalho) {
      colunas = cabecalho;
      emTotais = false;
      continue;
    }

    // ---- Client header (any column, any casing, any dash) --------------------
    const cliente = lerCabecalhoCliente(row);
    if (cliente) {
      emTotais = false;
      const { codigo: codigoCliente, nome: nomeCliente, coluna } = cliente;
      const chave = chaveDoCliente(codigoCliente, nomeCliente);

      // Layer 1: Existence — the name must exist. A client printed without a
      // code is still a valid identification (by name), so it is only a warning.
      validator.checkCrossField(
        codigoCliente !== "",
        excelRow,
        coluna,
        "codigoCliente",
        "Cliente identificado por código",
        `Cliente "${nomeCliente}" sem código — identificado pelo nome em todos os meses.`,
      );
      validator.checkExistence(nomeCliente, excelRow, coluna, "nomeCliente");

      // Layer 3: Integrity — client name should not be too short
      if (nomeCliente) {
        validator.checkStringIntegrity(nomeCliente, excelRow, coluna, "nomeCliente", {
          minLength: 2,
        });
      }

      clientEntries.push({ key: chave, label: nomeCliente, row: excelRow });

      // The same client can be listed twice in one tab (two report pages):
      // keep the first block and append to it instead of dropping the earlier rows.
      const existente = clientes.get(chave);
      if (existente) {
        currentClient = existente;
        blocos.get(existente)!.push({ valores: [], qtds: [], linha: excelRow });
        // A later block may carry fields the first one did not print.
        existente.campos = { ...cliente.campos, ...existente.campos };
        existente.dataUltimaCompra ??= cliente.dataUltimaCompra;
        existente.fone ??= cliente.fone;
        existente.celular ??= cliente.celular;
        existente.funcionario ||= cliente.funcionario;
      } else {
        currentClient = {
          chave,
          nome: nomeCliente,
          codigoCliente,
          dataUltimaCompra: cliente.dataUltimaCompra,
          fone: cliente.fone,
          celular: cliente.celular,
          campos: cliente.campos,
          produtos: [],
          totalGeral: 0,
          totalRecalculado: 0,
          totalConsistente: true,
          origemLinha: excelRow,
          funcionario: cliente.funcionario,
        };
        clientes.set(chave, currentClient);
        blocos.set(currentClient, [{ valores: [], qtds: [], linha: excelRow }]);
      }
      continue;
    }

    const codigo = texto(row, colunas.codigo);
    const descricao = texto(row, colunas.descricao);
    // Units sold on the row: the "Qtd. venda" column of the product.
    const qtdLida = lerNumero(celula(row, colunas.quantidade));
    const vendaLida = lerNumero(celula(row, colunas.vendaBruta));
    const temValor = qtdLida !== null || vendaLida !== null;

    // A sale row: under a client, with a product code and/or description that is
    // not a totals label, and a quantity or a value. Checked before the summary
    // rows so a product named "TOTAL FLEX 5L" (with its code) is still a sale.
    const codigoValido = codigo !== "" && !ehRotulo(codigo) && !codigo.includes(":");
    const descricaoValida =
      descricao !== "" && !ehNumero(descricao) && (codigoValido || !ehRotulo(descricao));
    // "Total do cliente:" printed in any other column marks a totals line.
    const rotuloNaLinha = row.find(
      (c, i) => i !== colunas.descricao && typeof c === "string" && ehRotulo(c),
    );
    const somaDoCliente =
      !!currentClient &&
      ehSomaDoCliente(currentClient, blocos.get(currentClient)!, {
        codigo,
        descricao,
        codigoValido,
        qtd: qtdLida,
        venda: vendaLida,
      });
    const ehProduto =
      !!currentClient &&
      !emTotais &&
      temValor &&
      (codigoValido || descricaoValida) &&
      rotuloNaLinha === undefined &&
      !somaDoCliente;

    // ---- Summary / footer rows ----------------------------------------------
    // Nothing here is a sale: totals and client sums are skipped, never added
    // to the revenue or to the units sold.
    if (!ehProduto) {
      const marcador = rotuloNaLinha !== undefined ? normalizarTexto(rotuloNaLinha) : primeiroTexto(row);

      if (RE_FIM_DE_BLOCO.test(marcador)) {
        // "Totais por unidade:" — the per-unit rows that follow are totals too
        if (RE_TOTAIS.test(marcador)) emTotais = true;
        // The totals end the client: what follows is not that client's products.
        currentClient = null;
        continue;
      }

      if (RE_LINHA_IGNORADA.test(marcador)) {
        // Page furniture, skipped without closing the block (paginated exports).
        continue;
      }

      // "UN | | | 120 | | 4.500,00" — one unit's total
      if (emTotais) continue;

      // Sum of the client's rows (no product, or the client itself, or the very
      // sum of the block): already counted in the sales above it.
      if (currentClient && temValor && ((codigo === "" && descricao === "") || somaDoCliente)) {
        continue;
      }

      // A value that is neither a sale nor a total: never dropped quietly.
      if (vendaLida !== null && vendaLida !== 0) {
        descartar(
          excelRow,
          currentClient
            ? "linha com valor que não é produto nem somatório do cliente"
            : "linha com valor fora de qualquer bloco de cliente",
          vendaLida,
        );
      }
      continue;
    }

    // ---- Sale row -------------------------------------------------------------
    const client = currentClient!;

    // A sale row missing its code or its description still is a sale: it counts
    // in the revenue and is flagged so the source can be fixed.
    if (!codigoValido || !descricaoValida) {
      validator.checkCrossField(
        false,
        excelRow,
        codigoValido ? colunas.descricao : colunas.codigo,
        codigoValido ? "descricao" : "codigo",
        "Linha de venda com código e descrição do produto",
        `Linha ${excelRow} da aba ${nome}: venda ${codigoValido ? `do código ${codigo} sem descrição` : `de "${descricao}" sem código`} — ` +
          `contada no faturamento; confira a linha na planilha.`,
      );
    }
    const codigoProduto = codigoValido ? codigo : "";
    const descricaoProduto = descricaoValida ? descricao : `(sem descrição) ${codigo}`;

    // Layer 4: Structure — the row must reach the value column
    if (
      !validator.checkStructure(
        row,
        excelRow,
        Math.max(colunas.quantidade, colunas.vendaBruta) + 1,
        `Produto ${codigoProduto || descricaoProduto}`,
      )
    ) {
      continue;
    }

    // Layer 5: Traceability — trace every value to its cell
    const tCodigo = validator.traceValue(codigoProduto, excelRow, colunas.codigo);
    const tDescricao = validator.traceValue(descricaoProduto, excelRow, colunas.descricao);
    const tUnidade = validator.traceValue(
      texto(row, colunas.unidade),
      excelRow,
      colunas.unidade === COLUNA_AUSENTE ? colunas.descricao : colunas.unidade,
    );

    // Layer 1: Existence — mandatory fields
    validator.checkExistence(celula(row, colunas.quantidade), excelRow, colunas.quantidade, "quantidade");
    validator.checkExistence(celula(row, colunas.vendaBruta), excelRow, colunas.vendaBruta, "vendaBruta");
    if (colunas.precoMedio !== COLUNA_AUSENTE) {
      validator.checkExistence(celula(row, colunas.precoMedio), excelRow, colunas.precoMedio, "precoMedio");
    }

    // Layer 2: Type — numeric fields must be numbers (Brazilian format included)
    validator.checkType(celula(row, colunas.quantidade), "number", excelRow, colunas.quantidade, "quantidade");
    validator.checkType(celula(row, colunas.vendaBruta), "number", excelRow, colunas.vendaBruta, "vendaBruta");
    validator.checkType(descricaoProduto, "string", excelRow, colunas.descricao, "descricao");
    if (colunas.precoMedio !== COLUNA_AUSENTE) {
      validator.checkType(celula(row, colunas.precoMedio), "number", excelRow, colunas.precoMedio, "precoMedio");
    }

    // Units sold: the "Qtd. venda" of the product row.
    const quantidade = qtdLida ?? 0;
    const quantidadePrincipal = quantidade;
    const vendaBruta = vendaLida ?? 0;
    // Without a unit-price column the average price is derived, never invented.
    const precoMedio =
      colunas.precoMedio !== COLUNA_AUSENTE
        ? paraNumero(celula(row, colunas.precoMedio))
        : quantidadePrincipal !== 0
          ? preciseRound(vendaBruta / quantidadePrincipal, 4)
          : 0;

    // Traceability for numeric values
    const tQuantidade = validator.traceValue(quantidade, excelRow, colunas.quantidade);
    const tPrecoMedio = validator.traceValue(
      precoMedio,
      excelRow,
      colunas.precoMedio === COLUNA_AUSENTE ? colunas.vendaBruta : colunas.precoMedio,
    );
    const tVendaBruta = validator.traceValue(vendaBruta, excelRow, colunas.vendaBruta);

    // Layer 3: Integrity — numeric ranges
    validator.checkNumericIntegrity(quantidade, excelRow, colunas.quantidade, "quantidade");
    validator.checkNumericIntegrity(precoMedio, excelRow, tPrecoMedio.origin.col, "precoMedio");
    validator.checkNumericIntegrity(vendaBruta, excelRow, colunas.vendaBruta, "vendaBruta");

    // Cross-field: if qty > 0 then price should > 0 and vice versa
    validator.checkCrossField(
      !(quantidadePrincipal > 0 && precoMedio === 0),
      excelRow,
      tPrecoMedio.origin.col,
      "precoMedio",
      "Se quantidade > 0, preco medio deve ser > 0",
      `Produto ${codigoProduto || descricaoProduto}: Quantidade ${quantidade} com preco medio zero — possivel dado faltante.`,
    );

    // Math consistency: qty × price = total, within the rounding of the
    // printed average price.
    const mathCheck = validator.checkMathConsistency(
      quantidadePrincipal,
      precoMedio,
      vendaBruta,
      excelRow,
      `${client.nome} → Produto ${codigoProduto || descricaoProduto}`,
      MATH_TOLERANCE + Math.abs(quantidadePrincipal) * ARREDONDAMENTO_PRECO_POR_UNIDADE,
      colunas.vendaBruta,
    );

    client.produtos.push({
      produto: tDescricao.value,
      codigo: tCodigo.value,
      unidade: tUnidade.value,
      valorUnitario: precoMedio,
      quantidade,
      valorTotal: vendaBruta,
      validado: mathCheck.valid,
      valorCalculado: mathCheck.calculated,
      divergencia: mathCheck.difference,
      origens: {
        codigo: tCodigo.origin,
        produto: tDescricao.origin,
        unidade: tUnidade.origin,
        quantidade: tQuantidade.origin,
        valorUnitario: tPrecoMedio.origin,
        valorTotal: tVendaBruta.origin,
      },
    });
    const bloco = blocos.get(client)!;
    bloco[bloco.length - 1].valores.push(vendaBruta);
    bloco[bloco.length - 1].qtds.push(quantidade);
    produtosLidos++;
  }

  // ---- Duplicate detection for clients ---------------------------------------
  validator.checkDuplicates(clientEntries, `Periodo ${nome}`);

  // ---- Revenue per client: the sum of its sale rows ----------------------------
  for (const c of clientes.values()) {
    const productValues = c.produtos.map((p) => p.valorTotal);
    const sumProdutos = somarValores(productValues);
    c.totalRecalculado = sumProdutos;
    c.totalGeral = sumProdutos;

    // Duplicate product detection within same client
    const prodEntries = c.produtos.map((p) => ({
      key: chaveDoProduto(p.codigo, p.produto),
      label: `${p.produto} (${p.codigo})`,
      row: p.origens.codigo.row,
    }));
    validator.checkDuplicates(prodEntries, `${nome} → ${c.nome}`);
  }

  const todos = [...clientes.values()];
  const clientesCarteira = new Map(todos.filter((c) => !c.funcionario).map((c) => [c.chave, c]));
  const funcionarios = new Map(todos.filter((c) => c.funcionario).map((c) => [c.chave, c]));

  const totalFaturamento = somarValores(todos.map((c) => c.totalGeral));
  const faturamentoCarteira = somarValores([...clientesCarteira.values()].map((c) => c.totalGeral));
  const faturamentoFuncionarios = somarValores([...funcionarios.values()].map((c) => c.totalGeral));

  const vendas = todos.flatMap((c) => c.produtos);

  return {
    nome,
    sheetName: aba.sheetName,
    clientes,
    clientesCarteira,
    funcionarios,
    totalFaturamento,
    faturamentoCarteira,
    faturamentoFuncionarios,
    leituraAba: {
      registrosVenda: produtosLidos,
      quantidadeTotal: somarValores(vendas.map((p) => p.quantidade)),
      produtosDistintos: new Set(vendas.map((p) => chaveDoProduto(p.codigo, p.produto))).size,
      faturamentoRegistros: somarValores(vendas.map((p) => p.valorTotal)),
      linhasDescartadas,
    },
    validationReport: validator.getReport(),
    origemPeriodo: aba.origemPeriodo,
    mes: aba.mes,
    ano: aba.ano,
    oculta: aba.oculta,
    linhasLidas: aba.linhas,
    produtosLidos,
  };
}

// ---------------------------------------------------------------------------
// Chronological ordering of period sheets
// ---------------------------------------------------------------------------

/**
 * Orders period tabs chronologically so the analysis window does not depend on
 * the order the tabs happen to be saved in. Tab names that don't name a
 * recognizable month keep their slot — one unrecognised tab must not stop the
 * months around it from being ordered.
 *
 * Kept as a thin wrapper over the shared discovery helpers: the ordering rule
 * lives in one place and is the same one `analyzeWorkbook` applies.
 */
export function ordenarPeriodosCronologicamente(names: string[]): string[] {
  const itens = names.map((name, index) => ({ name, index, ...lerPeriodoDoNome(name) }));
  return ordenarPorPeriodo(itens).map((p) => p.name);
}

// ---------------------------------------------------------------------------
// Portfolio classifier over the whole window
// ---------------------------------------------------------------------------

/**
 * Reads a series of monthly values as a direction. A 5% dead band keeps ordinary
 * month-to-month noise from being reported as growth or decline; a series that
 * starts or ends at zero is a start or a stop, which is a direction either way.
 */
function tendenciaDaSerie(valores: number[]): { tendencia: Tendencia; variacao: number | null } {
  // A single month is a snapshot, not a direction: there is nothing to compare
  // it against, so it gets no trend rather than a fabricated "stable".
  if (valores.length < 2) return { tendencia: "estavel", variacao: null };

  const primeiro = valores[0] ?? 0;
  const ultimo = valores[valores.length - 1] ?? 0;

  if (primeiro === 0 && ultimo === 0) return { tendencia: "estavel", variacao: null };
  if (primeiro === 0) return { tendencia: "crescimento", variacao: null };
  if (ultimo === 0) return { tendencia: "queda", variacao: -100 };

  const v = preciseRound(((ultimo - primeiro) / primeiro) * 100, 4);
  if (Math.abs(v) < 5) return { tendencia: "estavel", variacao: v };
  return { tendencia: v > 0 ? "crescimento" : "queda", variacao: v };
}

function classificarCarteira(periodos: PeriodData[]): ClienteCarteira[] {
  const n = periodos.length;
  // Keyed by client identity, never by the raw header text.
  const allClients = new Map<string, { indicesPresente: number[] }>();

  for (let i = 0; i < n; i++) {
    // Employees ("*") are buyers, not portfolio clients.
    for (const chave of periodos[i].clientesCarteira.keys()) {
      if (!allClients.has(chave)) allClients.set(chave, { indicesPresente: [] });
      allClients.get(chave)!.indicesPresente.push(i);
    }
  }

  const result: ClienteCarteira[] = [];
  for (const [chave, data] of allClients) {
    const lastIdx = Math.max(...data.indicesPresente);
    // 0 = current period, 1 = one period back, etc.
    const distanceFromEnd = n - 1 - lastIdx;

    let status: StatusCarteira;
    if (distanceFromEnd === 0) {
      status = "ativo";
    } else if (distanceFromEnd === 1) {
      status = "frequencia_reduzida";
    } else if (distanceFromEnd <= 3) {
      status = "risco_abandono";
    } else {
      status = "inativo";
    }

    // Revenue in every month of the window, so the client can be read as a
    // series instead of as a single snapshot.
    const faturamentoPorPeriodo: Record<string, number> = {};
    for (const p of periodos) {
      faturamentoPorPeriodo[p.nome] = p.clientesCarteira.get(chave)?.totalGeral ?? 0;
    }

    // Measured between the client's own first and last purchase: a client that
    // arrived in month 3 is not "growing" just because months 1-2 were zero.
    const firstIdx = Math.min(...data.indicesPresente);
    const serie = periodos
      .slice(firstIdx, lastIdx + 1)
      .map((p) => p.clientesCarteira.get(chave)?.totalGeral ?? 0);
    const { tendencia, variacao: variacaoPeriodo } = tendenciaDaSerie(serie);

    // Name and contact details as printed in the last month the client bought.
    const ultimo = periodos[lastIdx].clientesCarteira.get(chave)!;

    result.push({
      chave,
      nome: ultimo.nome,
      codigoCliente: ultimo.codigoCliente || undefined,
      dataUltimaCompra: ultimo.dataUltimaCompra,
      fone: ultimo.fone,
      celular: ultimo.celular,
      status,
      ultimoPeriodo: periodos[lastIdx].nome,
      totalPeriodosPresente: data.indicesPresente.length,
      periodosPresente: data.indicesPresente.map((i) => periodos[i].nome),
      faturamentoUltimoPeriodo: ultimo.totalGeral,
      faturamentoPorPeriodo,
      variacaoPeriodo,
      tendencia,
    });
  }

  return result.sort((a, b) => a.nome.localeCompare(b.nome, "pt-BR"));
}

/** % change, or null when there is no base to compare against. */
function variacao(de: number, para: number): number | null {
  if (de === 0) return null;
  return preciseRound(((para - de) / de) * 100, 4);
}

/**
 * Revenue per portfolio client. Employees are neither in the count nor in the
 * revenue it divides, so their purchases do not inflate the client ticket.
 */
function ticketMedio(periodo: PeriodData): number {
  const n = periodo.clientesCarteira.size;
  return n > 0 ? preciseRound(periodo.faturamentoCarteira / n, 4) : 0;
}

/** Turns a client of one month into a movement entry, keeping every field. */
function movimento(cliente: ClientData): ClienteMovimento {
  return {
    chave: cliente.chave,
    nome: cliente.nome,
    codigoCliente: cliente.codigoCliente || undefined,
    dataUltimaCompra: cliente.dataUltimaCompra,
    fone: cliente.fone,
    celular: cliente.celular,
    faturamento: preciseRound(cliente.totalGeral, 4),
  };
}

function grupo(clientes: ClienteMovimento[]): GrupoClientes {
  const ordenados = [...clientes].sort((a, b) => a.nome.localeCompare(b.nome, "pt-BR"));
  return {
    clientes: ordenados,
    total: ordenados.length,
    faturamento: preciseSum(ordenados.map((c) => c.faturamento)),
  };
}

/**
 * Compares each month with the one before it, over the whole window.
 *
 * The client groups are read against the full history, not just the pair: a
 * client that bought in m1, skipped m2 and came back in m3 is "retornado" in the
 * m2→m3 transition, never "novo" — which is only possible because every month of
 * the file is available here.
 */
function compararPeriodos(periodos: PeriodData[]): TransicaoPeriodo[] {
  const transicoes: TransicaoPeriodo[] = [];

  for (let i = 1; i < periodos.length; i++) {
    const de = periodos[i - 1];
    const para = periodos[i];

    // Everyone seen before the previous month — tells novo from retornado.
    const vistoAntesDeDe = new Set<string>();
    for (let j = 0; j < i - 1; j++) {
      for (const chave of periodos[j].clientesCarteira.keys()) vistoAntesDeDe.add(chave);
    }

    const novos: ClienteMovimento[] = [];
    const retornados: ClienteMovimento[] = [];
    const mantidos: ClienteMovimento[] = [];
    const perdidos: ClienteMovimento[] = [];

    let expansao = 0;
    let contracao = 0;

    // Client movement is read on portfolio clients only: an employee never
    // enters as "novo" nor leaves as "perdido".
    for (const [chave, cliente] of para.clientesCarteira) {
      const anterior = de.clientesCarteira.get(chave);
      if (anterior) {
        mantidos.push(movimento(cliente));
        const delta = cliente.totalGeral - anterior.totalGeral;
        if (delta >= 0) expansao += delta;
        else contracao += delta;
      } else if (vistoAntesDeDe.has(chave)) {
        retornados.push(movimento(cliente));
      } else {
        novos.push(movimento(cliente));
      }
    }

    // Measured in the month they last bought — the month they are leaving.
    for (const [chave, cliente] of de.clientesCarteira) {
      if (!para.clientesCarteira.has(chave)) perdidos.push(movimento(cliente));
    }

    const gNovos = grupo(novos);
    const gRetornados = grupo(retornados);
    const gMantidos = grupo(mantidos);
    const gPerdidos = grupo(perdidos);

    const diferenca = preciseRound(para.totalFaturamento - de.totalFaturamento, 4);
    const tmDe = ticketMedio(de);
    const tmPara = ticketMedio(para);

    transicoes.push({
      de: de.nome,
      para: para.nome,
      faturamentoDe: de.totalFaturamento,
      faturamentoPara: para.totalFaturamento,
      variacaoFaturamento: variacao(de.totalFaturamento, para.totalFaturamento),
      diferencaFaturamento: diferenca,
      clientesDe: de.clientesCarteira.size,
      clientesPara: para.clientesCarteira.size,
      variacaoClientes: variacao(de.clientesCarteira.size, para.clientesCarteira.size),
      ticketMedioDe: tmDe,
      ticketMedioPara: tmPara,
      variacaoTicketMedio: variacao(tmDe, tmPara),
      novos: gNovos,
      retornados: gRetornados,
      mantidos: gMantidos,
      perdidos: gPerdidos,
      ponte: {
        novos: gNovos.faturamento,
        retornados: gRetornados.faturamento,
        expansao: preciseRound(expansao, 4),
        contracao: preciseRound(contracao, 4),
        perdidos: preciseRound(-gPerdidos.faturamento, 4),
        funcionarios: preciseRound(para.faturamentoFuncionarios - de.faturamentoFuncionarios, 4),
        total: diferenca,
      },
      taxaRetencao:
        de.clientesCarteira.size > 0
          ? preciseRound((mantidos.length / de.clientesCarteira.size) * 100, 4)
          : null,
    });
  }

  return transicoes;
}

/** Reads the whole window as one movement: where it started, where it ended. */
function consolidarEvolucao(
  periodos: PeriodData[],
  transicoes: TransicaoPeriodo[],
  classificacao: ClienteCarteira[],
): EvolucaoConsolidada {
  const primeiro = periodos[0];
  const ultimo = periodos[periodos.length - 1];

  const faturamentos = periodos.map((p) => p.totalFaturamento);
  const faturamentoTotal = somarValores(faturamentos);

  const melhor = periodos.reduce((a, b) => (b.totalFaturamento > a.totalFaturamento ? b : a));
  const pior = periodos.reduce((a, b) => (b.totalFaturamento < a.totalFaturamento ? b : a));

  const variacoes = transicoes
    .map((t) => t.variacaoFaturamento)
    .filter((v): v is number => v !== null);
  const retencoes = transicoes
    .map((t) => t.taxaRetencao)
    .filter((v): v is number => v !== null);

  const variacaoTotal = variacao(primeiro.totalFaturamento, ultimo.totalFaturamento);

  // A 5-point dead band keeps small month-to-month noise from being read as a
  // trend in either direction.
  let tendencia: Tendencia = "estavel";
  if (variacaoTotal !== null && Math.abs(variacaoTotal) >= 5) {
    tendencia = variacaoTotal > 0 ? "crescimento" : "queda";
  }

  const tmInicial = ticketMedio(primeiro);
  const tmFinal = ticketMedio(ultimo);

  return {
    meses: periodos.length,
    primeiro: primeiro.nome,
    ultimo: ultimo.nome,
    faturamentoInicial: primeiro.totalFaturamento,
    faturamentoFinal: ultimo.totalFaturamento,
    faturamentoTotal,
    faturamentoMedio: preciseRound(faturamentoTotal / periodos.length, 4),
    variacaoTotal,
    variacaoMediaMensal: variacoes.length
      ? preciseRound(variacoes.reduce((a, b) => a + b, 0) / variacoes.length, 4)
      : null,
    clientesInicial: primeiro.clientesCarteira.size,
    clientesFinal: ultimo.clientesCarteira.size,
    variacaoClientes: variacao(primeiro.clientesCarteira.size, ultimo.clientesCarteira.size),
    ticketMedioInicial: tmInicial,
    ticketMedioFinal: tmFinal,
    variacaoTicketMedio: variacao(tmInicial, tmFinal),
    melhorMes: { nome: melhor.nome, faturamento: melhor.totalFaturamento },
    piorMes: { nome: pior.nome, faturamento: pior.totalFaturamento },
    mesesEmAlta: transicoes.filter((t) => t.diferencaFaturamento > 0).length,
    mesesEmQueda: transicoes.filter((t) => t.diferencaFaturamento < 0).length,
    tendencia,
    clientesDistintos: classificacao.length,
    clientesRecorrentes: classificacao.filter((c) => c.totalPeriodosPresente === periodos.length)
      .length,
    novosNoPeriodo: transicoes.reduce((s, t) => s + t.novos.total, 0),
    perdidosNoPeriodo: classificacao.filter(
      (c) => c.status === "risco_abandono" || c.status === "inativo",
    ).length,
    taxaRetencaoMedia: retencoes.length
      ? preciseRound(retencoes.reduce((a, b) => a + b, 0) / retencoes.length, 4)
      : null,
  };
}

/**
 * Summarises every month of the window and how each moved against the one
 * before, reading the movement straight off the transitions so the table and the
 * transition cards can never disagree.
 */
function resumirEvolucao(
  periodos: PeriodData[],
  transicoes: TransicaoPeriodo[],
): PeriodoResumo[] {
  return periodos.map((p, i) => {
    // transicoes[i - 1] is the move that ends in this month; the first month of
    // the window has nothing before it to compare against.
    const t = i > 0 ? transicoes[i - 1] : null;

    return {
      nome: p.nome,
      faturamento: p.totalFaturamento,
      clientes: p.clientesCarteira.size,
      ticketMedio: ticketMedio(p),
      variacaoFaturamento: t ? t.variacaoFaturamento : null,
      variacaoClientes: t ? t.variacaoClientes : null,
      entraram: t ? t.novos.total + t.retornados.total : null,
      sairam: t ? t.perdidos.total : null,
      novos: t ? t.novos.total : null,
      retornados: t ? t.retornados.total : null,
      mantidos: t ? t.mantidos.total : null,
      perdidos: t ? t.perdidos.total : null,
      faturamentoNovos: t ? t.novos.faturamento : null,
      faturamentoPerdidos: t ? t.perdidos.faturamento : null,
      taxaRetencao: t ? t.taxaRetencao : null,
      quantidade: p.leituraAba.quantidadeTotal,
      variacaoQuantidade:
        i > 0 ? variacao(periodos[i - 1].leituraAba.quantidadeTotal, p.leituraAba.quantidadeTotal) : null,
      movimentacao:
        i > 0 ? tendenciaDaSerie([periodos[i - 1].totalFaturamento, p.totalFaturamento]).tendencia : null,
      funcionarios: p.funcionarios.size,
      faturamentoFuncionarios: p.faturamentoFuncionarios,
    };
  });
}

interface ProdutoAgregado {
  qtd: number;
  valor: number;
  codigo: string;
  produto: string;
}

/**
 * Revenue/quantity per product, for one period. Every buyer counts here —
 * employees included: their purchases are product sales like any other.
 */
function agregarProdutos(periodo: PeriodData): Map<string, ProdutoAgregado> {
  const m = new Map<string, ProdutoAgregado>();
  for (const c of periodo.clientes.values()) {
    for (const p of c.produtos) {
      const chave = chaveDoProduto(p.codigo, p.produto);
      const e = m.get(chave) ?? { qtd: 0, valor: 0, codigo: p.codigo, produto: p.produto };
      e.qtd = somarValores([e.qtd, p.quantidade]);
      e.valor = somarValores([e.valor, p.valorTotal]);
      if (!e.codigo) e.codigo = p.codigo;
      m.set(chave, e);
    }
  }
  return m;
}

function resumirCarteira(
  periodos: PeriodData[],
  classificacao: ClienteCarteira[],
): CarteiraResumo {
  const atual = periodos[periodos.length - 1];
  const anterior = periodos[periodos.length - 2];
  const anteriores = periodos.slice(0, -1);

  const vistoAntes = new Set<string>();
  for (const p of anteriores) for (const nome of p.clientesCarteira.keys()) vistoAntes.add(nome);

  const ativos: ClientData[] = [];
  const novos: ClientData[] = [];
  const retornados: ClientData[] = [];

  for (const [nome, cliente] of atual.clientesCarteira) {
    ativos.push(cliente);
    if (!vistoAntes.has(nome)) {
      novos.push(cliente);
    } else if (!anterior.clientesCarteira.has(nome)) {
      // Skipped last month but bought before — a longer buying cycle, not a new client.
      retornados.push(cliente);
    }
  }

  const frequenciaReduzida = classificacao.filter((c) => c.status === "frequencia_reduzida");
  const perdidos = classificacao.filter(
    (c) => c.status === "risco_abandono" || c.status === "inativo",
  );

  return {
    periodos: periodos.map((p) => p.nome),
    ativos,
    novos,
    retornados,
    frequenciaReduzida,
    perdidos,
    faturamentoNovos: preciseRound(
      novos.reduce((s, c) => s + c.totalGeral, 0),
      4,
    ),
    faturamentoPerdidos: preciseRound(
      perdidos.reduce((s, c) => s + c.faturamentoUltimoPeriodo, 0),
      4,
    ),
  };
}

// ---------------------------------------------------------------------------
// Main analysis entry point
// ---------------------------------------------------------------------------

export function analyzeWorkbook(buffer: ArrayBuffer): AnalysisResult {
  const wb = lerPastaDeTrabalho(buffer);

  // Every tab that holds report data becomes a period — identified by content,
  // ordered by the month it covers, with no cap on how many months are read.
  const descoberta = descobrirAbas(wb);

  if (descoberta.abas.length < 2) {
    const encontradas = descoberta.todasAsAbas.length
      ? descoberta.todasAsAbas.join(", ")
      : "nenhuma";
    const motivos = descoberta.ignoradas.length
      ? ` Abas descartadas: ${descoberta.ignoradas
          .map((a) => `"${a.sheetName}" (${a.motivo})`)
          .join("; ")}.`
      : "";
    throw new Error(
      `A planilha precisa ter pelo menos 2 abas (períodos) com dados de relatório. ` +
        `Abas encontradas: ${encontradas}.${motivos}`,
    );
  }

  const periodos = descoberta.abas.map(parsePeriod);

  const leitura: LeituraPlanilha = {
    abasNoArquivo: descoberta.todasAsAbas,
    abasLidas: periodos.map((p) => p.sheetName),
    abasIgnoradas: descoberta.ignoradas,
  };

  const periodoAnterior = periodos[periodos.length - 2];
  const periodoAtual = periodos[periodos.length - 1];

  // ---- Client comparison -----------------------------------------------------
  // Portfolio clients only — employees are not new or lost clients.
  const nomesAnterior = new Set(periodoAnterior.clientesCarteira.keys());
  const nomesAtual = new Set(periodoAtual.clientesCarteira.keys());

  const clientesNovos: ClientData[] = [];
  const clientesPerdidos: ClientData[] = [];
  const clientesMantidos: string[] = [];

  for (const nome of nomesAtual) {
    if (!nomesAnterior.has(nome)) {
      clientesNovos.push(periodoAtual.clientesCarteira.get(nome)!);
    } else {
      clientesMantidos.push(nome);
    }
  }
  for (const nome of nomesAnterior) {
    if (!nomesAtual.has(nome)) {
      clientesPerdidos.push(periodoAnterior.clientesCarteira.get(nome)!);
    }
  }

  // ---- Revenue calculations with precision -----------------------------------
  const faturamentoNovos = preciseRound(
    clientesNovos.reduce((s, c) => s + c.totalGeral, 0),
    4,
  );
  const faturamentoPerdidos = preciseRound(
    clientesPerdidos.reduce((s, c) => s + c.totalGeral, 0),
    4,
  );

  const variacaoFaturamento =
    periodoAnterior.totalFaturamento > 0
      ? preciseRound(
          ((periodoAtual.totalFaturamento - periodoAnterior.totalFaturamento) /
            periodoAnterior.totalFaturamento) *
            100,
          4,
        )
      : 0;

  // ---- Product comparison ----------------------------------------------------
  // Aggregated once per period, so every month of the file feeds the product
  // analysis instead of only the two compared months.
  const produtosPorPeriodo = periodos.map(agregarProdutos);
  const vazio: ProdutoAgregado = { qtd: 0, valor: 0, codigo: "", produto: "" };

  const produtosAnterior = produtosPorPeriodo[produtosPorPeriodo.length - 2];
  const produtosAtual = produtosPorPeriodo[produtosPorPeriodo.length - 1];

  // The universe of products is the union of every period — a product sold in an
  // earlier month and not in the last two must still show up in the comparison.
  const allProducts = new Set<string>();
  for (const mapa of produtosPorPeriodo) {
    for (const nome of mapa.keys()) allProducts.add(nome);
  }

  const comparativoProdutos: ProductComparison[] = [];
  for (const prod of allProducts) {
    const ant = produtosAnterior.get(prod) ?? vazio;
    const atu = produtosAtual.get(prod) ?? vazio;

    const valorPorPeriodo: Record<string, number> = {};
    const qtdPorPeriodo: Record<string, number> = {};
    let codigo = atu.codigo || ant.codigo;
    // Shown with the description of the most recent month it was sold in.
    let produto = "";
    periodos.forEach((p, i) => {
      const e = produtosPorPeriodo[i].get(prod);
      valorPorPeriodo[p.nome] = e?.valor ?? 0;
      qtdPorPeriodo[p.nome] = e?.qtd ?? 0;
      if (!codigo && e?.codigo) codigo = e.codigo;
      if (e?.produto) produto = e.produto;
    });

    const serie = periodos.map((p) => valorPorPeriodo[p.nome]);
    const { tendencia, variacao: variacaoPeriodo } = tendenciaDaSerie(serie);

    comparativoProdutos.push({
      chave: prod,
      produto,
      codigo,
      mesesComVenda: serie.filter((v) => v !== 0).length,
      variacaoPeriodo,
      tendencia,
      valorTotalPeriodo: somarValores(serie),
      qtdAnterior: ant.qtd,
      qtdAtual: atu.qtd,
      variacaoQtd: ant.qtd > 0
        ? preciseRound(((atu.qtd - ant.qtd) / ant.qtd) * 100, 4)
        : atu.qtd > 0 ? 100 : 0,
      valorAnterior: ant.valor,
      valorAtual: atu.valor,
      variacaoValor: ant.valor > 0
        ? preciseRound(((atu.valor - ant.valor) / ant.valor) * 100, 4)
        : atu.valor > 0 ? 100 : 0,
      valorPorPeriodo,
      qtdPorPeriodo,
    });
  }
  comparativoProdutos.sort((a, b) => b.valorAtual - a.valorAtual);

  // ---- Build human-readable inconsistencies ----------------------------------
  // Every period of the file is audited — an inconsistency in an older month is
  // still an inconsistency in the numbers the analysis is built on.
  const inconsistencias: string[] = [];
  for (const period of periodos) {
    for (const c of period.clientes.values()) {
      for (const p of c.produtos) {
        if (!p.validado) {
          inconsistencias.push(
            `${period.nome} → ${c.nome} → ${p.produto}: ` +
            `Valor (${toDisplay(p.valorTotal).toFixed(2)}) ≠ ` +
            `Preço médio (${toDisplay(p.valorUnitario).toFixed(2)}) × Qtd (${p.quantidade}) = ` +
            `${toDisplay(p.valorCalculado).toFixed(2)} ` +
            `[Divergência: R$ ${p.divergencia.toFixed(4)}] ` +
            `[Origem: ${p.origens.valorTotal.cellRef}]`
          );
        }
      }
    }
    for (const d of period.leituraAba.linhasDescartadas) {
      inconsistencias.push(
        `${period.nome} → linha ${d.linha}: ${d.motivo} ` +
        `[Valor fora do faturamento: R$ ${d.valor.toFixed(2)}] [Origem: ${period.sheetName}]`,
      );
    }
  }

  // ---- Portfolio classification over the whole window --------------------------
  const classificacaoCarteira = classificarCarteira(periodos);
  const carteira = resumirCarteira(periodos, classificacaoCarteira);
  // Month against month first; the per-month summary and the window view are
  // both derived from it, so every number on screen tells the same story.
  const transicoes = compararPeriodos(periodos);
  const evolucaoPeriodos = resumirEvolucao(periodos, transicoes);
  const evolucaoConsolidada = consolidarEvolucao(periodos, transicoes, classificacaoCarteira);
  const funcionarios = resumirFuncionarios(periodos);

  // ---- Conferência: what the spreadsheet holds, tab by tab ---------------------
  // Read straight from the sale rows; nothing is compared with or corrected by
  // anything outside the spreadsheet.
  const conferencia = resumirPlanilha(periodos);

  // ---- Merge validation reports into one combined report ----------------------
  // Covers every period read from the file: a month left out of the validation
  // is a month nobody checked, which is how a whole tab used to go missing
  // without a single alert.
  const relatorios = periodos.map((p) => p.validationReport);
  const allDiagnostics = relatorios.flatMap((r) => r.diagnostics);
  const totalChecks = relatorios.reduce((s, r) => s + r.totalChecks, 0);
  const passedChecks = relatorios.reduce((s, r) => s + r.passedChecks, 0);

  // A report tab that produced no client at all is a reading failure, not an
  // empty month: surface it instead of letting the period show up as zero.
  for (const p of periodos) {
    if (p.clientes.size === 0) {
      allDiagnostics.push({
        layer: "structure",
        severity: "error",
        cellRef: `${p.nome}!A1`,
        field: `aba ${p.sheetName}`,
        actual: "nenhum cliente reconhecido",
        expected: 'linhas no formato "Cliente: CÓDIGO - NOME"',
        impact:
          `A aba "${p.sheetName}" foi carregada (${p.linhasLidas} linhas) mas nenhum cliente ` +
          `foi reconhecido — o período entra na análise zerado. Confira o layout dessa aba.`,
      });
    }
  }

  // The same month in two tabs would be counted twice in the window.
  for (const r of descoberta.periodosRepetidos) {
    allDiagnostics.push({
      layer: "structure",
      severity: "error",
      cellRef: `${r.abas[1]}!A1`,
      field: `mês ${r.periodo}`,
      actual: `abas ${r.abas.map((a) => `"${a}"`).join(", ")}`,
      expected: "uma aba por mês",
      impact:
        `O mês ${r.periodo} aparece em mais de uma aba (${r.abas.join(", ")}) — ` +
        `as vendas desse mês seriam contadas mais de uma vez na janela.`,
    });
  }

  const layerSummary: Record<string, { passed: number; failed: number }> = {};
  for (const relatorio of relatorios) {
    for (const [layer, counts] of Object.entries(relatorio.layerSummary)) {
      if (!layerSummary[layer]) layerSummary[layer] = { passed: 0, failed: 0 };
      layerSummary[layer].passed += counts.passed;
      layerSummary[layer].failed += counts.failed;
    }
  }

  const validationReport: ValidationReport = {
    valid: allDiagnostics.filter((d) => d.severity === "error").length === 0,
    diagnostics: allDiagnostics,
    totalChecks,
    passedChecks,
    failedChecks: totalChecks - passedChecks,
    layerSummary,
  };

  return {
    periodoAnterior,
    periodoAtual,
    periodos,
    leitura,
    clientesNovos,
    clientesPerdidos,
    clientesMantidos,
    faturamentoNovos,
    faturamentoPerdidos,
    variacaoFaturamento,
    comparativoProdutos,
    inconsistencias,
    classificacaoCarteira,
    carteira,
    evolucaoPeriodos,
    transicoes,
    evolucaoConsolidada,
    funcionarios,
    conferencia,
    validationReport,
  };
}

/** Employees found in any month, with their purchases per month. */
function resumirFuncionarios(periodos: PeriodData[]): FuncionarioResumo[] {
  const mapa = new Map<string, FuncionarioResumo>();
  for (const p of periodos) {
    for (const f of p.funcionarios.values()) {
      const e =
        mapa.get(f.chave) ??
        {
          chave: f.chave,
          nome: f.nome,
          codigoCliente: f.codigoCliente || undefined,
          faturamentoPorPeriodo: Object.fromEntries(periodos.map((q) => [q.nome, 0])),
          faturamentoTotal: 0,
          periodosPresente: [],
        };
      e.nome = f.nome;
      e.faturamentoPorPeriodo[p.nome] = f.totalGeral;
      e.faturamentoTotal = somarValores([e.faturamentoTotal, f.totalGeral]);
      e.periodosPresente.push(p.nome);
      mapa.set(f.chave, e);
    }
  }
  return [...mapa.values()].sort((a, b) => a.nome.localeCompare(b.nome, "pt-BR"));
}
