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
}

export interface PeriodData {
  /** Label used across the analysis (tab name, or the period read from content) */
  nome: string;
  /** Tab name exactly as it appears in the file */
  sheetName: string;
  /** Clients of the month, keyed by ClientData.chave */
  clientes: Map<string, ClientData>;
  totalFaturamento: number;
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
  /** Combined validation report from all layers */
  validationReport: ValidationReport;
}

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

/** Tolerance for math checks (in currency units) */
const MATH_TOLERANCE = 0.01;
/** Tolerance for aggregation checks */
const AGGREGATION_TOLERANCE = 0.05;

/** Product codes seen in the exports: "5001", "PA-001", "12.345", "AB/2" */
const RE_CODIGO_PRODUTO = /^[A-Za-z0-9][A-Za-z0-9._/-]*$/;

/**
 * Lines that close the current client block: whatever comes after them
 * ("Totais por unidade" and its per-unit rows) is no longer that client's
 * products.
 */
const RE_FIM_DE_BLOCO = /^(totais?\b|total geral|subtotal|quantidade total|valor total geral)/;

/**
 * Page furniture. Skipped without closing the block — a paginated export prints
 * these in the middle of a client and then goes on with that client's products.
 */
const RE_LINHA_IGNORADA =
  /^(gerado em|emitido|impresso|pagina\b|página\b|folha\b|filtros?\b|relatorio|relatório|periodo\b|período\b|vendas de produto)/;

// ---------------------------------------------------------------------------
// Row helpers
// ---------------------------------------------------------------------------

function ehNumero(value: unknown): boolean {
  if (typeof value === "number") return Number.isFinite(value);
  const texto = String(value ?? "").trim();
  if (!texto) return false;
  // "1.234,56" and "1234.56" both show up depending on the export
  return Number.isFinite(Number(texto.replace(/\./g, "").replace(",", ".")));
}

function paraNumero(value: unknown): number {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const texto = String(value ?? "").trim();
  if (!texto) return 0;
  const n = Number(texto.replace(/\s/g, "").replace(/\./g, "").replace(",", "."));
  return Number.isFinite(n) ? n : 0;
}

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

/** Cell at `col`, or "" when the column is absent from this sheet. */
function celula(row: SheetRows[number], col: number): unknown {
  return col === COLUNA_AUSENTE ? null : row[col];
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

/**
 * Reads one report tab into a period.
 *
 * The tab is read through the layout discovered for it, not through fixed column
 * indices, so an export that indents the report or renames a column still gets
 * read instead of coming back empty.
 */
function parsePeriod(aba: AbaDescoberta): PeriodData {
  const { rows, nome } = aba;

  const validator = new ExcelValidator(nome);
  const clientes = new Map<string, ClientData>();
  let currentClient: ClientData | null = null;
  /**
   * Reported subtotal of each block of a client — a client listed on two report
   * pages has two blocks, and its total is the sum of both, not the last one.
   */
  const subtotais = new Map<ClientData, number[]>();
  /** Refreshed every time a product header row shows up in the tab. */
  let colunas: ColumnMap = aba.colunas;
  let produtosLidos = 0;

  // Track client entries for duplicate detection
  const clientEntries: { key: string; label: string; row: number }[] = [];

  for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
    const row = rows[rowIdx];
    const excelRow = rowIdx + 1; // 1-based for Excel reference

    if (!row || row.length === 0) continue;

    // ---- Product header: re-reads the column layout for the rows below --------
    const cabecalho = lerCabecalhoProdutos(row);
    if (cabecalho) {
      colunas = cabecalho;
      continue;
    }

    // ---- Client header (any column, any casing, any dash) --------------------
    const cliente = lerCabecalhoCliente(row);
    if (cliente) {
      const { codigo: codigoCliente, nome: nomeCliente, coluna } = cliente;
      const chave = chaveDoCliente(codigoCliente, nomeCliente);

      // Layer 1: Existence — client code and name must exist
      validator.checkExistence(codigoCliente, excelRow, coluna, "codigoCliente");
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
        subtotais.get(existente)!.push(0);
        // A later block may carry fields the first one did not print.
        existente.campos = { ...cliente.campos, ...existente.campos };
        existente.dataUltimaCompra ??= cliente.dataUltimaCompra;
        existente.fone ??= cliente.fone;
        existente.celular ??= cliente.celular;
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
        };
        clientes.set(chave, currentClient);
        subtotais.set(currentClient, [0]);
      }
      continue;
    }

    const codigo = String(celula(row, colunas.codigo) ?? "").trim();
    const descricao = String(celula(row, colunas.descricao) ?? "").trim();
    const temNumero =
      ehNumero(celula(row, colunas.quantidade)) ||
      ehNumero(celula(row, colunas.vendaBruta)) ||
      ehNumero(celula(row, colunas.precoMedio));

    // A product row carries a code and a description under the current client.
    const ehProduto =
      !!currentClient && !!codigo && RE_CODIGO_PRODUTO.test(codigo) && descricao !== "";

    // ---- Summary / footer rows ----------------------------------------------
    // Checked after the product test so a product named "TOTAL FLEX 5L" is not
    // mistaken for a totals line.
    if (!ehProduto) {
      const marcador = primeiroTexto(row);
      if (RE_FIM_DE_BLOCO.test(marcador)) {
        // The totals block ends the client: its rows are not products.
        currentClient = null;
        continue;
      }
      if (RE_LINHA_IGNORADA.test(marcador)) continue;
    }

    // A coded line with values but no description cannot join the product
    // comparison — flag it instead of letting it in as a nameless product.
    if (currentClient && !ehProduto && codigo && RE_CODIGO_PRODUTO.test(codigo) && temNumero) {
      validator.checkCrossField(
        false,
        excelRow,
        colunas.descricao,
        "descricao",
        "Linha com código e valores precisa de descrição do produto",
        `Linha ${excelRow} da aba ${nome}: código ${codigo} com valores mas sem descrição — ` +
          `linha ignorada no comparativo de produtos.`,
      );
      continue;
    }

    // ---- Product row ---------------------------------------------------------
    if (ehProduto && currentClient) {
      // Layer 4: Structure — the row must reach the columns the layout declares
      const colunasNecessarias =
        Math.max(
          colunas.codigo,
          colunas.descricao,
          colunas.quantidade,
          colunas.vendaBruta,
          colunas.precoMedio,
          colunas.unidade,
        ) + 1;
      if (!validator.checkStructure(row, excelRow, colunasNecessarias, `Produto ${codigo}`)) {
        continue;
      }

      // Layer 5: Traceability — trace every value to its cell
      const tCodigo = validator.traceValue(codigo, excelRow, colunas.codigo);
      const tDescricao = validator.traceValue(descricao, excelRow, colunas.descricao);
      const tUnidade = validator.traceValue(
        String(celula(row, colunas.unidade) ?? ""),
        excelRow,
        colunas.unidade === COLUNA_AUSENTE ? colunas.descricao : colunas.unidade,
      );

      // Layer 1: Existence — mandatory fields
      validator.checkExistence(celula(row, colunas.descricao), excelRow, colunas.descricao, "descricao");
      validator.checkExistence(celula(row, colunas.quantidade), excelRow, colunas.quantidade, "quantidade");
      validator.checkExistence(celula(row, colunas.vendaBruta), excelRow, colunas.vendaBruta, "vendaBruta");
      if (colunas.precoMedio !== COLUNA_AUSENTE) {
        validator.checkExistence(celula(row, colunas.precoMedio), excelRow, colunas.precoMedio, "precoMedio");
      }

      // Layer 2: Type — numeric fields must be numbers
      validator.checkType(celula(row, colunas.quantidade), "number", excelRow, colunas.quantidade, "quantidade");
      validator.checkType(celula(row, colunas.vendaBruta), "number", excelRow, colunas.vendaBruta, "vendaBruta");
      validator.checkType(celula(row, colunas.descricao), "string", excelRow, colunas.descricao, "descricao");
      if (colunas.precoMedio !== COLUNA_AUSENTE) {
        validator.checkType(celula(row, colunas.precoMedio), "number", excelRow, colunas.precoMedio, "precoMedio");
      }

      const quantidade = paraNumero(celula(row, colunas.quantidade));
      const vendaBruta = paraNumero(celula(row, colunas.vendaBruta));
      // Without a unit-price column the average price is derived, never invented.
      const precoMedio =
        colunas.precoMedio !== COLUNA_AUSENTE
          ? paraNumero(celula(row, colunas.precoMedio))
          : quantidade > 0
            ? preciseRound(vendaBruta / quantidade, 4)
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
        !(quantidade > 0 && precoMedio === 0),
        excelRow,
        tPrecoMedio.origin.col,
        "precoMedio",
        "Se quantidade > 0, preco medio deve ser > 0",
        `Produto ${codigo}: Quantidade ${quantidade} com preco medio zero — possivel dado faltante.`,
      );

      // Math consistency: qty × price = total (with precision)
      const mathCheck = validator.checkMathConsistency(
        quantidade,
        precoMedio,
        vendaBruta,
        excelRow,
        `${currentClient.nome} → Produto ${codigo}`,
        MATH_TOLERANCE,
      );

      currentClient.produtos.push({
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
      produtosLidos++;
      continue;
    }

    // ---- Subtotal row --------------------------------------------------------
    if (
      !codigo &&
      currentClient &&
      ehNumero(celula(row, colunas.quantidade)) &&
      ehNumero(celula(row, colunas.vendaBruta))
    ) {
      const subtotal = paraNumero(celula(row, colunas.vendaBruta));
      const blocos = subtotais.get(currentClient)!;
      blocos[blocos.length - 1] = subtotal;

      // Layer 5: trace the subtotal
      validator.traceValue(subtotal, excelRow, colunas.vendaBruta);
      continue;
    }
  }

  // ---- Duplicate detection for clients ---------------------------------------
  validator.checkDuplicates(clientEntries, `Periodo ${nome}`);

  // ---- Aggregation integrity per client --------------------------------------
  let totalFaturamento = 0;
  for (const c of clientes.values()) {
    const productValues = c.produtos.map((p) => p.valorTotal);
    const sumProdutos = preciseSum(productValues);
    c.totalRecalculado = sumProdutos;
    c.totalGeral = preciseSum(subtotais.get(c) ?? []);

    if (c.totalGeral === 0) {
      c.totalGeral = sumProdutos;
    } else {
      c.totalConsistente = validator.checkAggregation(
        productValues,
        c.totalGeral,
        c.origemLinha,
        `${nome} → ${c.nome}`,
        AGGREGATION_TOLERANCE,
      );
    }

    // Duplicate product detection within same client
    const prodEntries = c.produtos.map((p) => ({
      key: p.codigo,
      label: `${p.produto} (${p.codigo})`,
      row: p.origens.codigo.row,
    }));
    validator.checkDuplicates(prodEntries, `${nome} → ${c.nome}`);

    totalFaturamento += c.totalGeral;
  }

  totalFaturamento = preciseRound(totalFaturamento, 4);

  return {
    nome,
    sheetName: aba.sheetName,
    clientes,
    totalFaturamento,
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
    for (const chave of periodos[i].clientes.keys()) {
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
      faturamentoPorPeriodo[p.nome] = p.clientes.get(chave)?.totalGeral ?? 0;
    }

    // Measured between the client's own first and last purchase: a client that
    // arrived in month 3 is not "growing" just because months 1-2 were zero.
    const firstIdx = Math.min(...data.indicesPresente);
    const serie = periodos
      .slice(firstIdx, lastIdx + 1)
      .map((p) => p.clientes.get(chave)?.totalGeral ?? 0);
    const { tendencia, variacao: variacaoPeriodo } = tendenciaDaSerie(serie);

    // Name and contact details as printed in the last month the client bought.
    const ultimo = periodos[lastIdx].clientes.get(chave)!;

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

function ticketMedio(periodo: PeriodData): number {
  const n = periodo.clientes.size;
  return n > 0 ? preciseRound(periodo.totalFaturamento / n, 4) : 0;
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
      for (const chave of periodos[j].clientes.keys()) vistoAntesDeDe.add(chave);
    }

    const novos: ClienteMovimento[] = [];
    const retornados: ClienteMovimento[] = [];
    const mantidos: ClienteMovimento[] = [];
    const perdidos: ClienteMovimento[] = [];

    let expansao = 0;
    let contracao = 0;

    for (const [chave, cliente] of para.clientes) {
      const anterior = de.clientes.get(chave);
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
    for (const [chave, cliente] of de.clientes) {
      if (!para.clientes.has(chave)) perdidos.push(movimento(cliente));
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
      clientesDe: de.clientes.size,
      clientesPara: para.clientes.size,
      variacaoClientes: variacao(de.clientes.size, para.clientes.size),
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
        total: diferenca,
      },
      taxaRetencao:
        de.clientes.size > 0
          ? preciseRound((mantidos.length / de.clientes.size) * 100, 4)
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
  const faturamentoTotal = preciseSum(faturamentos);

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
    clientesInicial: primeiro.clientes.size,
    clientesFinal: ultimo.clientes.size,
    variacaoClientes: variacao(primeiro.clientes.size, ultimo.clientes.size),
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
      clientes: p.clientes.size,
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
    };
  });
}

interface ProdutoAgregado {
  qtd: number;
  valor: number;
  codigo: string;
}

/** Revenue/quantity per product, for one period. */
function agregarProdutos(periodo: PeriodData): Map<string, ProdutoAgregado> {
  const m = new Map<string, ProdutoAgregado>();
  for (const c of periodo.clientes.values()) {
    for (const p of c.produtos) {
      const e = m.get(p.produto) ?? { qtd: 0, valor: 0, codigo: p.codigo };
      e.qtd = preciseRound(e.qtd + p.quantidade, 4);
      e.valor = preciseRound(e.valor + p.valorTotal, 4);
      if (!e.codigo) e.codigo = p.codigo;
      m.set(p.produto, e);
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
  for (const p of anteriores) for (const nome of p.clientes.keys()) vistoAntes.add(nome);

  const ativos: ClientData[] = [];
  const novos: ClientData[] = [];
  const retornados: ClientData[] = [];

  for (const [nome, cliente] of atual.clientes) {
    ativos.push(cliente);
    if (!vistoAntes.has(nome)) {
      novos.push(cliente);
    } else if (!anterior.clientes.has(nome)) {
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
  const nomesAnterior = new Set(periodoAnterior.clientes.keys());
  const nomesAtual = new Set(periodoAtual.clientes.keys());

  const clientesNovos: ClientData[] = [];
  const clientesPerdidos: ClientData[] = [];
  const clientesMantidos: string[] = [];

  for (const nome of nomesAtual) {
    if (!nomesAnterior.has(nome)) {
      clientesNovos.push(periodoAtual.clientes.get(nome)!);
    } else {
      clientesMantidos.push(nome);
    }
  }
  for (const nome of nomesAnterior) {
    if (!nomesAtual.has(nome)) {
      clientesPerdidos.push(periodoAnterior.clientes.get(nome)!);
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
  const vazio: ProdutoAgregado = { qtd: 0, valor: 0, codigo: "" };

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
    periodos.forEach((p, i) => {
      const e = produtosPorPeriodo[i].get(prod);
      valorPorPeriodo[p.nome] = e?.valor ?? 0;
      qtdPorPeriodo[p.nome] = e?.qtd ?? 0;
      if (!codigo && e?.codigo) codigo = e.codigo;
    });

    const serie = periodos.map((p) => valorPorPeriodo[p.nome]);
    const { tendencia, variacao: variacaoPeriodo } = tendenciaDaSerie(serie);

    comparativoProdutos.push({
      produto: prod,
      codigo,
      mesesComVenda: serie.filter((v) => v > 0).length,
      variacaoPeriodo,
      tendencia,
      valorTotalPeriodo: preciseSum(serie),
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
      if (!c.totalConsistente) {
        inconsistencias.push(
          `${period.nome} → ${c.nome}: ` +
          `Total reportado (${toDisplay(c.totalGeral).toFixed(2)}) ≠ ` +
          `Soma dos produtos (${toDisplay(c.totalRecalculado).toFixed(2)}) ` +
          `[Divergência: R$ ${Math.abs(c.totalGeral - c.totalRecalculado).toFixed(4)}]`
        );
      }
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

  // ---- Merge validation reports into one combined report ----------------------
  // Covers every period read from the file: a month left out of the validation
  // is a month nobody checked, which is how a whole tab used to go missing
  // without a single alert.
  const allDiagnostics = periodos.flatMap((p) => p.validationReport.diagnostics);
  const totalChecks = periodos.reduce((s, p) => s + p.validationReport.totalChecks, 0);
  const passedChecks = periodos.reduce((s, p) => s + p.validationReport.passedChecks, 0);

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

  const layerSummary: Record<string, { passed: number; failed: number }> = {};
  for (const period of periodos) {
    for (const [layer, counts] of Object.entries(period.validationReport.layerSummary)) {
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
    validationReport,
  };
}
