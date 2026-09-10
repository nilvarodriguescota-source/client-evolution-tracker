import * as XLSX from "xlsx";
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
  nome: string;
  codigoCliente: string;
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
  nome: string;
  clientes: Map<string, ClientData>;
  totalFaturamento: number;
  /** Validation report for this period */
  validationReport: ValidationReport;
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
}

export type StatusCarteira =
  | "ativo"
  | "frequencia_reduzida"
  | "risco_abandono"
  | "inativo";

export interface ClienteCarteira {
  nome: string;
  codigoCliente?: string;
  status: StatusCarteira;
  ultimoPeriodo: string;
  totalPeriodosPresente: number;
  periodosPresente: string[];
  faturamentoUltimoPeriodo: number;
}

/**
 * Client portfolio read over the whole analysis window (up to 6 months), as
 * opposed to the strict month-over-month comparison kept in AnalysisResult.
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

export interface AnalysisResult {
  periodoAnterior: PeriodData;
  periodoAtual: PeriodData;
  /** All parsed periods used for the 6-month portfolio analysis (up to 6, oldest first) */
  periodos: PeriodData[];
  clientesNovos: ClientData[];
  clientesPerdidos: ClientData[];
  clientesMantidos: string[];
  faturamentoNovos: number;
  faturamentoPerdidos: number;
  variacaoFaturamento: number;
  comparativoProdutos: ProductComparison[];
  inconsistencias: string[];
  /** 6-month portfolio classification for every client found in any period */
  classificacaoCarteira: ClienteCarteira[];
  /** Portfolio figures computed over the whole window (drives the KPI cards) */
  carteira: CarteiraResumo;
  /** Month-by-month comparison across the whole window, oldest first */
  evolucaoPeriodos: PeriodoResumo[];
  /** Combined validation report from all layers */
  validationReport: ValidationReport;
}

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

/** Minimum columns expected in a product row */
const MIN_PRODUCT_COLS = 6;

/** Column indices */
const COL = {
  CODIGO: 0,
  DESCRICAO: 1,
  UNIDADE: 2,
  QUANTIDADE: 3,
  PRECO_MEDIO: 4,
  VENDA_BRUTA: 5,
} as const;

/** Maximum number of monthly sheets considered by the portfolio analysis */
export const MAX_PERIODOS_CARTEIRA = 6;

/** Tolerance for math checks (in currency units) */
const MATH_TOLERANCE = 0.01;
/** Tolerance for aggregation checks */
const AGGREGATION_TOLERANCE = 0.05;

// ---------------------------------------------------------------------------
// Period Parser with full validation
// ---------------------------------------------------------------------------

function parsePeriod(sheet: XLSX.WorkSheet, nome: string): PeriodData {
  const rows: (string | number | null)[][] = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: null,
  });

  const validator = new ExcelValidator(nome);
  const clientes = new Map<string, ClientData>();
  let currentClient: ClientData | null = null;

  // Track client entries for duplicate detection
  const clientEntries: { key: string; label: string; row: number }[] = [];

  for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
    const row = rows[rowIdx];
    const excelRow = rowIdx + 1; // 1-based for Excel reference

    if (!row || row.length === 0) continue;

    const firstCell = String(row[0] ?? "").trim();

    // ---- Detect client header ------------------------------------------------
    if (firstCell.startsWith("Cliente:")) {
      const match = firstCell.match(/Cliente:\s*(\d+)\s*-\s*(.+?)(?:\s*-\s*Data|$)/);
      if (match) {
        const codigoCliente = match[1];
        const nomeCliente = match[2].trim();

        // Layer 1: Existence — client code and name must exist
        validator.checkExistence(codigoCliente, excelRow, COL.CODIGO, "codigoCliente");
        validator.checkExistence(nomeCliente, excelRow, COL.CODIGO, "nomeCliente");

        // Layer 3: Integrity — client name should not be too short
        if (nomeCliente) {
          validator.checkStringIntegrity(nomeCliente, excelRow, COL.CODIGO, "nomeCliente", { minLength: 2 });
        }

        clientEntries.push({ key: nomeCliente, label: nomeCliente, row: excelRow });

        currentClient = {
          nome: nomeCliente,
          codigoCliente,
          produtos: [],
          totalGeral: 0,
          totalRecalculado: 0,
          totalConsistente: true,
          origemLinha: excelRow,
        };
        clientes.set(nomeCliente, currentClient);
      }
      continue;
    }

    // Skip header / empty / summary rows
    if (
      firstCell === "Código" ||
      firstCell === "" ||
      firstCell === "Totais por unidade:" ||
      firstCell.startsWith("Quantidade total") ||
      firstCell.startsWith("Gerado em")
    ) {
      continue;
    }

    // ---- Product row ---------------------------------------------------------
    const codigo = firstCell;
    const isProductRow = /^\d+$/.test(codigo) && currentClient;

    if (isProductRow && currentClient) {
      // Layer 4: Structure — row must have enough columns
      if (!validator.checkStructure(row, excelRow, MIN_PRODUCT_COLS, `Produto ${codigo}`)) {
        continue;
      }

      // Layer 5: Traceability — trace every value to its cell
      const tCodigo = validator.traceValue(codigo, excelRow, COL.CODIGO);
      const tDescricao = validator.traceValue(String(row[COL.DESCRICAO] ?? ""), excelRow, COL.DESCRICAO);
      const tUnidade = validator.traceValue(String(row[COL.UNIDADE] ?? ""), excelRow, COL.UNIDADE);

      // Layer 1: Existence — mandatory fields
      validator.checkExistence(row[COL.DESCRICAO], excelRow, COL.DESCRICAO, "descricao");
      validator.checkExistence(row[COL.QUANTIDADE], excelRow, COL.QUANTIDADE, "quantidade");
      validator.checkExistence(row[COL.PRECO_MEDIO], excelRow, COL.PRECO_MEDIO, "precoMedio");
      validator.checkExistence(row[COL.VENDA_BRUTA], excelRow, COL.VENDA_BRUTA, "vendaBruta");

      // Layer 2: Type — numeric fields must be numbers
      validator.checkType(row[COL.QUANTIDADE], "number", excelRow, COL.QUANTIDADE, "quantidade");
      validator.checkType(row[COL.PRECO_MEDIO], "number", excelRow, COL.PRECO_MEDIO, "precoMedio");
      validator.checkType(row[COL.VENDA_BRUTA], "number", excelRow, COL.VENDA_BRUTA, "vendaBruta");
      validator.checkType(row[COL.DESCRICAO], "string", excelRow, COL.DESCRICAO, "descricao");

      const quantidade = Number(row[COL.QUANTIDADE]) || 0;
      const precoMedio = Number(row[COL.PRECO_MEDIO]) || 0;
      const vendaBruta = Number(row[COL.VENDA_BRUTA]) || 0;

      // Traceability for numeric values
      const tQuantidade = validator.traceValue(quantidade, excelRow, COL.QUANTIDADE);
      const tPrecoMedio = validator.traceValue(precoMedio, excelRow, COL.PRECO_MEDIO);
      const tVendaBruta = validator.traceValue(vendaBruta, excelRow, COL.VENDA_BRUTA);

      // Layer 3: Integrity — numeric ranges
      validator.checkNumericIntegrity(quantidade, excelRow, COL.QUANTIDADE, "quantidade");
      validator.checkNumericIntegrity(precoMedio, excelRow, COL.PRECO_MEDIO, "precoMedio");
      validator.checkNumericIntegrity(vendaBruta, excelRow, COL.VENDA_BRUTA, "vendaBruta");

      // Cross-field: if qty > 0 then price should > 0 and vice versa
      validator.checkCrossField(
        !(quantidade > 0 && precoMedio === 0),
        excelRow,
        COL.PRECO_MEDIO,
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
        codigo: tCodigo.value as string,
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
      continue;
    }

    // ---- Subtotal row --------------------------------------------------------
    if (!firstCell && currentClient && row[COL.QUANTIDADE] && row[COL.VENDA_BRUTA]) {
      currentClient.totalGeral = Number(row[COL.VENDA_BRUTA]) || 0;

      // Layer 5: trace the subtotal
      validator.traceValue(currentClient.totalGeral, excelRow, COL.VENDA_BRUTA);
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
    const prodEntries = c.produtos.map((p, i) => ({
      key: p.codigo,
      label: `${p.produto} (${p.codigo})`,
      row: p.origens.codigo.row,
    }));
    validator.checkDuplicates(prodEntries, `${nome} → ${c.nome}`);

    totalFaturamento += c.totalGeral;
  }

  totalFaturamento = preciseRound(totalFaturamento, 4);

  return { nome, clientes, totalFaturamento, validationReport: validator.getReport() };
}

// ---------------------------------------------------------------------------
// Chronological ordering of period sheets
// ---------------------------------------------------------------------------

const MESES: Record<string, number> = {
  janeiro: 1, jan: 1,
  fevereiro: 2, fev: 2,
  marco: 3, mar: 3,
  abril: 4, abr: 4,
  maio: 5, mai: 5,
  junho: 6, jun: 6,
  julho: 7, jul: 7,
  agosto: 8, ago: 8,
  setembro: 9, set: 9, sep: 9,
  outubro: 10, out: 10,
  novembro: 11, nov: 11,
  dezembro: 12, dez: 12,
};

interface SheetPeriod {
  name: string;
  index: number;
  month?: number;
  year?: number;
}

/** Extracts month/year from a sheet name ("junho", "junho/2025", "06/2025", "2025-06"). */
function parseSheetPeriod(name: string, index: number): SheetPeriod {
  const norm = name
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, ""); // "março" -> "marco"

  let m = norm.match(/^(\d{4})[-/.](\d{1,2})$/);
  if (m) return { name, index, year: Number(m[1]), month: Number(m[2]) };

  m = norm.match(/^(\d{1,2})[-/.](\d{4})$/);
  if (m) return { name, index, month: Number(m[1]), year: Number(m[2]) };

  const month = (norm.match(/[a-z]+/g) ?? [])
    .map((w) => MESES[w])
    .find((v) => v !== undefined);
  const yearMatch = norm.match(/\b(\d{4})\b/);

  return { name, index, month, year: yearMatch ? Number(yearMatch[1]) : undefined };
}

/**
 * Orders period sheets chronologically so the analysis window does not depend on
 * the order the tabs happen to be saved in. Sheet names that don't name a
 * recognizable month are left in workbook order (nothing reliable to sort by).
 */
export function ordenarPeriodosCronologicamente(names: string[]): string[] {
  const parsed = names.map(parseSheetPeriod);

  if (parsed.some((p) => p.month === undefined)) return names;

  // Explicit years on every sheet: a plain (year, month) sort is unambiguous.
  if (parsed.every((p) => p.year !== undefined)) {
    return [...parsed]
      .sort((a, b) => a.year! - b.year! || a.month! - b.month! || a.index - b.index)
      .map((p) => p.name);
  }

  // No years: pick the 12-month rotation that packs the months into the
  // shortest span, so novembro → dezembro → janeiro sorts across the year end
  // instead of collapsing to janeiro → novembro → dezembro.
  const meses = [...new Set(parsed.map((p) => p.month!))];
  let inicio = meses[0];
  let menorSpan = Infinity;
  for (const s of meses) {
    const span = Math.max(...meses.map((m) => (m - s + 12) % 12));
    if (span < menorSpan) {
      menorSpan = span;
      inicio = s;
    }
  }
  const rank = (m: number) => (m - inicio + 12) % 12;

  return [...parsed]
    .sort((a, b) => rank(a.month!) - rank(b.month!) || a.index - b.index)
    .map((p) => p.name);
}

// ---------------------------------------------------------------------------
// 6-month portfolio classifier
// ---------------------------------------------------------------------------

function classificarCarteira(periodos: PeriodData[]): ClienteCarteira[] {
  const n = periodos.length;
  const allClients = new Map<string, { indicesPresente: number[]; codigoCliente?: string }>();

  for (let i = 0; i < n; i++) {
    for (const [nome, client] of periodos[i].clientes) {
      if (!allClients.has(nome)) {
        allClients.set(nome, { indicesPresente: [], codigoCliente: client.codigoCliente });
      }
      allClients.get(nome)!.indicesPresente.push(i);
    }
  }

  const result: ClienteCarteira[] = [];
  for (const [nome, data] of allClients) {
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

    result.push({
      nome,
      codigoCliente: data.codigoCliente,
      status,
      ultimoPeriodo: periodos[lastIdx].nome,
      totalPeriodosPresente: data.indicesPresente.length,
      periodosPresente: data.indicesPresente.map((i) => periodos[i].nome),
      faturamentoUltimoPeriodo: periodos[lastIdx].clientes.get(nome)?.totalGeral ?? 0,
    });
  }

  return result.sort((a, b) => a.nome.localeCompare(b.nome, "pt-BR"));
}

/** Summarises every month of the window and how each moved against the one before. */
function resumirEvolucao(periodos: PeriodData[]): PeriodoResumo[] {
  return periodos.map((p, i) => {
    const ant = i > 0 ? periodos[i - 1] : null;
    const clientes = p.clientes.size;

    return {
      nome: p.nome,
      faturamento: p.totalFaturamento,
      clientes,
      ticketMedio: clientes > 0 ? preciseRound(p.totalFaturamento / clientes, 4) : 0,
      variacaoFaturamento:
        ant && ant.totalFaturamento > 0
          ? preciseRound(
              ((p.totalFaturamento - ant.totalFaturamento) / ant.totalFaturamento) * 100,
              4,
            )
          : null,
      variacaoClientes:
        ant && ant.clientes.size > 0
          ? preciseRound(((clientes - ant.clientes.size) / ant.clientes.size) * 100, 4)
          : null,
      entraram: ant
        ? [...p.clientes.keys()].filter((n) => !ant.clientes.has(n)).length
        : null,
      sairam: ant
        ? [...ant.clientes.keys()].filter((n) => !p.clientes.has(n)).length
        : null,
    };
  });
}

/** Revenue/quantity per product, for one period. */
function agregarProdutos(periodo: PeriodData): Map<string, { qtd: number; valor: number }> {
  const m = new Map<string, { qtd: number; valor: number }>();
  for (const c of periodo.clientes.values()) {
    for (const p of c.produtos) {
      const e = m.get(p.produto) ?? { qtd: 0, valor: 0 };
      e.qtd = preciseRound(e.qtd + p.quantidade, 4);
      e.valor = preciseRound(e.valor + p.valorTotal, 4);
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
  // NOTE: legacy .xls (CFB/BIFF8) files fail to parse when a raw ArrayBuffer is
  // passed — SheetJS returns a single bogus "Sheet1" instead of the real tabs.
  // Wrapping in a Uint8Array parses every sheet correctly for both .xls and .xlsx.
  const wb = XLSX.read(new Uint8Array(buffer), { type: "array" });

  if (wb.SheetNames.length < 2) {
    throw new Error(
      `A planilha precisa ter pelo menos 2 abas (períodos). ` +
        `Abas encontradas: ${wb.SheetNames.length ? wb.SheetNames.join(", ") : "nenhuma"}.`,
    );
  }

  // Order the tabs by month first, so the window is the 6 most recent months
  // rather than the last 6 tabs in whatever order they happen to be saved.
  const sheetNames = ordenarPeriodosCronologicamente(wb.SheetNames);
  const maxPeriods = Math.min(sheetNames.length, MAX_PERIODOS_CARTEIRA);
  const periodSheetNames = sheetNames.slice(sheetNames.length - maxPeriods);

  const periodos = periodSheetNames.map((name) =>
    parsePeriod(wb.Sheets[name], name),
  );

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

  // ---- Product comparison with precision -------------------------------------
  const produtosAnterior = new Map<string, { qtd: number; valor: number; codigo: string }>();
  const produtosAtual = new Map<string, { qtd: number; valor: number; codigo: string }>();

  for (const c of periodoAnterior.clientes.values()) {
    for (const p of c.produtos) {
      const key = p.produto;
      const existing = produtosAnterior.get(key) || { qtd: 0, valor: 0, codigo: p.codigo };
      existing.qtd = preciseRound(existing.qtd + p.quantidade, 4);
      existing.valor = preciseRound(existing.valor + p.valorTotal, 4);
      produtosAnterior.set(key, existing);
    }
  }
  for (const c of periodoAtual.clientes.values()) {
    for (const p of c.produtos) {
      const key = p.produto;
      const existing = produtosAtual.get(key) || { qtd: 0, valor: 0, codigo: p.codigo };
      existing.qtd = preciseRound(existing.qtd + p.quantidade, 4);
      existing.valor = preciseRound(existing.valor + p.valorTotal, 4);
      produtosAtual.set(key, existing);
    }
  }

  // Per-period product totals, so the summary can chart every month of the window.
  const produtosPorPeriodo = periodos.map(agregarProdutos);

  const allProducts = new Set([...produtosAnterior.keys(), ...produtosAtual.keys()]);
  const comparativoProdutos: ProductComparison[] = [];
  for (const prod of allProducts) {
    const ant = produtosAnterior.get(prod) || { qtd: 0, valor: 0, codigo: "" };
    const atu = produtosAtual.get(prod) || { qtd: 0, valor: 0, codigo: ant.codigo };

    const valorPorPeriodo: Record<string, number> = {};
    const qtdPorPeriodo: Record<string, number> = {};
    periodos.forEach((p, i) => {
      const e = produtosPorPeriodo[i].get(prod);
      valorPorPeriodo[p.nome] = e?.valor ?? 0;
      qtdPorPeriodo[p.nome] = e?.qtd ?? 0;
    });

    comparativoProdutos.push({
      produto: prod,
      codigo: atu.codigo || ant.codigo,
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
  const inconsistencias: string[] = [];
  for (const period of [periodoAnterior, periodoAtual]) {
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

  // ---- 6-month portfolio classification ---------------------------------------
  const classificacaoCarteira = classificarCarteira(periodos);
  const carteira = resumirCarteira(periodos, classificacaoCarteira);
  const evolucaoPeriodos = resumirEvolucao(periodos);

  // ---- Merge validation reports into one combined report ----------------------
  // Scoped to the two comparison periods only, so every already-validated number
  // in the Validação/Alertas tabs stays exactly as it was. The extra historical
  // periods feed the portfolio classification alone.
  const periodosValidacao = [periodoAnterior, periodoAtual];
  const allDiagnostics = periodosValidacao.flatMap((p) => p.validationReport.diagnostics);
  const totalChecks = periodosValidacao.reduce((s, p) => s + p.validationReport.totalChecks, 0);
  const passedChecks = periodosValidacao.reduce((s, p) => s + p.validationReport.passedChecks, 0);

  const layerSummary: Record<string, { passed: number; failed: number }> = {};
  for (const period of periodosValidacao) {
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
    validationReport,
  };
}
