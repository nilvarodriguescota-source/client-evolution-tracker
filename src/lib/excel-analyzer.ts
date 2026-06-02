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

// ---------------------------------------------------------------------------
// Main analysis entry point
// ---------------------------------------------------------------------------

export function analyzeWorkbook(buffer: ArrayBuffer): AnalysisResult {
  const wb = XLSX.read(buffer, { type: "array" });

  if (wb.SheetNames.length < 2) {
    throw new Error("A planilha precisa ter pelo menos 2 abas (períodos).");
  }

  const sheetNames = wb.SheetNames;
  // Use up to 6 most recent sheets for the portfolio analysis window
  const maxPeriods = Math.min(sheetNames.length, 6);
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

  const allProducts = new Set([...produtosAnterior.keys(), ...produtosAtual.keys()]);
  const comparativoProdutos: ProductComparison[] = [];
  for (const prod of allProducts) {
    const ant = produtosAnterior.get(prod) || { qtd: 0, valor: 0, codigo: "" };
    const atu = produtosAtual.get(prod) || { qtd: 0, valor: 0, codigo: ant.codigo };
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

  // ---- Merge validation reports into one combined report ----------------------
  const allDiagnostics = periodos.flatMap((p) => p.validationReport.diagnostics);
  const totalChecks = periodos.reduce((s, p) => s + p.validationReport.totalChecks, 0);
  const passedChecks = periodos.reduce((s, p) => s + p.validationReport.passedChecks, 0);

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
    clientesNovos,
    clientesPerdidos,
    clientesMantidos,
    faturamentoNovos,
    faturamentoPerdidos,
    variacaoFaturamento,
    comparativoProdutos,
    inconsistencias,
    classificacaoCarteira,
    validationReport,
  };
}
