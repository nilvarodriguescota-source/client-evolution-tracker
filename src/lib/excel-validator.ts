/**
 * Excel Data Validation Engine
 *
 * Implements 5+ validation layers:
 *   1. Existence   – field is present, not null/undefined/empty
 *   2. Type        – value matches expected type (number, string, date)
 *   3. Integrity   – value is not corrupted, truncated, or out of range
 *   4. Structure   – row/column layout matches expected spreadsheet schema
 *   5. Traceability – every value maps back to an exact Excel cell reference
 *
 * Additional cross-cutting checks:
 *   - Mathematical consistency (qty × price ≈ total)
 *   - Duplicate detection (clients and products)
 *   - Cross-column relationship validation
 *   - Aggregation integrity (sum of parts = reported total)
 */

import { lerNumero } from "./numero-br";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export type ValidationSeverity = "error" | "warning" | "info";

export interface ValidationDiagnostic {
  /** Which layer detected the issue */
  layer:
    | "existence"
    | "type"
    | "integrity"
    | "structure"
    | "traceability"
    | "math"
    | "duplicate"
    | "cross-field"
    /** Spreadsheet × processed figures: what the file says against what the system computed */
    | "conferencia";
  severity: ValidationSeverity;
  /** Excel cell reference (e.g. "Sheet1!B12") */
  cellRef: string;
  /** Human-readable field name */
  field: string;
  /** What was found */
  actual: string;
  /** What was expected */
  expected: string;
  /** Business impact description */
  impact: string;
}

export interface CellOrigin {
  sheet: string;
  row: number;   // 1-based row in the sheet
  col: number;   // 0-based column index
  colLetter: string;
  cellRef: string;
  rawValue: unknown;
}

export interface TracedValue<T = unknown> {
  value: T;
  origin: CellOrigin;
}

export interface ValidationReport {
  valid: boolean;
  diagnostics: ValidationDiagnostic[];
  totalChecks: number;
  passedChecks: number;
  failedChecks: number;
  /** Per-layer summary */
  layerSummary: Record<string, { passed: number; failed: number }>;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const PRECISION_DIGITS = 4;

/** Convert 0-based column index to Excel letter (0→A, 25→Z, 26→AA) */
export function colIndexToLetter(idx: number): string {
  let letter = "";
  let n = idx;
  while (n >= 0) {
    letter = String.fromCharCode((n % 26) + 65) + letter;
    n = Math.floor(n / 26) - 1;
  }
  return letter;
}

/** Build an Excel-style cell reference */
export function cellRef(sheet: string, row: number, col: number): string {
  return `${sheet}!${colIndexToLetter(col)}${row}`;
}

/** High-precision rounding to N decimal places */
export function preciseRound(value: number, decimals: number): number {
  const factor = Math.pow(10, decimals);
  return Math.round((value + Number.EPSILON) * factor) / factor;
}

/** Internal calculation preserving up to PRECISION_DIGITS decimals */
export function preciseMultiply(a: number, b: number): number {
  return preciseRound(a * b, PRECISION_DIGITS);
}

/** Sum an array with high-precision accumulation */
export function preciseSum(values: number[]): number {
  const sum = values.reduce((acc, v) => acc + v, 0);
  return preciseRound(sum, PRECISION_DIGITS);
}

/** Format for presentation: exactly 2 decimal places with proper rounding */
export function toDisplay(value: number): number {
  return preciseRound(value, 2);
}

// ---------------------------------------------------------------------------
// Validation Engine
// ---------------------------------------------------------------------------

export class ExcelValidator {
  private diagnostics: ValidationDiagnostic[] = [];
  private totalChecks = 0;
  private passedChecks = 0;
  private layerCounts: Record<string, { passed: number; failed: number }> = {};

  constructor(private sheetName: string) {}

  // ---- internal bookkeeping --------------------------------------------------

  private pass(layer: string): void {
    this.totalChecks++;
    this.passedChecks++;
    if (!this.layerCounts[layer]) this.layerCounts[layer] = { passed: 0, failed: 0 };
    this.layerCounts[layer].passed++;
  }

  private fail(d: ValidationDiagnostic): void {
    this.totalChecks++;
    this.diagnostics.push(d);
    if (!this.layerCounts[d.layer]) this.layerCounts[d.layer] = { passed: 0, failed: 0 };
    this.layerCounts[d.layer].failed++;
  }

  // ---- Layer 1: Existence ----------------------------------------------------

  checkExistence(
    value: unknown,
    row: number,
    col: number,
    fieldName: string,
  ): boolean {
    const ref = cellRef(this.sheetName, row, col);
    if (value === null || value === undefined || (typeof value === "string" && value.trim() === "")) {
      this.fail({
        layer: "existence",
        severity: "error",
        cellRef: ref,
        field: fieldName,
        actual: String(value ?? "null"),
        expected: "Valor presente e nao vazio",
        impact: `Campo obrigatorio '${fieldName}' ausente — registro sera ignorado ou processado incorretamente.`,
      });
      return false;
    }
    this.pass("existence");
    return true;
  }

  // ---- Layer 2: Type ---------------------------------------------------------

  checkType(
    value: unknown,
    expectedType: "number" | "string" | "date",
    row: number,
    col: number,
    fieldName: string,
  ): boolean {
    const ref = cellRef(this.sheetName, row, col);

    if (expectedType === "number") {
      // Same reading the analysis uses: "1.234,56" is a number, "abc" is not.
      if (lerNumero(value) === null) {
        this.fail({
          layer: "type",
          severity: "error",
          cellRef: ref,
          field: fieldName,
          actual: `${typeof value} (${String(value)})`,
          expected: "number",
          impact: `Valor '${fieldName}' nao e numerico — calculos dependentes serao comprometidos.`,
        });
        return false;
      }
    }

    if (expectedType === "string") {
      if (typeof value !== "string" && typeof value !== "number") {
        this.fail({
          layer: "type",
          severity: "warning",
          cellRef: ref,
          field: fieldName,
          actual: typeof value,
          expected: "string",
          impact: `Valor '${fieldName}' nao e texto — pode haver perda de informacao na conversao.`,
        });
        return false;
      }
    }

    this.pass("type");
    return true;
  }

  // ---- Layer 3: Integrity ----------------------------------------------------

  checkNumericIntegrity(
    value: number,
    row: number,
    col: number,
    fieldName: string,
    options: { min?: number; max?: number; allowNegative?: boolean } = {},
  ): boolean {
    const ref = cellRef(this.sheetName, row, col);
    const { min, max, allowNegative = false } = options;

    if (!Number.isFinite(value)) {
      this.fail({
        layer: "integrity",
        severity: "error",
        cellRef: ref,
        field: fieldName,
        actual: String(value),
        expected: "Numero finito",
        impact: `Valor '${fieldName}' e infinito ou NaN — dado corrompido.`,
      });
      return false;
    }

    if (!allowNegative && value < 0) {
      this.fail({
        layer: "integrity",
        severity: "error",
        cellRef: ref,
        field: fieldName,
        actual: String(value),
        expected: "Valor >= 0",
        impact: `Valor negativo inesperado em '${fieldName}' — possivel corrupcao ou erro de entrada.`,
      });
      return false;
    }

    if (min !== undefined && value < min) {
      this.fail({
        layer: "integrity",
        severity: "warning",
        cellRef: ref,
        field: fieldName,
        actual: String(value),
        expected: `>= ${min}`,
        impact: `Valor '${fieldName}' abaixo do minimo esperado.`,
      });
      return false;
    }

    if (max !== undefined && value > max) {
      this.fail({
        layer: "integrity",
        severity: "warning",
        cellRef: ref,
        field: fieldName,
        actual: String(value),
        expected: `<= ${max}`,
        impact: `Valor '${fieldName}' acima do maximo esperado.`,
      });
      return false;
    }

    // Check for possible truncation (numbers that look oddly rounded)
    const strVal = String(value);
    if (strVal.includes("e+") || strVal.includes("e-")) {
      this.fail({
        layer: "integrity",
        severity: "warning",
        cellRef: ref,
        field: fieldName,
        actual: strVal,
        expected: "Notacao decimal padrao",
        impact: `Valor '${fieldName}' em notacao cientifica — possivel perda de precisao.`,
      });
      return false;
    }

    this.pass("integrity");
    return true;
  }

  checkStringIntegrity(
    value: string,
    row: number,
    col: number,
    fieldName: string,
    options: { maxLength?: number; minLength?: number } = {},
  ): boolean {
    const ref = cellRef(this.sheetName, row, col);
    const { maxLength = 500, minLength = 1 } = options;

    if (value.length < minLength) {
      this.fail({
        layer: "integrity",
        severity: "warning",
        cellRef: ref,
        field: fieldName,
        actual: `"${value}" (${value.length} chars)`,
        expected: `Comprimento >= ${minLength}`,
        impact: `Texto '${fieldName}' possivelmente truncado ou incompleto.`,
      });
      return false;
    }

    if (value.length > maxLength) {
      this.fail({
        layer: "integrity",
        severity: "warning",
        cellRef: ref,
        field: fieldName,
        actual: `${value.length} chars`,
        expected: `Comprimento <= ${maxLength}`,
        impact: `Texto '${fieldName}' excede comprimento maximo esperado — possivel dado corrompido.`,
      });
      return false;
    }

    this.pass("integrity");
    return true;
  }

  // ---- Layer 4: Structure ----------------------------------------------------

  checkStructure(
    row: unknown[],
    rowIndex: number,
    expectedMinCols: number,
    context: string,
  ): boolean {
    const ref = cellRef(this.sheetName, rowIndex, 0);

    if (!Array.isArray(row)) {
      this.fail({
        layer: "structure",
        severity: "error",
        cellRef: ref,
        field: context,
        actual: typeof row,
        expected: "Array (linha da planilha)",
        impact: `Linha ${rowIndex} nao e um array — estrutura da planilha incompativel.`,
      });
      return false;
    }

    if (row.length < expectedMinCols) {
      this.fail({
        layer: "structure",
        severity: "error",
        cellRef: ref,
        field: context,
        actual: `${row.length} colunas`,
        expected: `>= ${expectedMinCols} colunas`,
        impact: `Linha ${rowIndex} possui colunas insuficientes — dados incompletos para processamento.`,
      });
      return false;
    }

    this.pass("structure");
    return true;
  }

  // ---- Layer 5: Traceability -------------------------------------------------

  traceValue<T>(
    value: T,
    row: number,
    col: number,
  ): TracedValue<T> {
    this.pass("traceability");
    return {
      value,
      origin: {
        sheet: this.sheetName,
        row,
        col,
        colLetter: colIndexToLetter(col),
        cellRef: cellRef(this.sheetName, row, col),
        rawValue: value,
      },
    };
  }

  // ---- Cross-cutting: Math consistency --------------------------------------

  checkMathConsistency(
    quantity: number,
    unitPrice: number,
    reportedTotal: number,
    row: number,
    context: string,
    tolerance = 0.01,
    /** Column of the reported total, for the cell reference */
    col = 5,
  ): { valid: boolean; calculated: number; difference: number } {
    const calculated = preciseMultiply(quantity, unitPrice);
    const difference = preciseRound(Math.abs(calculated - reportedTotal), PRECISION_DIGITS);
    const ref = cellRef(this.sheetName, row, col);

    if (difference > tolerance) {
      this.fail({
        layer: "math",
        severity: "error",
        cellRef: ref,
        field: "valorTotal",
        actual: `${reportedTotal} (planilha)`,
        expected: `${calculated} (qtd ${quantity} x preco ${unitPrice})`,
        impact: `${context}: Divergencia de R$ ${difference.toFixed(4)} entre valor reportado e calculado — erro no calculo ou na planilha de origem.`,
      });
      return { valid: false, calculated, difference };
    }

    this.pass("math");
    return { valid: true, calculated, difference };
  }

  // ---- Cross-cutting: Aggregation integrity ---------------------------------

  checkAggregation(
    parts: number[],
    reportedTotal: number,
    row: number,
    context: string,
    tolerance = 0.01,
    col = 5,
  ): boolean {
    const calculated = preciseSum(parts);
    const diff = Math.abs(calculated - reportedTotal);
    const ref = cellRef(this.sheetName, row, col);

    if (diff > tolerance) {
      this.fail({
        layer: "math",
        severity: "error",
        cellRef: ref,
        field: "totalGeral",
        actual: `${reportedTotal}`,
        expected: `${calculated} (soma dos itens)`,
        impact: `${context}: Total geral diverge da soma dos produtos em R$ ${diff.toFixed(4)}.`,
      });
      return false;
    }

    this.pass("math");
    return true;
  }

  // ---- Cross-cutting: Duplicate detection -----------------------------------

  checkDuplicates(
    items: { key: string; label: string; row: number }[],
    context: string,
  ): string[] {
    const seen = new Map<string, number>();
    const duplicates: string[] = [];

    for (const item of items) {
      const prevRow = seen.get(item.key);
      if (prevRow !== undefined) {
        duplicates.push(item.key);
        this.fail({
          layer: "duplicate",
          severity: "warning",
          cellRef: cellRef(this.sheetName, item.row, 0),
          field: item.label,
          actual: `Duplicado (tambem na linha ${prevRow})`,
          expected: "Registro unico",
          impact: `${context}: '${item.label}' aparece mais de uma vez — pode causar duplicidade nos calculos.`,
        });
      } else {
        this.pass("duplicate");
      }
      seen.set(item.key, item.row);
    }

    return duplicates;
  }

  // ---- Conferência: spreadsheet × processed ---------------------------------

  /**
   * Compares a figure taken from the spreadsheet with the figure the system
   * computed from it. A divergence is always an error: a number that does not
   * match the file must never be presented as valid.
   */
  checkConferencia(
    esperado: number,
    apurado: number,
    cellRefOrigem: string,
    field: string,
    descricao: string,
    tolerance = 0.005,
  ): boolean {
    const diff = preciseRound(Math.abs(esperado - apurado), PRECISION_DIGITS);
    // NaN (a figure that was never computed) is a divergence, not a pass.
    if (!(diff <= tolerance)) {
      this.fail({
        layer: "conferencia",
        severity: "error",
        cellRef: cellRefOrigem,
        field,
        actual: `${apurado} (apurado pelo sistema)`,
        expected: `${esperado} (planilha)`,
        impact: `${descricao}: divergência de ${diff.toFixed(4)} entre a planilha e o valor processado.`,
      });
      return false;
    }
    this.pass("conferencia");
    return true;
  }

  // ---- Cross-field consistency -----------------------------------------------

  checkCrossField(
    condition: boolean,
    row: number,
    col: number,
    fieldName: string,
    description: string,
    impact: string,
  ): boolean {
    if (!condition) {
      this.fail({
        layer: "cross-field",
        severity: "warning",
        cellRef: cellRef(this.sheetName, row, col),
        field: fieldName,
        actual: "Inconsistente",
        expected: description,
        impact,
      });
      return false;
    }
    this.pass("cross-field");
    return true;
  }

  // ---- Report ----------------------------------------------------------------

  getReport(): ValidationReport {
    return {
      valid: this.diagnostics.filter((d) => d.severity === "error").length === 0,
      diagnostics: [...this.diagnostics],
      totalChecks: this.totalChecks,
      passedChecks: this.passedChecks,
      failedChecks: this.totalChecks - this.passedChecks,
      layerSummary: { ...this.layerCounts },
    };
  }

  getDiagnostics(): ValidationDiagnostic[] {
    return [...this.diagnostics];
  }

  hasErrors(): boolean {
    return this.diagnostics.some((d) => d.severity === "error");
  }

  merge(other: ExcelValidator): void {
    this.diagnostics.push(...other.diagnostics);
    this.totalChecks += other.totalChecks;
    this.passedChecks += other.passedChecks;
    for (const [layer, counts] of Object.entries(other.layerCounts)) {
      if (!this.layerCounts[layer]) this.layerCounts[layer] = { passed: 0, failed: 0 };
      this.layerCounts[layer].passed += counts.passed;
      this.layerCounts[layer].failed += counts.failed;
    }
  }
}
