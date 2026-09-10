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
  /** Label used across the analysis (tab name, or the period read from content) */
  nome: string;
  /** Tab name exactly as it appears in the file */
  sheetName: string;
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

      // Layer 1: Existence — client code and name must exist
      validator.checkExistence(codigoCliente, excelRow, coluna, "codigoCliente");
      validator.checkExistence(nomeCliente, excelRow, coluna, "nomeCliente");

      // Layer 3: Integrity — client name should not be too short
      if (nomeCliente) {
        validator.checkStringIntegrity(nomeCliente, excelRow, coluna, "nomeCliente", {
          minLength: 2,
        });
      }

      clientEntries.push({ key: nomeCliente, label: nomeCliente, row: excelRow });

      // The same client can be listed twice in one tab (two report pages):
      // keep the first block and append to it instead of dropping the earlier rows.
      const existente = clientes.get(nomeCliente);
      if (existente) {
        currentClient = existente;
        subtotais.get(existente)!.push(0);
      } else {
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

    comparativoProdutos.push({
      produto: prod,
      codigo,
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
  const evolucaoPeriodos = resumirEvolucao(periodos);

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
    validationReport,
  };
}
