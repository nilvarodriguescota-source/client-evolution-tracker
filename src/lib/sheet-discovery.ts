/**
 * Sheet discovery — the single place where the workbook is opened, its tabs are
 * enumerated, identified and ordered.
 *
 * Everything that needs to know "which tabs hold report data and in which order"
 * must go through here, so a tab can never be silently dropped by one module
 * while another module sees it.
 *
 * Design rules:
 *   - Every tab of the file is considered, including hidden ones.
 *   - A tab is a report because of what it *contains*, never because of its name
 *     or its position in the file.
 *   - There is no fixed number of periods: 2, 3 or 12 months all work.
 *   - A tab that is left out is reported with a reason instead of disappearing.
 */

import * as XLSX from "xlsx";

export type SheetRows = (string | number | boolean | null)[][];

// ---------------------------------------------------------------------------
// Text helpers
// ---------------------------------------------------------------------------

/** lower-case, accent-free, trimmed — for comparing spreadsheet labels. */
export function normalizarTexto(value: unknown): string {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

/** Any dash used by the exporters: -, – (en), — (em). */
export const TRACOS = "-–—";

// ---------------------------------------------------------------------------
// Column mapping
// ---------------------------------------------------------------------------

export interface ColumnMap {
  codigo: number;
  descricao: number;
  unidade: number;
  quantidade: number;
  precoMedio: number;
  vendaBruta: number;
}

/** Layout of the standard "Vendas de produto no período" export. */
export const COLUNAS_PADRAO: ColumnMap = {
  codigo: 0,
  descricao: 1,
  unidade: 2,
  quantidade: 3,
  precoMedio: 4,
  vendaBruta: 5,
};

/** A column is absent from the sheet. */
export const COLUNA_AUSENTE = -1;

/**
 * Header labels, most specific first — the first pattern that matches a cell
 * and whose field is still unassigned wins.
 */
const PADROES_CABECALHO: { field: keyof ColumnMap; test: RegExp }[] = [
  { field: "codigo", test: /^(cod|codigo|cod\.)/ },
  { field: "descricao", test: /^(desc|produto|mercadoria|item)/ },
  { field: "unidade", test: /^(un|unid)/ },
  { field: "quantidade", test: /(qtd|qtde|quant)/ },
  { field: "precoMedio", test: /(preco|prc|vlr\.? unit|valor unit)/ },
  { field: "vendaBruta", test: /(venda bruta|vlr\.? bruto|valor bruto|vlr\.? total|valor total|faturamento|total)/ },
];

/** Fields without which a row cannot be read as a product. */
const CAMPOS_ESSENCIAIS: (keyof ColumnMap)[] = ["codigo", "descricao", "quantidade", "vendaBruta"];

/**
 * Reads a product header row ("Código | Descrição | UN | Qtd | Preço | Venda").
 * Returns null when the row is not a header, so the caller keeps the map it has.
 */
export function lerCabecalhoProdutos(row: SheetRows[number]): ColumnMap | null {
  if (!row) return null;

  const mapa: Partial<ColumnMap> = {};
  for (let col = 0; col < row.length; col++) {
    const texto = normalizarTexto(row[col]);
    if (!texto) continue;
    for (const { field, test } of PADROES_CABECALHO) {
      if (mapa[field] === undefined && test.test(texto)) {
        mapa[field] = col;
        break;
      }
    }
  }

  // A header row must name the fields the analysis cannot do without.
  if (!CAMPOS_ESSENCIAIS.every((f) => mapa[f] !== undefined)) return null;

  return {
    codigo: mapa.codigo!,
    descricao: mapa.descricao!,
    unidade: mapa.unidade ?? COLUNA_AUSENTE,
    quantidade: mapa.quantidade!,
    precoMedio: mapa.precoMedio ?? COLUNA_AUSENTE,
    vendaBruta: mapa.vendaBruta!,
  };
}

/** Shifts the standard layout so it starts at `offset` (export with lead columns). */
export function deslocarColunas(offset: number): ColumnMap {
  return {
    codigo: COLUNAS_PADRAO.codigo + offset,
    descricao: COLUNAS_PADRAO.descricao + offset,
    unidade: COLUNAS_PADRAO.unidade + offset,
    quantidade: COLUNAS_PADRAO.quantidade + offset,
    precoMedio: COLUNAS_PADRAO.precoMedio + offset,
    vendaBruta: COLUNAS_PADRAO.vendaBruta + offset,
  };
}

// ---------------------------------------------------------------------------
// Client header
// ---------------------------------------------------------------------------

export interface CabecalhoCliente {
  codigo: string;
  /** Client name alone — no dates, no phone numbers appended */
  nome: string;
  /**
   * Every "Rótulo: valor" the export appends to the client line, keyed by the
   * normalised label ("data da ultima compra", "fone", "celular", …). Kept whole
   * so nothing the report carries is thrown away.
   */
  campos: Record<string, string>;
  /** Convenience readings of the fields the analysis uses by name */
  dataUltimaCompra?: string;
  fone?: string;
  celular?: string;
  /** Column the "Cliente:" label was found in */
  coluna: number;
}

/**
 * The client line, in the three shapes the exports use. Tried in order, because
 * "Cliente: SUPER-MERCADO ZE" has no code while "Cliente: AB-10 - INDUSTRIA" has
 * one that contains a dash.
 */
const RE_CLIENTE = [
  // code (may contain dashes) separated by a spaced dash: "AB-10 - INDUSTRIA X"
  new RegExp(
    `^clientes?\\s*:?\\s*([^\\s${TRACOS}]+(?:[${TRACOS}][^\\s${TRACOS}]+)*)\\s+[${TRACOS}]\\s+(.+)$`,
    "i",
  ),
  // numeric code, dash with or without spaces: "Cliente:1001-INDUSTRIA X"
  new RegExp(`^clientes?\\s*:?\\s*(\\d[\\w.]*)\\s*[${TRACOS}]\\s*(.+)$`, "i"),
  // no code at all: "Cliente: INDUSTRIA X"
  /^clientes?\s*:?\s*()(.+)$/i,
];

/** Splits on a dash that has spaces around it — "A - B", never "AB-10". */
const RE_SEPARADOR = new RegExp(`\\s+[${TRACOS}]\\s+`);

/** A trailing segment shaped like "Rótulo: valor". */
const RE_CAMPO = /^([^:]{2,40}?)\s*:\s*(.*)$/;

/** Field labels the analysis reads by name, in order of preference. */
const ROTULOS = {
  dataUltimaCompra: ["data da ultima compra", "data ultima compra", "ultima compra"],
  fone: ["fone", "telefone", "tel"],
  celular: ["celular", "cel", "whatsapp", "whats"],
};

function primeiroCampo(campos: Record<string, string>, rotulos: string[]): string | undefined {
  for (const r of rotulos) {
    const v = campos[r];
    if (v) return v;
  }
  return undefined;
}

/**
 * Splits the text after the client code into the name and the labelled fields
 * the exports append to it:
 *
 *   "BRASA - Data da última compra: 08/09/2026 - Fone: (11) 3333-3333"
 *   →  nome "BRASA", campos { "data da ultima compra": "08/09/2026", fone: "(11) 3333-3333" }
 *
 * Segments before the first labelled field stay in the name, so a client called
 * "SUPER - MERCADO ZE" keeps its dash.
 */
export function separarCamposCliente(resto: string): {
  nome: string;
  campos: Record<string, string>;
} {
  const segmentos = resto.split(RE_SEPARADOR);
  const campos: Record<string, string> = {};
  let nome = segmentos[0].trim();
  let rotuloAtual: string | null = null;

  for (let i = 1; i < segmentos.length; i++) {
    const seg = segmentos[i].trim();
    if (!seg) continue;

    const m = seg.match(RE_CAMPO);
    if (m) {
      rotuloAtual = normalizarTexto(m[1]);
      campos[rotuloAtual] = m[2].trim();
    } else if (rotuloAtual) {
      // The value itself contained a spaced dash — put it back together.
      campos[rotuloAtual] = `${campos[rotuloAtual]} - ${seg}`.trim();
    } else {
      // Still the name: no labelled field has started yet.
      nome = `${nome} - ${seg}`;
    }
  }

  return { nome: nome.trim(), campos };
}

/**
 * Finds the client header anywhere in the row — some exports indent the report
 * or merge the title into a column other than the first.
 */
export function lerCabecalhoCliente(row: SheetRows[number]): CabecalhoCliente | null {
  if (!row) return null;

  for (let col = 0; col < row.length; col++) {
    const bruto = String(row[col] ?? "").trim();
    if (!bruto || !/^clientes?\s*:/i.test(bruto)) continue;

    for (const re of RE_CLIENTE) {
      const match = bruto.match(re);
      if (!match) continue;

      // "NOME - Data da última compra: … - Fone: …" becomes a name plus fields.
      const { nome, campos } = separarCamposCliente(match[2]);
      if (!nome) continue;

      return {
        codigo: (match[1] ?? "").trim(),
        nome,
        campos,
        dataUltimaCompra: primeiroCampo(campos, ROTULOS.dataUltimaCompra),
        fone: primeiroCampo(campos, ROTULOS.fone),
        celular: primeiroCampo(campos, ROTULOS.celular),
        coluna: col,
      };
    }
  }

  return null;
}

// ---------------------------------------------------------------------------
// Period detection (name and content)
// ---------------------------------------------------------------------------

export const MESES: Record<string, number> = {
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

const NOME_MES = [
  "janeiro", "fevereiro", "março", "abril", "maio", "junho",
  "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
];

export interface Periodo {
  mes?: number;
  ano?: number;
}

/** Reads month/year out of a sheet name ("junho", "junho/2025", "06/2025", "2025-06"). */
export function lerPeriodoDoNome(name: string): Periodo {
  const norm = normalizarTexto(name);

  let m = norm.match(/(?:^|\D)(\d{4})[-/.](\d{1,2})(?:\D|$)/);
  if (m && Number(m[2]) >= 1 && Number(m[2]) <= 12) {
    return { ano: Number(m[1]), mes: Number(m[2]) };
  }

  m = norm.match(/(?:^|\D)(\d{1,2})[-/.](\d{4})(?:\D|$)/);
  if (m && Number(m[1]) >= 1 && Number(m[1]) <= 12) {
    return { mes: Number(m[1]), ano: Number(m[2]) };
  }

  const mes = (norm.match(/[a-z]+/g) ?? [])
    .map((w) => MESES[w])
    .find((v) => v !== undefined);

  const anoMatch = norm.match(/\b(\d{4})\b/) ?? norm.match(/[a-z]+[-/. ](\d{2})\b/);
  let ano: number | undefined;
  if (anoMatch) {
    const n = Number(anoMatch[1]);
    ano = n < 100 ? 2000 + n : n;
  }

  return { mes, ano };
}

/** "Gerado em 05/09/2025" is when the file was printed, not the period it covers. */
const RE_RODAPE = /^(gerado|emitido|impresso|pagina|página)/i;

/**
 * "Data da última compra: 08/09/2026" is the client's own history, not the month
 * the tab covers — it is removed before the period is read off the content.
 */
const RE_ULTIMA_COMPRA = /(data\s+d[ae]\s+)?[uú]ltima\s+compra\s*:\s*\d{1,2}\/\d{1,2}\/\d{2,4}/gi;

/**
 * Reads the period out of the sheet content, for tabs whose name says nothing
 * ("Plan1", "Relatório (2)"). Looks for the date ranges the exports print in the
 * report title and in each client header.
 */
export function lerPeriodoDoConteudo(rows: SheetRows): Periodo {
  const contagem = new Map<string, { mes: number; ano: number; n: number }>();

  for (const row of rows) {
    if (!row) continue;
    for (const cell of row) {
      const texto = String(cell ?? "")
        .replace(RE_ULTIMA_COMPRA, " ")
        .trim();
      if (!texto || RE_RODAPE.test(normalizarTexto(texto))) continue;

      // "01/06/2025 a 30/06/2025" or "Data: 01/06/2025"
      const datas = texto.match(/\b(\d{1,2})\/(\d{1,2})\/(\d{4})\b/g);
      if (datas) {
        for (const d of datas) {
          const [, , mes, ano] = d.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/)!;
          const chave = `${ano}-${mes}`;
          const e = contagem.get(chave) ?? { mes: Number(mes), ano: Number(ano), n: 0 };
          e.n++;
          contagem.set(chave, e);
        }
        continue;
      }

      // "Período: junho/2025"
      const norm = normalizarTexto(texto);
      if (/(periodo|referencia|competencia|mes)/.test(norm)) {
        const { mes, ano } = lerPeriodoDoNome(norm);
        if (mes) {
          const chave = `${ano ?? "?"}-${mes}`;
          const e = contagem.get(chave) ?? { mes, ano: ano ?? 0, n: 0 };
          e.n++;
          contagem.set(chave, e);
        }
      }
    }
  }

  if (contagem.size === 0) return {};

  // The period the sheet covers is the month that shows up the most.
  const melhor = [...contagem.values()].sort((a, b) => b.n - a.n)[0];
  return { mes: melhor.mes, ano: melhor.ano || undefined };
}

/** Human label for a detected period: "junho/2025" or "junho". */
export function rotularPeriodo({ mes, ano }: Periodo): string | null {
  if (!mes) return null;
  const nome = NOME_MES[mes - 1];
  return ano ? `${nome}/${ano}` : nome;
}

// ---------------------------------------------------------------------------
// Chronological ordering
// ---------------------------------------------------------------------------

export interface Ordenavel extends Periodo {
  /** position in the workbook, used as a stable tie-break */
  index: number;
}

/**
 * Orders the items that name a month chronologically. Items with no detectable
 * month keep their original slot instead of disabling the ordering for
 * everybody — one unrecognised tab must not scramble the months.
 */
export function ordenarPorPeriodo<T extends Ordenavel>(items: T[]): T[] {
  const posicoes = items.map((_, i) => i).filter((i) => items[i].mes !== undefined);
  if (posicoes.length < 2) return [...items];

  const comMes = posicoes.map((i) => items[i]);
  const todosComAno = comMes.every((p) => p.ano !== undefined);

  let ordenados: T[];
  if (todosComAno) {
    ordenados = [...comMes].sort(
      (a, b) => a.ano! - b.ano! || a.mes! - b.mes! || a.index - b.index,
    );
  } else {
    // No explicit years: choose the 12-month rotation that packs the months into
    // the shortest span, so novembro → dezembro → janeiro sorts across the year
    // end instead of collapsing into janeiro → novembro → dezembro.
    const meses = [...new Set(comMes.map((p) => p.mes!))];
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
    ordenados = [...comMes].sort(
      (a, b) => rank(a.mes!) - rank(b.mes!) || a.index - b.index,
    );
  }

  const saida = [...items];
  posicoes.forEach((pos, i) => {
    saida[pos] = ordenados[i];
  });
  return saida;
}

// ---------------------------------------------------------------------------
// Workbook discovery
// ---------------------------------------------------------------------------

export interface AbaDescoberta {
  /** Tab name exactly as it appears in the file */
  sheetName: string;
  /** Label used across the analysis (detected period, or the tab name) */
  nome: string;
  rows: SheetRows;
  /** Column layout detected for this tab */
  colunas: ColumnMap;
  mes?: number;
  ano?: number;
  /** Where the period came from */
  origemPeriodo: "nome-da-aba" | "conteudo" | "ordem-do-arquivo";
  oculta: boolean;
  /** Position in the workbook */
  index: number;
  /** Rows that carry data (non-empty) */
  linhas: number;
}

export interface AbaIgnorada {
  sheetName: string;
  motivo: string;
}

export interface DescobertaPlanilha {
  /** Report tabs, oldest first */
  abas: AbaDescoberta[];
  ignoradas: AbaIgnorada[];
  /** Every tab name in the file, in file order */
  todasAsAbas: string[];
}

/**
 * Opens the workbook.
 *
 * NOTE: legacy .xls (CFB/BIFF8) files fail to parse when a raw ArrayBuffer is
 * passed — SheetJS returns a single bogus "Sheet1" instead of the real tabs.
 * Wrapping in a Uint8Array parses every sheet correctly for both .xls and .xlsx.
 */
export function lerPastaDeTrabalho(buffer: ArrayBuffer): XLSX.WorkBook {
  return XLSX.read(new Uint8Array(buffer), { type: "array" });
}

function linhasDaAba(sheet: XLSX.WorkSheet): SheetRows {
  return XLSX.utils.sheet_to_json<SheetRows[number]>(sheet, { header: 1, defval: null });
}

function abasOcultas(wb: XLSX.WorkBook): Set<string> {
  const ocultas = new Set<string>();
  for (const meta of wb.Workbook?.Sheets ?? []) {
    // 0 = visible, 1 = hidden, 2 = very hidden
    if (meta && meta.Hidden && meta.name) ocultas.add(meta.name);
  }
  return ocultas;
}

/** True when the tab carries report data — decided by content, never by name. */
function pareceRelatorio(rows: SheetRows): boolean {
  for (const row of rows) {
    if (!row) continue;
    if (lerCabecalhoCliente(row)) return true;
    if (lerCabecalhoProdutos(row)) return true;
  }
  return false;
}

/** Column layout of a tab: from its header row, or from where the report starts. */
function detectarColunas(rows: SheetRows): ColumnMap {
  for (const row of rows) {
    const mapa = lerCabecalhoProdutos(row);
    if (mapa) return mapa;
  }

  // No header row: fall back to the standard layout, shifted to the column the
  // client header sits in (exports that indent the whole report).
  for (const row of rows) {
    const cliente = lerCabecalhoCliente(row);
    if (cliente) return deslocarColunas(cliente.coluna);
  }

  return COLUNAS_PADRAO;
}

/**
 * Finds every tab of the workbook that holds report data, names its period and
 * puts them in chronological order. No cap on how many periods come back.
 */
export function descobrirAbas(wb: XLSX.WorkBook): DescobertaPlanilha {
  const ocultas = abasOcultas(wb);
  const abas: AbaDescoberta[] = [];
  const ignoradas: AbaIgnorada[] = [];

  wb.SheetNames.forEach((sheetName, index) => {
    const sheet = wb.Sheets[sheetName];
    if (!sheet) {
      ignoradas.push({ sheetName, motivo: "aba declarada no arquivo mas sem conteúdo legível" });
      return;
    }

    const rows = linhasDaAba(sheet);
    const linhas = rows.filter((r) => r && r.some((c) => c !== null && String(c).trim() !== "")).length;

    if (linhas === 0) {
      ignoradas.push({ sheetName, motivo: "aba vazia" });
      return;
    }

    if (!pareceRelatorio(rows)) {
      ignoradas.push({
        sheetName,
        motivo:
          "não tem estrutura de relatório (nenhuma linha \"Cliente:\" nem cabeçalho de produtos)",
      });
      return;
    }

    const doNome = lerPeriodoDoNome(sheetName);
    const doConteudo = doNome.mes === undefined ? lerPeriodoDoConteudo(rows) : {};
    const periodo = doNome.mes !== undefined ? doNome : doConteudo;

    const origemPeriodo: AbaDescoberta["origemPeriodo"] =
      doNome.mes !== undefined
        ? "nome-da-aba"
        : doConteudo.mes !== undefined
          ? "conteudo"
          : "ordem-do-arquivo";

    // Keep the tab name when it already identifies the period; only fall back to
    // the detected label when the name says nothing ("Plan1", "Relatório (2)").
    const rotulo = rotularPeriodo(periodo);
    const nome = origemPeriodo === "conteudo" && rotulo ? rotulo : sheetName;

    abas.push({
      sheetName,
      nome,
      rows,
      colunas: detectarColunas(rows),
      mes: periodo.mes,
      ano: periodo.ano,
      origemPeriodo,
      oculta: ocultas.has(sheetName),
      index,
      linhas,
    });
  });

  return {
    abas: ordenarPorPeriodo(abas),
    ignoradas,
    todasAsAbas: [...wb.SheetNames],
  };
}
