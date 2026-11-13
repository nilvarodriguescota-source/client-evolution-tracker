/**
 * Leitura de números como o relatório os imprime.
 *
 * É o ÚNICO ponto do sistema que converte uma célula em número. Quantidade,
 * preço médio, venda bruta, subtotais e totais da aba passam todos por aqui, para
 * que nenhum módulo leia "1.234,56" de um jeito e outro módulo de outro.
 *
 * Formatos aceitos:
 *   - número nativo da planilha ........................ 1234.56
 *   - brasileiro ....................................... "1.234,56"  "1234,56"  "0,5"
 *   - com moeda e espaços .............................. "R$ 1.234,56"  "1 234,56" (espaço comum ou não separável)
 *   - negativos ........................................ "-1.234,56"  "(1.234,56)"  "1.234,56-"
 *   - padrão americano, quando não há ambiguidade ...... "1,234.56"  "1234.56"
 *
 * Regra para o caso ambíguo de um único ponto: no Brasil "1.234" é mil duzentos e
 * trinta e quatro (separador de milhar seguido de exatamente 3 dígitos); qualquer
 * outro uso do ponto ("1.5", "0.500", "12.3456") é casa decimal.
 *
 * Texto que não é número ("PARAFUSO", "01/06/2025", "12%") devolve null — nunca 0
 * — para que quem chama saiba distinguir "zero" de "não é número".
 */

/** Espaços comuns, inclusive os não separáveis que o Excel e o HTML usam. */
const RE_ESPACOS = /[\s\u00a0\u202f\u2007]/g;

export function lerNumero(value: unknown): number | null {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value !== "string") return null;

  let texto = value.replace(RE_ESPACOS, "");
  if (!texto) return null;

  // Moeda
  texto = texto.replace(/^R\$/i, "").replace(/^\$/, "");

  // Sinal: "-10", "+10", "(10)", "10-"
  let negativo = false;
  if (/^\(.*\)$/.test(texto)) {
    negativo = true;
    texto = texto.slice(1, -1);
  }
  if (/^[-+]/.test(texto)) {
    negativo = negativo !== (texto[0] === "-");
    texto = texto.slice(1);
  } else if (/-$/.test(texto)) {
    negativo = !negativo;
    texto = texto.slice(0, -1);
  }
  texto = texto.replace(/^R\$/i, "");

  if (!/^[\d.,]+$/.test(texto) || !/\d/.test(texto)) return null;

  const pontos = (texto.match(/\./g) ?? []).length;
  const virgulas = (texto.match(/,/g) ?? []).length;

  let normalizado: string;
  if (pontos > 0 && virgulas > 0) {
    // O separador que aparece por último é o decimal.
    const decimal = texto.lastIndexOf(",") > texto.lastIndexOf(".") ? "," : ".";
    const milhar = decimal === "," ? "." : ",";
    if ((decimal === "," ? virgulas : pontos) > 1) return null;
    if (!grupoDeMilharValido(texto.slice(0, texto.lastIndexOf(decimal)), milhar)) return null;
    normalizado = texto.split(milhar).join("").replace(decimal, ".");
  } else if (virgulas > 0) {
    if (virgulas === 1) {
      // Vírgula única é decimal no padrão brasileiro: "123,45", "0,5".
      normalizado = texto.replace(",", ".");
    } else {
      if (!grupoDeMilharValido(texto, ",")) return null;
      normalizado = texto.split(",").join("");
    }
  } else if (pontos > 0) {
    if (pontos > 1) {
      if (!grupoDeMilharValido(texto, ".")) return null;
      normalizado = texto.split(".").join("");
    } else if (/^[1-9]\d{0,2}\.\d{3}$/.test(texto)) {
      // "1.234" — milhar brasileiro
      normalizado = texto.replace(".", "");
    } else {
      normalizado = texto;
    }
  } else {
    normalizado = texto;
  }

  if (!/^\d*\.?\d+$|^\d+\.?$/.test(normalizado)) return null;
  const n = Number(normalizado);
  if (!Number.isFinite(n)) return null;
  return negativo ? -n : n;
}

/** "1.234.567" é milhar válido; "1.23.4" não é. */
function grupoDeMilharValido(parteInteira: string, separador: string): boolean {
  const grupos = parteInteira.split(separador);
  if (grupos.length === 1) return /^\d+$/.test(grupos[0]);
  if (!/^\d{1,3}$/.test(grupos[0])) return false;
  return grupos.slice(1).every((g) => /^\d{3}$/.test(g));
}

/** A célula carrega um número (em qualquer dos formatos aceitos). */
export function ehNumero(value: unknown): boolean {
  return lerNumero(value) !== null;
}

/** Número da célula, ou 0 quando ela não carrega número. */
export function paraNumero(value: unknown): number {
  return lerNumero(value) ?? 0;
}

/**
 * Soma em centavos inteiros (4 casas), para que somar mil linhas de R$ 0,10 dê
 * exatamente R$ 100,00 e não R$ 99,99999999.
 */
export function somarValores(valores: number[]): number {
  let acc = 0;
  for (const v of valores) acc += Math.round(v * 10000);
  return acc / 10000;
}
