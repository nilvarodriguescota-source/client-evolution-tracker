/**
 * Conferência — planilha original × dados processados.
 *
 * Depois que as abas foram lidas e todos os módulos calcularam seus números, esta
 * rotina refaz as contas por outro caminho e compara:
 *
 *   1. faturamento de cada aba: soma das linhas de venda × subtotais e totais
 *      impressos pela própria aba;
 *   2. faturamento consolidado da janela, por aba, por produto e por cliente;
 *   3. quantidade de registros de venda;
 *   4. quantidade de clientes (carteira + funcionários);
 *   5. quantidade de produtos vendidos;
 *   6. valor de cada cliente em cada mês;
 *   7. valor de cada produto em cada mês;
 *   8. consistência dos comparativos mensais (tabela, transições e ponte).
 *
 * Toda divergência vira erro no relatório de validação e alerta na tela, para
 * que nenhum número que não confira com a planilha seja apresentado como válido.
 */

import { ExcelValidator, type ValidationReport } from "./excel-validator";
import { somarValores } from "./numero-br";
import { normalizarTexto } from "./sheet-discovery";
import type {
  CarteiraResumo,
  ClienteCarteira,
  EvolucaoConsolidada,
  FuncionarioResumo,
  PeriodData,
  PeriodoResumo,
  ProductComparison,
  TransicaoPeriodo,
} from "./excel-analyzer";

/** Diferença tolerada: meio centavo, só para absorver a representação binária. */
const TOLERANCIA = 0.005;

export interface ConferenciaAba {
  aba: string;
  periodo: string;
  registrosVenda: number;
  /** Compradores do mês (clientes + funcionários) */
  clientes: number;
  clientesCarteira: number;
  funcionarios: number;
  produtosDistintos: number;
  quantidadeTotal: number;
  /** Σ das linhas de venda — o faturamento da aba */
  faturamentoApurado: number;
  faturamentoCarteira: number;
  faturamentoFuncionarios: number;
  /** Σ dos subtotais impressos por cliente (null quando a aba não os imprime) */
  subtotaisImpressos: number | null;
  /** Total que a própria aba imprime (null quando não imprime) */
  totalImpresso: number | null;
  origemTotalImpresso: string | null;
  /** Valor de linhas que não puderam ser lidas como venda */
  valorDescartado: number;
  linhasDescartadas: number;
  /** Uma linha por verificação que falhou nesta aba */
  divergencias: string[];
  ok: boolean;
}

export interface ItemConferencia {
  /** Número da verificação pedida (1 a 8, ver o topo do arquivo) */
  grupo: number;
  /** Onde a verificação foi feita: uma aba ou "Consolidado" */
  escopo: string;
  descricao: string;
  esperado: string;
  apurado: string;
  ok: boolean;
  /** Explicação da divergência (vazio quando confere) */
  detalhe: string;
}

export interface Conferencia {
  /** true somente quando todas as verificações conferem */
  ok: boolean;
  abas: ConferenciaAba[];
  itens: ItemConferencia[];
  consolidado: {
    abas: number;
    registrosVenda: number;
    faturamento: number;
    faturamentoCarteira: number;
    faturamentoFuncionarios: number;
    clientesDistintos: number;
    funcionariosDistintos: number;
    produtosDistintos: number;
    quantidadeTotal: number;
  };
  validationReport: ValidationReport;
}

export interface EntradaConferencia {
  periodos: PeriodData[];
  comparativoProdutos: ProductComparison[];
  classificacaoCarteira: ClienteCarteira[];
  carteira: CarteiraResumo;
  transicoes: TransicaoPeriodo[];
  evolucaoPeriodos: PeriodoResumo[];
  evolucaoConsolidada: EvolucaoConsolidada;
  funcionarios: FuncionarioResumo[];
  periodosRepetidos: { periodo: string; abas: string[] }[];
}

const brl = (v: number) =>
  v.toLocaleString("pt-BR", { style: "currency", currency: "BRL", minimumFractionDigits: 2 });

const iguais = (a: number, b: number) => Math.abs(a - b) <= TOLERANCIA;

export function conferirAnalise(e: EntradaConferencia): Conferencia {
  const validator = new ExcelValidator("Conferência");
  const itens: ItemConferencia[] = [];

  /** Registers one check; a failure is an error in the validation report. */
  const conferir = (
    grupo: number,
    escopo: string,
    descricao: string,
    esperado: number,
    apurado: number,
    formato: (v: number) => string,
    cellRef: string,
    detalheFalha?: string,
  ): boolean => {
    const ok = validator.checkConferencia(
      esperado,
      apurado,
      cellRef,
      `${escopo} · ${descricao}`,
      `${escopo} · ${descricao}`,
      TOLERANCIA,
    );
    itens.push({
      grupo,
      escopo,
      descricao,
      esperado: formato(esperado),
      apurado: formato(apurado),
      ok,
      detalhe: ok
        ? ""
        : detalheFalha ??
          `planilha ${formato(esperado)} × sistema ${formato(apurado)} (diferença ${formato(Math.abs(esperado - apurado))})`,
    });
    return ok;
  };
  const conferirCondicao = (
    grupo: number,
    escopo: string,
    descricao: string,
    ok: boolean,
    detalhe: string,
    cellRef: string,
  ): boolean =>
    conferir(grupo, escopo, descricao, 0, ok ? 0 : 1, (v) => (v === 0 ? "confere" : "diverge"), cellRef, detalhe);

  const contagem = (v: number) => v.toLocaleString("pt-BR");
  const qtd = (v: number) => v.toLocaleString("pt-BR", { maximumFractionDigits: 4 });

  const { periodos } = e;
  const abas: ConferenciaAba[] = [];

  // ---- Per tab -----------------------------------------------------------------
  periodos.forEach((p, i) => {
    const antes = itens.length;
    const escopo = `Aba ${p.sheetName}`;
    const ref = `${p.sheetName}!A1`;
    const l = p.leituraAba;
    const todos = [...p.clientes.values()];

    // 1 · Faturamento da aba
    conferir(1, escopo, "Faturamento = soma das linhas de venda", l.faturamentoRegistros, p.totalFaturamento, brl, ref);
    conferir(
      1,
      escopo,
      "Faturamento = soma dos clientes",
      somarValores(todos.map((c) => c.totalGeral)),
      p.totalFaturamento,
      brl,
      ref,
    );
    conferir(
      1,
      escopo,
      "Clientes + funcionários = faturamento da aba",
      p.totalFaturamento,
      somarValores([p.faturamentoCarteira, p.faturamentoFuncionarios]),
      brl,
      ref,
    );
    if (l.subtotaisImpressos !== null) {
      conferir(
        1,
        escopo,
        "Subtotais impressos por cliente = soma das vendas desses clientes",
        l.subtotaisImpressos,
        l.faturamentoBlocosComSubtotal,
        brl,
        ref,
      );
    }
    if (l.totalImpresso) {
      conferir(
        1,
        escopo,
        `${l.totalImpresso.origem} impresso na aba = faturamento apurado`,
        l.totalImpresso.valor,
        p.totalFaturamento,
        brl,
        `${p.sheetName}!A${l.totalImpresso.linha}`,
      );
    }
    const valorDescartado = somarValores(l.linhasDescartadas.map((d) => d.valor));
    conferir(
      1,
      escopo,
      "Valores da aba fora do faturamento",
      0,
      valorDescartado,
      brl,
      ref,
      `${l.linhasDescartadas.length} linha(s) com valor não puderam ser lidas como venda ` +
        `(${l.linhasDescartadas.map((d) => `linha ${d.linha}: ${brl(d.valor)}`).join("; ")})`,
    );

    // 3 · Registros
    conferir(
      3,
      escopo,
      "Registros de venda lidos = registros nos clientes",
      l.registrosVenda,
      todos.reduce((s, c) => s + c.produtos.length, 0),
      contagem,
      ref,
    );

    // 4 · Clientes
    conferir(
      4,
      escopo,
      "Compradores = clientes da carteira + funcionários",
      p.clientes.size,
      p.clientesCarteira.size + p.funcionarios.size,
      contagem,
      ref,
    );
    conferirCondicao(
      4,
      escopo,
      "Nenhum funcionário (*) entre os clientes da carteira",
      [...p.clientesCarteira.values()].every((c) => !c.funcionario),
      "há cliente marcado com * contado na carteira",
      ref,
    );

    // 5 / 7 · Produtos do mês
    const somaProdutos = somarValores(e.comparativoProdutos.map((c) => c.valorPorPeriodo[p.nome] ?? 0));
    conferir(7, escopo, "Soma dos produtos = faturamento da aba", p.totalFaturamento, somaProdutos, brl, ref);
    conferir(
      5,
      escopo,
      "Quantidade vendida nos produtos = quantidade das linhas",
      l.quantidadeTotal,
      somarValores(e.comparativoProdutos.map((c) => c.qtdPorPeriodo[p.nome] ?? 0)),
      qtd,
      ref,
    );
    const chavesComparativo = new Set(e.comparativoProdutos.map((c) => c.chave));
    const produtosDaAba = new Set(
      todos.flatMap((c) => c.produtos.map((x) => `${x.codigo.trim()}|${normalizarTexto(x.produto)}`)),
    );
    conferir(
      5,
      escopo,
      "Produtos vendidos no mês presentes no comparativo",
      produtosDaAba.size,
      [...produtosDaAba].filter((k) => chavesComparativo.has(k)).length,
      contagem,
      ref,
    );

    // 6 · Clientes, mês a mês
    let divergentes = 0;
    for (const c of e.classificacaoCarteira) {
      const daAba = p.clientesCarteira.get(c.chave)?.totalGeral ?? 0;
      if (!iguais(daAba, c.faturamentoPorPeriodo[p.nome] ?? 0)) divergentes++;
    }
    for (const f of e.funcionarios) {
      const daAba = p.funcionarios.get(f.chave)?.totalGeral ?? 0;
      if (!iguais(daAba, f.faturamentoPorPeriodo[p.nome] ?? 0)) divergentes++;
    }
    conferir(6, escopo, "Clientes com valor divergente da aba", 0, divergentes, contagem, ref);

    // 8 · Linha do comparativo mês a mês
    const linha = e.evolucaoPeriodos[i];
    conferir(8, escopo, "Comparativo mensal: faturamento", p.totalFaturamento, linha?.faturamento ?? NaN, brl, ref);
    conferir(8, escopo, "Comparativo mensal: clientes", p.clientesCarteira.size, linha?.clientes ?? NaN, contagem, ref);

    const falhas = itens.slice(antes).filter((it) => !it.ok);
    abas.push({
      aba: p.sheetName,
      periodo: p.nome,
      registrosVenda: l.registrosVenda,
      clientes: p.clientes.size,
      clientesCarteira: p.clientesCarteira.size,
      funcionarios: p.funcionarios.size,
      produtosDistintos: l.produtosDistintos,
      quantidadeTotal: l.quantidadeTotal,
      faturamentoApurado: p.totalFaturamento,
      faturamentoCarteira: p.faturamentoCarteira,
      faturamentoFuncionarios: p.faturamentoFuncionarios,
      subtotaisImpressos: l.subtotaisImpressos,
      totalImpresso: l.totalImpresso?.valor ?? null,
      origemTotalImpresso: l.totalImpresso?.origem ?? null,
      valorDescartado,
      linhasDescartadas: l.linhasDescartadas.length,
      divergencias: falhas.map((f) => `${f.descricao}: ${f.detalhe}`),
      ok: falhas.length === 0,
    });
  });

  // ---- Whole window --------------------------------------------------------------
  const escopo = "Consolidado";
  const ref = "Conferência!A1";
  const faturamento = somarValores(periodos.map((p) => p.totalFaturamento));
  const faturamentoCarteira = somarValores(periodos.map((p) => p.faturamentoCarteira));
  const faturamentoFuncionarios = somarValores(periodos.map((p) => p.faturamentoFuncionarios));

  // 2 · Faturamento consolidado
  conferir(2, escopo, "Soma das abas = total da janela", faturamento, e.evolucaoConsolidada.faturamentoTotal, brl, ref);
  conferir(
    2,
    escopo,
    "Soma das abas = soma dos produtos na janela",
    faturamento,
    somarValores(e.comparativoProdutos.map((c) => c.valorTotalPeriodo)),
    brl,
    ref,
  );
  conferir(
    2,
    escopo,
    "Soma das abas = carteira + funcionários na janela",
    faturamento,
    somarValores([
      ...e.classificacaoCarteira.flatMap((c) => Object.values(c.faturamentoPorPeriodo)),
      ...e.funcionarios.map((f) => f.faturamentoTotal),
    ]),
    brl,
    ref,
  );

  // 3 · Registros
  const registrosVenda = periodos.reduce((s, p) => s + p.leituraAba.registrosVenda, 0);
  conferir(
    3,
    escopo,
    "Registros de venda na janela",
    registrosVenda,
    periodos.reduce((s, p) => s + p.produtosLidos, 0),
    contagem,
    ref,
  );

  // 4 · Clientes
  const carteiraDistintos = new Set(periodos.flatMap((p) => [...p.clientesCarteira.keys()]));
  const funcionariosDistintos = new Set(periodos.flatMap((p) => [...p.funcionarios.keys()]));
  conferir(
    4,
    escopo,
    "Clientes distintos na janela (sem funcionários)",
    carteiraDistintos.size,
    e.classificacaoCarteira.length,
    contagem,
    ref,
  );
  conferir(
    4,
    escopo,
    "Clientes distintos no resumo",
    carteiraDistintos.size,
    e.evolucaoConsolidada.clientesDistintos,
    contagem,
    ref,
  );

  // Funcionários fora de toda análise de clientes
  const ehFunc = (chave: string) => funcionariosDistintos.has(chave);
  const vazamentos = [
    ...e.classificacaoCarteira.filter((c) => ehFunc(c.chave)).map((c) => `carteira: ${c.nome}`),
    ...e.carteira.novos.filter((c) => c.funcionario).map((c) => `novos: ${c.nome}`),
    ...e.carteira.retornados.filter((c) => c.funcionario).map((c) => `retornados: ${c.nome}`),
    ...e.carteira.ativos.filter((c) => c.funcionario).map((c) => `ativos: ${c.nome}`),
    ...e.transicoes.flatMap((t) =>
      [t.novos, t.retornados, t.mantidos, t.perdidos].flatMap((g) =>
        g.clientes.filter((c) => ehFunc(c.chave)).map((c) => `${t.de} → ${t.para}: ${c.nome}`),
      ),
    ),
  ];
  conferirCondicao(
    4,
    escopo,
    "Funcionários (*) fora de clientes novos, perdidos e carteira",
    vazamentos.length === 0 && ![...carteiraDistintos].some(ehFunc),
    `funcionário contado como cliente — ${vazamentos.join("; ")}`,
    ref,
  );

  // 5 · Produtos
  const produtosDistintos = new Set(
    periodos.flatMap((p) =>
      [...p.clientes.values()].flatMap((c) => c.produtos.map((x) => `${x.codigo.trim()}|${normalizarTexto(x.produto)}`)),
    ),
  );
  conferir(5, escopo, "Produtos distintos na janela", produtosDistintos.size, e.comparativoProdutos.length, contagem, ref);
  const quantidadeTotal = somarValores(periodos.map((p) => p.leituraAba.quantidadeTotal));
  conferir(
    5,
    escopo,
    "Quantidade vendida na janela",
    quantidadeTotal,
    somarValores(e.comparativoProdutos.flatMap((c) => Object.values(c.qtdPorPeriodo))),
    qtd,
    ref,
  );

  // 8 · Comparativos mensais
  e.transicoes.forEach((t, i) => {
    const de = periodos[i];
    const para = periodos[i + 1];
    const esc = `${t.de} → ${t.para}`;
    conferir(8, esc, "Faturamento do mês anterior", de.totalFaturamento, t.faturamentoDe, brl, ref);
    conferir(8, esc, "Faturamento do mês", para.totalFaturamento, t.faturamentoPara, brl, ref);
    conferir(
      8,
      esc,
      "Ponte de receita fecha na variação do mês",
      t.diferencaFaturamento,
      somarValores([
        t.ponte.novos,
        t.ponte.retornados,
        t.ponte.expansao,
        t.ponte.contracao,
        t.ponte.perdidos,
        t.ponte.funcionarios,
      ]),
      brl,
      ref,
    );
    conferir(
      8,
      esc,
      "Novos + retornados + mantidos = clientes do mês",
      para.clientesCarteira.size,
      t.novos.total + t.retornados.total + t.mantidos.total,
      contagem,
      ref,
    );
    conferir(
      8,
      esc,
      "Mantidos + perdidos = clientes do mês anterior",
      de.clientesCarteira.size,
      t.mantidos.total + t.perdidos.total,
      contagem,
      ref,
    );
  });

  // Same month in two tabs, or two tabs with the very same records: the window
  // would count those sales twice.
  for (const r of e.periodosRepetidos) {
    conferirCondicao(
      2,
      escopo,
      `Mês ${r.periodo} em uma única aba`,
      false,
      `o mesmo mês aparece nas abas ${r.abas.map((a) => `"${a}"`).join(", ")} — as vendas seriam contadas mais de uma vez`,
      ref,
    );
  }
  const assinaturas = new Map<string, string>();
  for (const p of periodos) {
    if (p.leituraAba.registrosVenda === 0) continue;
    const assinatura = JSON.stringify(
      [...p.clientes.values()]
        .flatMap((c) => c.produtos.map((x) => [c.chave, x.codigo, x.produto, x.quantidade, x.valorTotal]))
        .sort(),
    );
    const igual = assinaturas.get(assinatura);
    if (igual) {
      conferirCondicao(
        2,
        escopo,
        `Abas "${igual}" e "${p.sheetName}" com registros diferentes`,
        false,
        `as abas "${igual}" e "${p.sheetName}" têm exatamente os mesmos registros — possível aba duplicada`,
        ref,
      );
    } else {
      assinaturas.set(assinatura, p.sheetName);
    }
  }

  return {
    ok: itens.every((it) => it.ok),
    abas,
    itens,
    consolidado: {
      abas: periodos.length,
      registrosVenda,
      faturamento,
      faturamentoCarteira,
      faturamentoFuncionarios,
      clientesDistintos: carteiraDistintos.size,
      funcionariosDistintos: funcionariosDistintos.size,
      produtosDistintos: produtosDistintos.size,
      quantidadeTotal,
    },
    validationReport: validator.getReport(),
  };
}
