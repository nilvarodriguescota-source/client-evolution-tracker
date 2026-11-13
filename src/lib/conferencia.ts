/**
 * Conferência — o que a planilha contém, aba por aba.
 *
 * A planilha é a única fonte oficial dos números. Este resumo não compara a
 * planilha com nada e não corrige nem completa nenhum valor: apenas conta e soma
 * o que foi encontrado nas linhas de venda de cada aba, para que se possa
 * conferir os números com a própria planilha.
 *
 *   - registros de venda: uma linha de produto de um cliente;
 *   - quantidade vendida: soma da coluna "Qtd. venda" de cada produto dessas
 *     linhas, sem as linhas de total (do cliente, por unidade ou da aba);
 *   - faturamento: soma da "Venda bruta" dessas mesmas linhas.
 */

import { somarValores } from "./numero-br";
import { normalizarTexto } from "./sheet-discovery";
import type { PeriodData } from "./excel-analyzer";

export interface ConferenciaAba {
  aba: string;
  periodo: string;
  registrosVenda: number;
  /** Compradores do mês (clientes + funcionários) */
  clientes: number;
  clientesCarteira: number;
  funcionarios: number;
  produtosDistintos: number;
  /** Unidades vendidas: soma da coluna "Qtd. venda" de cada produto */
  quantidadeTotal: number;
  /** Soma da "Venda bruta" das linhas de venda */
  faturamento: number;
  faturamentoCarteira: number;
  faturamentoFuncionarios: number;
  /** Linhas com valor que não puderam ser lidas como venda nem como total */
  linhasDescartadas: number;
  valorDescartado: number;
}

export interface Conferencia {
  abas: ConferenciaAba[];
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
}

const chaveProduto = (codigo: string, descricao: string) =>
  `${codigo.trim()}|${normalizarTexto(descricao)}`;

export function resumirPlanilha(periodos: PeriodData[]): Conferencia {
  const abas: ConferenciaAba[] = periodos.map((p) => {
    const l = p.leituraAba;
    return {
      aba: p.sheetName,
      periodo: p.nome,
      registrosVenda: l.registrosVenda,
      clientes: p.clientes.size,
      clientesCarteira: p.clientesCarteira.size,
      funcionarios: p.funcionarios.size,
      produtosDistintos: l.produtosDistintos,
      quantidadeTotal: l.quantidadeTotal,
      faturamento: p.totalFaturamento,
      faturamentoCarteira: p.faturamentoCarteira,
      faturamentoFuncionarios: p.faturamentoFuncionarios,
      linhasDescartadas: l.linhasDescartadas.length,
      valorDescartado: somarValores(l.linhasDescartadas.map((d) => d.valor)),
    };
  });

  const produtos = new Set(
    periodos.flatMap((p) =>
      [...p.clientes.values()].flatMap((c) => c.produtos.map((x) => chaveProduto(x.codigo, x.produto))),
    ),
  );

  return {
    abas,
    consolidado: {
      abas: periodos.length,
      registrosVenda: abas.reduce((s, a) => s + a.registrosVenda, 0),
      faturamento: somarValores(abas.map((a) => a.faturamento)),
      faturamentoCarteira: somarValores(abas.map((a) => a.faturamentoCarteira)),
      faturamentoFuncionarios: somarValores(abas.map((a) => a.faturamentoFuncionarios)),
      clientesDistintos: new Set(periodos.flatMap((p) => [...p.clientesCarteira.keys()])).size,
      funcionariosDistintos: new Set(periodos.flatMap((p) => [...p.funcionarios.keys()])).size,
      produtosDistintos: produtos.size,
      quantidadeTotal: somarValores(abas.map((a) => a.quantidadeTotal)),
    },
  };
}
