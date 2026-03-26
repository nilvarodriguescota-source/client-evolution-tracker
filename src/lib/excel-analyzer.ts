import * as XLSX from "xlsx";

export interface RawRow {
  Cliente: string;
  Produto: string;
  "Código do Produto": string | number;
  "Valor Unitário": number;
  Quantidade: number;
  "Valor Total por Produto": number;
  "Total Geral": number;
}

export interface ProductDetail {
  produto: string;
  codigo: string;
  valorUnitario: number;
  quantidade: number;
  valorTotal: number;
  validado: boolean;
}

export interface ClientData {
  nome: string;
  produtos: ProductDetail[];
  totalGeral: number;
}

export interface PeriodData {
  nome: string;
  clientes: Map<string, ClientData>;
  totalFaturamento: number;
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

export interface AnalysisResult {
  periodoAnterior: PeriodData;
  periodoAtual: PeriodData;
  clientesNovos: ClientData[];
  clientesPerdidos: ClientData[];
  clientesMantidos: string[];
  faturamentoNovos: number;
  faturamentoPerdidos: number;
  variacaoFaturamento: number;
  comparativoProdutos: ProductComparison[];
  inconsistencias: string[];
}

function parsePeriod(sheet: XLSX.WorkSheet, nome: string): PeriodData {
  const rows = XLSX.utils.sheet_to_json<RawRow>(sheet);
  const clientes = new Map<string, ClientData>();

  for (const row of rows) {
    const clienteNome = String(row.Cliente ?? "").trim();
    if (!clienteNome) continue;

    const produto: ProductDetail = {
      produto: String(row.Produto ?? ""),
      codigo: String(row["Código do Produto"] ?? ""),
      valorUnitario: Number(row["Valor Unitário"]) || 0,
      quantidade: Number(row.Quantidade) || 0,
      valorTotal: Number(row["Valor Total por Produto"]) || 0,
      validado: false,
    };

    const calculado = +(produto.valorUnitario * produto.quantidade).toFixed(2);
    produto.validado = Math.abs(calculado - produto.valorTotal) < 0.02;

    if (!clientes.has(clienteNome)) {
      clientes.set(clienteNome, {
        nome: clienteNome,
        produtos: [],
        totalGeral: 0,
      });
    }
    const c = clientes.get(clienteNome)!;
    // avoid duplicate product entries
    const existing = c.produtos.find(
      (p) => p.codigo === produto.codigo && p.produto === produto.produto
    );
    if (!existing) {
      c.produtos.push(produto);
    }
  }

  // recalculate totals
  let totalFaturamento = 0;
  for (const c of clientes.values()) {
    c.totalGeral = c.produtos.reduce((s, p) => s + p.valorTotal, 0);
    totalFaturamento += c.totalGeral;
  }

  return { nome, clientes, totalFaturamento };
}

export function analyzeWorkbook(buffer: ArrayBuffer): AnalysisResult {
  const wb = XLSX.read(buffer, { type: "array" });

  if (wb.SheetNames.length < 2) {
    throw new Error("A planilha precisa ter pelo menos 2 abas (períodos).");
  }

  const sheetNames = wb.SheetNames;
  const anteriorName = sheetNames[sheetNames.length - 2];
  const atualName = sheetNames[sheetNames.length - 1];

  const periodoAnterior = parsePeriod(wb.Sheets[anteriorName], anteriorName);
  const periodoAtual = parsePeriod(wb.Sheets[atualName], atualName);

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

  const faturamentoNovos = clientesNovos.reduce((s, c) => s + c.totalGeral, 0);
  const faturamentoPerdidos = clientesPerdidos.reduce((s, c) => s + c.totalGeral, 0);

  const variacaoFaturamento =
    periodoAnterior.totalFaturamento > 0
      ? ((periodoAtual.totalFaturamento - periodoAnterior.totalFaturamento) /
          periodoAnterior.totalFaturamento) *
        100
      : 0;

  // Product comparison
  const produtosAnterior = new Map<string, { qtd: number; valor: number; codigo: string }>();
  const produtosAtual = new Map<string, { qtd: number; valor: number; codigo: string }>();

  for (const c of periodoAnterior.clientes.values()) {
    for (const p of c.produtos) {
      const key = p.produto;
      const existing = produtosAnterior.get(key) || { qtd: 0, valor: 0, codigo: p.codigo };
      existing.qtd += p.quantidade;
      existing.valor += p.valorTotal;
      produtosAnterior.set(key, existing);
    }
  }
  for (const c of periodoAtual.clientes.values()) {
    for (const p of c.produtos) {
      const key = p.produto;
      const existing = produtosAtual.get(key) || { qtd: 0, valor: 0, codigo: p.codigo };
      existing.qtd += p.quantidade;
      existing.valor += p.valorTotal;
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
      variacaoQtd: ant.qtd > 0 ? ((atu.qtd - ant.qtd) / ant.qtd) * 100 : atu.qtd > 0 ? 100 : 0,
      valorAnterior: ant.valor,
      valorAtual: atu.valor,
      variacaoValor: ant.valor > 0 ? ((atu.valor - ant.valor) / ant.valor) * 100 : atu.valor > 0 ? 100 : 0,
    });
  }
  comparativoProdutos.sort((a, b) => b.valorAtual - a.valorAtual);

  // Inconsistencies
  const inconsistencias: string[] = [];
  for (const period of [periodoAnterior, periodoAtual]) {
    for (const c of period.clientes.values()) {
      for (const p of c.produtos) {
        if (!p.validado) {
          inconsistencias.push(
            `${period.nome} → ${c.nome} → ${p.produto}: Valor Total (${p.valorTotal.toFixed(2)}) ≠ Unitário (${p.valorUnitario.toFixed(2)}) × Qtd (${p.quantidade})`
          );
        }
      }
    }
  }

  return {
    periodoAnterior,
    periodoAtual,
    clientesNovos,
    clientesPerdidos,
    clientesMantidos,
    faturamentoNovos,
    faturamentoPerdidos,
    variacaoFaturamento,
    comparativoProdutos,
    inconsistencias,
  };
}
