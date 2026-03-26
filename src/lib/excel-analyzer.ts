import * as XLSX from "xlsx";

export interface ProductDetail {
  produto: string;
  codigo: string;
  unidade: string;
  valorUnitario: number;
  quantidade: number;
  valorTotal: number;
  validado: boolean;
}

export interface ClientData {
  nome: string;
  codigoCliente: string;
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
  const rows: (string | number | null)[][] = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: null,
  });

  const clientes = new Map<string, ClientData>();
  let currentClient: ClientData | null = null;

  for (const row of rows) {
    if (!row || row.length === 0) continue;

    const firstCell = String(row[0] ?? "").trim();

    // Detect client header: "Cliente: XXXXX - NOME - ..."
    if (firstCell.startsWith("Cliente:")) {
      const match = firstCell.match(/Cliente:\s*(\d+)\s*-\s*(.+?)(?:\s*-\s*Data|$)/);
      if (match) {
        const codigoCliente = match[1];
        const nomeCliente = match[2].trim();
        currentClient = {
          nome: nomeCliente,
          codigoCliente,
          produtos: [],
          totalGeral: 0,
        };
        clientes.set(nomeCliente, currentClient);
      }
      continue;
    }

    // Skip header row, empty rows, totals section
    if (
      firstCell === "Código" ||
      firstCell === "" ||
      firstCell === "Totais por unidade:" ||
      firstCell.startsWith("Quantidade total") ||
      firstCell.startsWith("Gerado em")
    ) {
      continue;
    }

    // Check if it's a product row (first cell is a numeric code)
    const codigo = firstCell;
    const isProductRow = /^\d+$/.test(codigo) && currentClient;

    if (isProductRow && currentClient) {
      const descricao = String(row[1] ?? "");
      const unidade = String(row[2] ?? "");
      const quantidade = Number(row[3]) || 0;
      const precoMedio = Number(row[4]) || 0;
      const vendaBruta = Number(row[5]) || 0;

      const calculado = +(precoMedio * quantidade).toFixed(2);
      const validado = Math.abs(calculado - vendaBruta) < 0.1;

      currentClient.produtos.push({
        produto: descricao,
        codigo,
        unidade,
        valorUnitario: precoMedio,
        quantidade,
        valorTotal: vendaBruta,
        validado,
      });
      continue;
    }

    // Subtotal row: first cell empty, col[3] has qty total, col[5] has value total
    if (!firstCell && currentClient && row[3] && row[5]) {
      currentClient.totalGeral = Number(row[5]) || 0;
      continue;
    }
  }

  // Recalculate totals from products for consistency
  let totalFaturamento = 0;
  for (const c of clientes.values()) {
    const sumProdutos = c.produtos.reduce((s, p) => s + p.valorTotal, 0);
    if (c.totalGeral === 0) c.totalGeral = sumProdutos;
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
            `${period.nome} → ${c.nome} → ${p.produto}: Valor (${p.valorTotal.toFixed(2)}) ≠ Preço médio (${p.valorUnitario.toFixed(2)}) × Qtd (${p.quantidade})`
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
