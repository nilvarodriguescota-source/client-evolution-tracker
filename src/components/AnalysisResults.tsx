import { useState } from "react";
import {
  Users,
  UserPlus,
  UserMinus,
  DollarSign,
  TrendingUp,
  TrendingDown,
  Package,
  AlertTriangle,
  ShieldCheck,
  ShieldAlert,
  CheckCircle2,
  XCircle,
  Info,
  CalendarRange,
  RotateCcw,
  FileSpreadsheet,
  ArrowRight,
  Minus,
} from "lucide-react";
import { KpiCard } from "./KpiCard";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { Badge } from "@/components/ui/badge";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  Legend,
  LabelList,
  ReferenceLine,
} from "recharts";
import type {
  AnalysisResult,
  ClientData,
  StatusCarteira,
  ClienteCarteira,
  TransicaoPeriodo,
  ClienteMovimento,
  Tendencia,
} from "@/lib/excel-analyzer";
import type { ValidationDiagnostic, ValidationSeverity } from "@/lib/excel-validator";

interface Props {
  result: AnalysisResult;
}

const fmt = (v: number) =>
  v.toLocaleString("pt-BR", {
    style: "currency",
    currency: "BRL",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

/** Units sold, Brazilian format, without trailing zeros. */
const fmtQtd = (v: number) => v.toLocaleString("pt-BR", { maximumFractionDigits: 4 });

export function AnalysisResults({ result }: Props) {
  const [tab, setTab] = useState("resumo");

  const periodNames = result.periodos.map((p) => p.nome);

  // One series per month of the window instead of just anterior/atual.
  const chartData = result.comparativoProdutos.slice(0, 10).map((p) => ({
    nome: p.produto.length > 18 ? p.produto.slice(0, 16) + "…" : p.produto,
    ...Object.fromEntries(periodNames.map((n) => [n, p.valorPorPeriodo[n] ?? 0])),
  }));

  // One measure per chart: revenue and client movement are read on their own
  // scale instead of being squeezed onto two axes of the same plot.
  const faturamentoChartData = result.evolucaoPeriodos.map((p) => ({
    nome: p.nome,
    Faturamento: p.faturamento,
    Clientes: p.clientes,
  }));

  // One group per transition: who stayed, who came in, who left.
  const movimentacaoChartData = result.transicoes.map((t) => ({
    nome: `${t.de} → ${t.para}`,
    Mantidos: t.mantidos.total,
    Retornados: t.retornados.total,
    Novos: t.novos.total,
    Perdidos: -t.perdidos.total,
  }));

  const ec = result.evolucaoConsolidada;
  const ultimaTransicao = result.transicoes[result.transicoes.length - 1] ?? null;

  // Tabs that were read but yielded nothing — the shape a lost month takes.
  const periodosVazios = result.periodos.filter((p) => p.clientes.size === 0);

  const vr = result.validationReport;
  const hasAlerts = result.inconsistencias.length > 0 || vr.diagnostics.length > 0;
  const errorCount = vr.diagnostics.filter((d) => d.severity === "error").length;
  const warningCount = vr.diagnostics.filter((d) => d.severity === "warning").length;

  return (
    <div className="space-y-6">
      {/* KPI cards — the current month measured against the one before it */}
      <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
        <KpiCard
          title={`Faturamento · ${result.periodoAtual.nome}`}
          value={fmt(result.periodoAtual.totalFaturamento)}
          subtitle={`${pct(ec.variacaoTotal)} desde ${ec.primeiro}`}
          icon={DollarSign}
          trend={ultimaTransicao?.variacaoFaturamento ?? undefined}
          color="primary"
        />
        <KpiCard
          title={`Clientes · ${result.periodoAtual.nome}`}
          value={String(result.periodoAtual.clientesCarteira.size)}
          subtitle={`${result.periodoAnterior.clientesCarteira.size} no mês anterior · ${ec.clientesDistintos} distintos na janela`}
          icon={Users}
          trend={ultimaTransicao?.variacaoClientes ?? undefined}
          color="primary"
        />
        <KpiCard
          title="Entraram no mês"
          value={String(
            ultimaTransicao ? ultimaTransicao.novos.total + ultimaTransicao.retornados.total : 0,
          )}
          subtitle={
            ultimaTransicao
              ? `${ultimaTransicao.novos.total} novos · ${ultimaTransicao.retornados.total} retornos · ${fmt(ultimaTransicao.novos.faturamento + ultimaTransicao.retornados.faturamento)}`
              : "sem mês anterior para comparar"
          }
          icon={UserPlus}
          color="success"
        />
        <KpiCard
          title="Saíram no mês"
          value={String(ultimaTransicao ? ultimaTransicao.perdidos.total : 0)}
          subtitle={
            ultimaTransicao
              ? `${fmt(ultimaTransicao.perdidos.faturamento)} deixaram de faturar · ${result.carteira.perdidos.length} há 2+ meses sem comprar`
              : "sem mês anterior para comparar"
          }
          icon={UserMinus}
          color="destructive"
        />
      </div>

      {/* Analysis window */}
      <Card className="border-primary/40 bg-primary/5">
        <CardContent className="flex flex-wrap items-center gap-x-3 gap-y-2 py-3">
          <CalendarRange className="h-5 w-5 shrink-0 text-primary" />
          <span className="text-sm font-medium">
            Janela de análise da carteira: {result.periodos.length}{" "}
            {result.periodos.length === 1 ? "mês" : "meses"}
          </span>
          <div className="flex flex-wrap items-center gap-1.5">
            {result.periodos.map((p, i) => (
              <Badge
                key={p.sheetName}
                variant="outline"
                title={`Aba "${p.sheetName}" · ${p.linhasLidas} linhas · ${p.clientesCarteira.size} clientes · ${p.funcionarios.size} funcionários · ${p.produtosLidos} registros de venda`}
                className={
                  i === result.periodos.length - 1
                    ? "border-primary/40 bg-primary/10 text-primary"
                    : "text-muted-foreground"
                }
              >
                {p.nome}
                {i === result.periodos.length - 1 && " (atual)"}
                <span className="ml-1.5 opacity-70">
                  {p.clientesCarteira.size} cli · {p.produtosLidos} prod
                </span>
              </Badge>
            ))}
          </div>
          <span className="text-xs text-muted-foreground">
            Todos os {result.periodos.length}{" "}
            {result.periodos.length === 1 ? "mês entra" : "meses entram"} nos cálculos,
            produtos, carteira e validação. A variação mês a mês do topo compara{" "}
            {result.periodoAnterior.nome} × {result.periodoAtual.nome}.
          </span>
        </CardContent>
      </Card>

      {/* Tabs left out of the analysis — never silently */}
      {(result.leitura.abasIgnoradas.length > 0 || periodosVazios.length > 0) && (
        <Card className="border-yellow-500/50 bg-yellow-500/5">
          <CardContent className="space-y-1 py-3 text-sm">
            <div className="flex items-center gap-2 font-medium text-yellow-700 dark:text-yellow-500">
              <AlertTriangle className="h-4 w-4 shrink-0" />
              Nem toda aba do arquivo virou período
            </div>
            {result.leitura.abasIgnoradas.map((a) => (
              <p key={a.sheetName} className="text-muted-foreground">
                Aba <span className="font-medium text-foreground">"{a.sheetName}"</span> não foi
                lida como período: {a.motivo}.
              </p>
            ))}
            {periodosVazios.map((p) => (
              <p key={p.sheetName} className="text-muted-foreground">
                Aba <span className="font-medium text-foreground">"{p.sheetName}"</span> tem{" "}
                {p.linhasLidas} linhas, mas nenhum cliente foi reconhecido — o período entra
                zerado. Confira o layout dessa aba.
              </p>
            ))}
          </CardContent>
        </Card>
      )}

      {/* Validation summary banner */}
      <Card className={vr.valid ? "border-accent/50 bg-accent/5" : "border-destructive/50 bg-destructive/5"}>
        <CardContent className="flex items-center gap-3 py-3">
          {vr.valid ? (
            <ShieldCheck className="h-5 w-5 text-accent" />
          ) : (
            <ShieldAlert className="h-5 w-5 text-destructive" />
          )}
          <div className="flex-1 text-sm">
            <span className="font-medium">
              {vr.valid ? "Dados validados com sucesso" : "Inconsistências detectadas nos dados"}
            </span>
            <span className="ml-2 text-muted-foreground">
              {vr.passedChecks}/{vr.totalChecks} verificações aprovadas
              {errorCount > 0 && ` · ${errorCount} erro${errorCount > 1 ? "s" : ""}`}
              {warningCount > 0 && ` · ${warningCount} aviso${warningCount > 1 ? "s" : ""}`}
            </span>
            <p className="text-xs text-muted-foreground">
              Dados da planilha: {result.conferencia.abas.length}{" "}
              {result.conferencia.abas.length === 1 ? "mês" : "meses"} ·{" "}
              {fmt(result.conferencia.consolidado.faturamento)} em{" "}
              {result.conferencia.consolidado.registrosVenda} registros de venda ·{" "}
              {fmtQtd(result.conferencia.consolidado.quantidadeTotal)} unidades vendidas.
            </p>
          </div>
        </CardContent>
      </Card>

      {/* Tabs */}
      <Tabs value={tab} onValueChange={setTab}>
        <TabsList className="w-full justify-start">
          <TabsTrigger value="resumo">Resumo</TabsTrigger>
          <TabsTrigger value="novos">
            Clientes Novos ({result.carteira.novos.length})
          </TabsTrigger>
          <TabsTrigger value="perdidos">
            Clientes Perdidos ({result.carteira.perdidos.length})
          </TabsTrigger>
          <TabsTrigger value="produtos">Comparativo Produtos</TabsTrigger>
          <TabsTrigger value="carteira">
            Carteira ({result.periodos.length}m)
          </TabsTrigger>
          {hasAlerts && (
            <TabsTrigger value="alertas" className="text-destructive">
              Alertas ({result.inconsistencias.length})
            </TabsTrigger>
          )}
          <TabsTrigger value="validacao">
            Validação
          </TabsTrigger>
        </TabsList>

        {/* Resumo */}
        <TabsContent value="resumo" className="space-y-6">
          {/* 1 · The whole window as one movement */}
          <EvolucaoJanela result={result} />

          {/* 2 · Month against month */}
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2 text-base">
                <CalendarRange className="h-4 w-4 text-primary" />
                Comparativo mês a mês · {result.periodos.length}{" "}
                {result.periodos.length === 1 ? "mês" : "meses"}
              </CardTitle>
              <p className="text-xs text-muted-foreground">
                Cada mês medido contra o mês anterior. Novos, retornados, mantidos e
                perdidos são apurados contra todo o histórico da janela, não só contra
                o mês anterior.
              </p>
            </CardHeader>
            <CardContent className="space-y-6">
              <div className="overflow-x-auto">
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Mês</TableHead>
                      <TableHead className="text-right">Faturamento</TableHead>
                      <TableHead className="text-right">Var. Faturamento</TableHead>
                      <TableHead>Movimentação</TableHead>
                      <TableHead className="text-right">Qtd. Vendida</TableHead>
                      <TableHead className="text-right">Clientes</TableHead>
                      <TableHead className="text-right">Var. Clientes</TableHead>
                      <TableHead className="text-right">Ticket Médio</TableHead>
                      <TableHead className="text-right">Novos</TableHead>
                      <TableHead className="text-right">Retornaram</TableHead>
                      <TableHead className="text-right">Mantidos</TableHead>
                      <TableHead className="text-right">Perdidos</TableHead>
                      <TableHead className="text-right">Retenção</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {result.evolucaoPeriodos.map((p, i) => (
                      <TableRow
                        key={p.nome}
                        className={i === result.evolucaoPeriodos.length - 1 ? "bg-primary/5" : ""}
                      >
                        <TableCell className="font-medium capitalize">
                          {p.nome}
                          {i === 0 && (
                            <span className="ml-2 text-xs text-muted-foreground">(base)</span>
                          )}
                          {i === result.evolucaoPeriodos.length - 1 && (
                            <span className="ml-2 text-xs text-muted-foreground">(atual)</span>
                          )}
                        </TableCell>
                        <TableCell className="text-right font-medium">{fmt(p.faturamento)}</TableCell>
                        <TableCell className="text-right">
                          <VariationCell value={p.variacaoFaturamento} />
                        </TableCell>
                        <TableCell>
                          {p.movimentacao === null ? (
                            <span className="text-xs text-muted-foreground">—</span>
                          ) : (
                            <MovimentacaoBadge tendencia={p.movimentacao} />
                          )}
                        </TableCell>
                        <TableCell className="text-right" title={
                          p.variacaoQuantidade === null
                            ? undefined
                            : `${pct(p.variacaoQuantidade)} vs mês anterior`
                        }>
                          {fmtQtd(p.quantidade)}
                        </TableCell>
                        <TableCell className="text-right">{p.clientes}</TableCell>
                        <TableCell className="text-right">
                          <VariationCell value={p.variacaoClientes} />
                        </TableCell>
                        <TableCell className="text-right">{fmt(p.ticketMedio)}</TableCell>
                        <ContagemCell valor={p.novos} sinal="+" tom="text-accent" />
                        <ContagemCell valor={p.retornados} sinal="+" tom="text-primary" />
                        <TableCell className="text-right">
                          {p.mantidos === null ? (
                            <span className="text-xs text-muted-foreground">—</span>
                          ) : (
                            p.mantidos
                          )}
                        </TableCell>
                        <ContagemCell valor={p.perdidos} sinal="−" tom="text-destructive" />
                        <TableCell className="text-right">
                          {p.taxaRetencao === null ? (
                            <span className="text-xs text-muted-foreground">—</span>
                          ) : (
                            `${p.taxaRetencao.toFixed(0)}%`
                          )}
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </div>

              <div>
                <p className="mb-2 text-xs font-medium uppercase tracking-wide text-muted-foreground">
                  Faturamento por mês
                </p>
                <ResponsiveContainer width="100%" height={240}>
                  <BarChart data={faturamentoChartData} margin={{ top: 20 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="hsl(220 15% 88%)" vertical={false} />
                    <XAxis dataKey="nome" tick={{ fontSize: 12 }} tickLine={false} />
                    {/* wide enough for "R$ 1,4 mil" to stay on one line */}
                    <YAxis tick={{ fontSize: 11 }} tickLine={false} axisLine={false} width={96}
                      tickFormatter={(v: number) => fmtCurto(v)} />
                    <Tooltip
                      cursor={{ fill: "hsl(220 15% 95%)" }}
                      formatter={(v: number) => [fmt(v), "Faturamento"]}
                      contentStyle={{ borderRadius: "8px", border: "1px solid hsl(220 15% 88%)" }}
                    />
                    <Bar dataKey="Faturamento" fill="hsl(217 91% 50%)" radius={[4, 4, 0, 0]}>
                      <LabelList
                        dataKey="Faturamento"
                        position="top"
                        formatter={(v: number) => fmtCurto(v)}
                        style={{ fontSize: 11, fill: "hsl(220 9% 46%)" }}
                      />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>

              <div>
                <p className="mb-2 text-xs font-medium uppercase tracking-wide text-muted-foreground">
                  Clientes por mês
                </p>
                <ResponsiveContainer width="100%" height={200}>
                  <BarChart data={faturamentoChartData} margin={{ top: 20 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="hsl(220 15% 88%)" vertical={false} />
                    <XAxis dataKey="nome" tick={{ fontSize: 12 }} tickLine={false} />
                    <YAxis tick={{ fontSize: 11 }} tickLine={false} axisLine={false} width={40} allowDecimals={false} />
                    <Tooltip
                      cursor={{ fill: "hsl(220 15% 95%)" }}
                      formatter={(v: number) => [v, "Clientes"]}
                      contentStyle={{ borderRadius: "8px", border: "1px solid hsl(220 15% 88%)" }}
                    />
                    <Bar dataKey="Clientes" fill="hsl(160 70% 42%)" radius={[4, 4, 0, 0]}>
                      <LabelList
                        dataKey="Clientes"
                        position="top"
                        style={{ fontSize: 11, fill: "hsl(220 9% 46%)" }}
                      />
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </CardContent>
          </Card>

          {/* 3 · Client movement between months */}
          {result.transicoes.length > 0 && (
            <Card>
              <CardHeader>
                <CardTitle className="flex items-center gap-2 text-base">
                  <Users className="h-4 w-4 text-primary" /> Movimentação de clientes
                </CardTitle>
                <p className="text-xs text-muted-foreground">
                  Entrada e saída em cada passagem de mês. Acima do zero quem comprou;
                  abaixo, quem deixou de comprar.
                </p>
              </CardHeader>
              <CardContent>
                <ResponsiveContainer width="100%" height={260}>
                  <BarChart data={movimentacaoChartData} margin={{ top: 8 }} stackOffset="sign">
                    <CartesianGrid strokeDasharray="3 3" stroke="hsl(220 15% 88%)" vertical={false} />
                    <XAxis dataKey="nome" tick={{ fontSize: 11 }} tickLine={false} />
                    <YAxis tick={{ fontSize: 11 }} tickLine={false} axisLine={false} width={40} allowDecimals={false} />
                    <Tooltip
                      cursor={{ fill: "hsl(220 15% 95%)" }}
                      formatter={(v: number, n: string) => [Math.abs(v), n]}
                      contentStyle={{ borderRadius: "8px", border: "1px solid hsl(220 15% 88%)" }}
                    />
                    <Legend formatter={legendaEmTinta} />
                    <ReferenceLine y={0} stroke="hsl(220 15% 70%)" />
                    {/* the surface-coloured stroke leaves a 2px gap between segments */}
                    <Bar dataKey="Mantidos" stackId="mov" fill="hsl(217 91% 50%)" stroke="hsl(0 0% 100%)" strokeWidth={2} />
                    <Bar dataKey="Retornados" stackId="mov" fill="hsl(265 60% 55%)" stroke="hsl(0 0% 100%)" strokeWidth={2} />
                    <Bar dataKey="Novos" stackId="mov" fill="hsl(160 70% 42%)" stroke="hsl(0 0% 100%)" strokeWidth={2} radius={[4, 4, 0, 0]} />
                    <Bar dataKey="Perdidos" stackId="mov" fill="hsl(0 72% 51%)" stroke="hsl(0 0% 100%)" strokeWidth={2} radius={[0, 0, 4, 4]} />
                  </BarChart>
                </ResponsiveContainer>
              </CardContent>
            </Card>
          )}

          {/* 4 · One card per transition, with the revenue bridge */}
          {result.transicoes.length > 0 && (
            <Card>
              <CardHeader>
                <CardTitle className="flex items-center gap-2 text-base">
                  <TrendingUp className="h-4 w-4 text-primary" /> De um mês para o outro
                </CardTitle>
                <p className="text-xs text-muted-foreground">
                  O que explica a variação de cada mês: quanto entrou com clientes novos e
                  que retornaram, quanto os clientes mantidos cresceram ou recuaram, e
                  quanto saiu com quem parou de comprar.
                </p>
              </CardHeader>
              <CardContent className="space-y-4">
                {result.transicoes.map((t) => (
                  <TransicaoCard key={`${t.de}-${t.para}`} t={t} />
                ))}
              </CardContent>
            </Card>
          )}

          {/* 5 · Executive summary + products */}
          <div className="grid gap-6 lg:grid-cols-2">
            <Card>
              <CardHeader>
                <CardTitle className="flex items-center gap-2 text-base">
                  <TrendingUp className="h-4 w-4 text-accent" /> Resumo Executivo
                </CardTitle>
              </CardHeader>
              <CardContent className="space-y-3 text-sm">
                <p className="text-xs font-medium uppercase tracking-wide text-muted-foreground">
                  Faturamento por mês · {result.periodos.length}{" "}
                  {result.periodos.length === 1 ? "mês" : "meses"} analisados
                </p>
                {result.evolucaoPeriodos.map((p) => (
                  <div key={p.nome} className="flex items-center justify-between gap-2">
                    <span className="capitalize text-muted-foreground">{p.nome}</span>
                    <span className="flex items-center gap-2">
                      <span className="font-medium text-foreground">{fmt(p.faturamento)}</span>
                      {p.variacaoFaturamento === null ? (
                        <Badge variant="secondary" className="text-xs">base</Badge>
                      ) : (
                        <VariationBadge value={p.variacaoFaturamento} />
                      )}
                    </span>
                  </div>
                ))}
                <Row
                  label={`Variação total (${ec.primeiro} → ${ec.ultimo})`}
                  value={pct(ec.variacaoTotal)}
                  highlight
                />
                <Row label="Variação média mensal" value={pct(ec.variacaoMediaMensal)} />
                <hr className="border-border" />
                <p className="text-xs font-medium uppercase tracking-wide text-muted-foreground">
                  Clientes na janela
                </p>
                <Row label="Clientes distintos no período" value={String(ec.clientesDistintos)} />
                <Row
                  label={`Compraram em todos os ${ec.meses} meses`}
                  value={String(ec.clientesRecorrentes)}
                />
                <Row label="Entraram ao longo do período" value={String(ec.novosNoPeriodo)} />
                <Row label="Clientes ativos no mês atual" value={String(result.carteira.ativos.length)} />
                <Row label="Frequência reduzida (1 mês sem comprar)" value={String(result.carteira.frequenciaReduzida.length)} />
                <Row label="Perdidos (2+ meses sem comprar)" value={String(result.carteira.perdidos.length)} />
                <Row
                  label="Retenção média mês a mês"
                  value={ec.taxaRetencaoMedia === null ? "—" : `${ec.taxaRetencaoMedia.toFixed(1)}%`}
                />
                <hr className="border-border" />
                <p className="text-xs font-medium uppercase tracking-wide text-muted-foreground">
                  Último mês · {ultimaTransicao ? `${ultimaTransicao.de} × ${ultimaTransicao.para}` : result.periodoAtual.nome}
                </p>
                {ultimaTransicao && (
                  <>
                    <Row label="Clientes mantidos" value={String(ultimaTransicao.mantidos.total)} />
                    <Row label="Clientes novos" value={String(ultimaTransicao.novos.total)} />
                    <Row label="Clientes que retornaram" value={String(ultimaTransicao.retornados.total)} />
                    <Row label="Deixaram de comprar" value={String(ultimaTransicao.perdidos.total)} />
                    <Row label="Receita que entrou" value={fmt(ultimaTransicao.novos.faturamento + ultimaTransicao.retornados.faturamento)} />
                    <Row label="Receita que saiu" value={fmt(ultimaTransicao.perdidos.faturamento)} />
                  </>
                )}
              </CardContent>
            </Card>

            <Card>
              <CardHeader>
                <CardTitle className="flex items-center gap-2 text-base">
                  <Package className="h-4 w-4 text-primary" /> Top 10 Produtos por Faturamento
                </CardTitle>
                <p className="text-xs text-muted-foreground">
                  Uma série por mês da janela, na ordem cronológica.
                </p>
              </CardHeader>
              <CardContent>
                <ResponsiveContainer width="100%" height={320}>
                  <BarChart data={chartData} margin={{ bottom: 70 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="hsl(220 15% 88%)" vertical={false} />
                    <XAxis dataKey="nome" angle={-35} textAnchor="end" tick={{ fontSize: 11 }} interval={0} />
                    <YAxis tick={{ fontSize: 11 }} tickLine={false} axisLine={false} width={88}
                      tickFormatter={(v: number) => fmtCurto(v)} />
                    <Tooltip
                      cursor={{ fill: "hsl(220 15% 95%)" }}
                      formatter={(v: number) => fmt(v)}
                      contentStyle={{
                        borderRadius: "8px",
                        border: "1px solid hsl(220 15% 88%)",
                      }}
                    />
                    <Legend verticalAlign="top" height={28} formatter={legendaEmTinta} />
                    {periodNames.map((n, i) => (
                      <Bar
                        key={n}
                        dataKey={n}
                        fill={periodColor(i, periodNames.length)}
                        radius={[4, 4, 0, 0]}
                      />
                    ))}
                  </BarChart>
                </ResponsiveContainer>
              </CardContent>
            </Card>
          </div>
        </TabsContent>

        {/* Clientes Novos */}
        <TabsContent value="novos" className="space-y-4">
          <p className="text-sm text-muted-foreground">
            Entrada de clientes em cada passagem de mês da janela. É "novo" quem
            comprou pela primeira vez em toda a janela; quem já havia comprado antes e
            voltou depois de um intervalo aparece como retorno.
          </p>

          <EntradaSaidaPorMes result={result} tipo="entrada" />

          <h3 className="flex items-center gap-2 pt-2 text-sm font-semibold">
            <UserPlus className="h-4 w-4 text-accent" />
            Detalhe dos clientes novos em {result.periodoAtual.nome} (
            {result.carteira.novos.length})
          </h3>
          <ClientTable
            clients={result.carteira.novos}
            emptyMessage={`Nenhum cliente comprou pela primeira vez em ${result.periodoAtual.nome}.`}
          />
          {result.carteira.retornados.length > 0 && (
            <>
              <h3 className="flex items-center gap-2 pt-2 text-sm font-semibold">
                <RotateCcw className="h-4 w-4 text-primary" />
                Clientes que retornaram em {result.periodoAtual.nome} (
                {result.carteira.retornados.length})
              </h3>
              <ClientTable clients={result.carteira.retornados} emptyMessage="" />
            </>
          )}
        </TabsContent>

        {/* Clientes Perdidos */}
        <TabsContent value="perdidos" className="space-y-4">
          <p className="text-sm text-muted-foreground">
            Saída de clientes em cada passagem de mês. Quem deixou de comprar em um mês
            pode voltar no seguinte — a lista final é de quem está há 2 meses ou mais
            sem comprar.
          </p>

          <EntradaSaidaPorMes result={result} tipo="saida" />

          <h3 className="flex items-center gap-2 pt-2 text-sm font-semibold">
            <UserMinus className="h-4 w-4 text-destructive" />
            Clientes perdidos na janela ({result.carteira.perdidos.length})
          </h3>
          <p className="text-sm text-muted-foreground">
            Sem nenhuma compra há 2 meses ou mais. Quem faltou apenas no último mês
            aparece como frequência reduzida ({result.carteira.frequenciaReduzida.length})
            na aba Carteira.
          </p>
          <CarteiraTable
            clientes={result.carteira.perdidos}
            periodos={periodNames}
            emptyMessage="Nenhum cliente perdido nesta janela."
          />
        </TabsContent>

        {/* Comparativo Produtos */}
        <TabsContent value="produtos" className="space-y-4">
          <p className="text-sm text-muted-foreground">
            Faturamento de cada produto em todos os {result.periodos.length} meses da
            janela. A tendência e a variação do período comparam {ec.primeiro} com{" "}
            {ec.ultimo}; as colunas de quantidade seguem comparando{" "}
            {result.periodoAnterior.nome} × {result.periodoAtual.nome}.
          </p>
          <Card>
            <CardContent className="overflow-x-auto p-0">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>Produto</TableHead>
                    <TableHead>Código</TableHead>
                    {periodNames.map((n) => (
                      <TableHead key={n} className="text-right capitalize">
                        {n}
                      </TableHead>
                    ))}
                    <TableHead>Tendência</TableHead>
                    <TableHead className="text-right">Var. período</TableHead>
                    <TableHead className="text-right">Meses c/ venda</TableHead>
                    <TableHead className="text-right">Var. mês atual</TableHead>
                    <TableHead className="text-right">Qtd Anterior</TableHead>
                    <TableHead className="text-right">Qtd Atual</TableHead>
                    <TableHead className="text-right">Var. Qtd</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {result.comparativoProdutos.map((p) => (
                    <TableRow key={p.chave}>
                      <TableCell className="font-medium">{p.produto}</TableCell>
                      <TableCell className="text-muted-foreground">{p.codigo}</TableCell>
                      {periodNames.map((n, i) => (
                        <TableCell
                          key={n}
                          className={
                            i === periodNames.length - 1
                              ? "text-right font-medium"
                              : "text-right text-muted-foreground"
                          }
                        >
                          {fmt(p.valorPorPeriodo[n] ?? 0)}
                        </TableCell>
                      ))}
                      <TableCell>
                        <TendenciaBadge tendencia={p.tendencia} />
                      </TableCell>
                      <TableCell className="text-right">
                        <VariationCell value={p.variacaoPeriodo} />
                      </TableCell>
                      <TableCell className="text-right text-muted-foreground">
                        {p.mesesComVenda}/{result.periodos.length}
                      </TableCell>
                      <TableCell className="text-right">
                        <VariationBadge value={p.variacaoValor} />
                      </TableCell>
                      <TableCell className="text-right">{p.qtdAnterior}</TableCell>
                      <TableCell className="text-right">{p.qtdAtual}</TableCell>
                      <TableCell className="text-right">
                        <VariationBadge value={p.variacaoQtd} />
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </CardContent>
          </Card>
        </TabsContent>

        {/* Carteira · janela completa */}
        <TabsContent value="carteira" className="space-y-4">
          <CarteiraAnalysis result={result} />
        </TabsContent>

        {/* Alertas */}
        {hasAlerts && (
          <TabsContent value="alertas">
            <Card>
              <CardHeader>
                <CardTitle className="flex items-center gap-2 text-base text-destructive">
                  <AlertTriangle className="h-4 w-4" /> Inconsistências Encontradas
                </CardTitle>
              </CardHeader>
              <CardContent className="space-y-2">
                {result.inconsistencias.map((msg, i) => (
                  <p key={i} className="text-sm text-muted-foreground">
                    {msg}
                  </p>
                ))}
              </CardContent>
            </Card>
          </TabsContent>
        )}

        {/* Validation Report Tab */}
        <TabsContent value="validacao" className="space-y-4">
          {/* What was read from the file, tab by tab */}
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2 text-base">
                <FileSpreadsheet className="h-4 w-4 text-primary" /> Leitura do arquivo ·{" "}
                {result.leitura.abasLidas.length} de {result.leitura.abasNoArquivo.length} abas
              </CardTitle>
            </CardHeader>
            <CardContent className="p-0">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>Aba</TableHead>
                    <TableHead>Período</TableHead>
                    <TableHead>Identificado por</TableHead>
                    <TableHead className="text-right">Linhas</TableHead>
                    <TableHead className="text-right">Clientes</TableHead>
                    <TableHead className="text-right">Funcionários</TableHead>
                    <TableHead className="text-right">Produtos</TableHead>
                    <TableHead className="text-right">Faturamento</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {result.periodos.map((p) => (
                    <TableRow key={p.sheetName}>
                      <TableCell className="font-medium">
                        {p.sheetName}
                        {p.oculta && (
                          <Badge variant="outline" className="ml-2 text-muted-foreground">
                            oculta
                          </Badge>
                        )}
                      </TableCell>
                      <TableCell>{p.nome}</TableCell>
                      <TableCell className="text-muted-foreground">
                        {origemPeriodoLabel(p.origemPeriodo)}
                      </TableCell>
                      <TableCell className="text-right">{p.linhasLidas}</TableCell>
                      <TableCell
                        className={`text-right ${p.clientes.size === 0 ? "font-semibold text-destructive" : ""}`}
                      >
                        {p.clientesCarteira.size}
                      </TableCell>
                      <TableCell className="text-right text-muted-foreground">{p.funcionarios.size}</TableCell>
                      <TableCell className="text-right">{p.produtosLidos}</TableCell>
                      <TableCell className="text-right">{fmt(p.totalFaturamento)}</TableCell>
                    </TableRow>
                  ))}
                  {result.leitura.abasIgnoradas.map((a) => (
                    <TableRow key={a.sheetName} className="text-muted-foreground">
                      <TableCell className="font-medium">{a.sheetName}</TableCell>
                      <TableCell colSpan={7}>Não lida — {a.motivo}</TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </CardContent>
          </Card>

          {/* Spreadsheet × processed */}
          <ConferenciaCard result={result} />

          {/* Employees: in the revenue, out of the client analysis */}
          <FuncionariosCard result={result} />

          {/* Layer summary */}
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2 text-base">
                <ShieldCheck className="h-4 w-4 text-primary" /> Resumo de Validação por Camada
              </CardTitle>
            </CardHeader>
            <CardContent className="p-0">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>Camada</TableHead>
                    <TableHead className="text-right">Aprovadas</TableHead>
                    <TableHead className="text-right">Falhas</TableHead>
                    <TableHead className="text-right">Total</TableHead>
                    <TableHead className="text-right">Taxa</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {Object.entries(vr.layerSummary).map(([layer, counts]) => {
                    const total = counts.passed + counts.failed;
                    const rate = total > 0 ? (counts.passed / total) * 100 : 100;
                    return (
                      <TableRow key={layer}>
                        <TableCell className="font-medium capitalize">{layerLabel(layer)}</TableCell>
                        <TableCell className="text-right text-accent">{counts.passed}</TableCell>
                        <TableCell className="text-right text-destructive">{counts.failed}</TableCell>
                        <TableCell className="text-right">{total}</TableCell>
                        <TableCell className="text-right">
                          <Badge
                            variant="outline"
                            className={
                              rate === 100
                                ? "border-accent/30 bg-accent/10 text-accent"
                                : rate >= 90
                                ? "border-yellow-500/30 bg-yellow-500/10 text-yellow-600"
                                : "border-destructive/30 bg-destructive/10 text-destructive"
                            }
                          >
                            {rate.toFixed(1)}%
                          </Badge>
                        </TableCell>
                      </TableRow>
                    );
                  })}
                  <TableRow className="font-semibold">
                    <TableCell>Total</TableCell>
                    <TableCell className="text-right text-accent">{vr.passedChecks}</TableCell>
                    <TableCell className="text-right text-destructive">{vr.failedChecks}</TableCell>
                    <TableCell className="text-right">{vr.totalChecks}</TableCell>
                    <TableCell className="text-right">
                      <Badge
                        variant="outline"
                        className={
                          vr.valid
                            ? "border-accent/30 bg-accent/10 text-accent"
                            : "border-destructive/30 bg-destructive/10 text-destructive"
                        }
                      >
                        {vr.totalChecks > 0 ? ((vr.passedChecks / vr.totalChecks) * 100).toFixed(1) : "100.0"}%
                      </Badge>
                    </TableCell>
                  </TableRow>
                </TableBody>
              </Table>
            </CardContent>
          </Card>

          {/* Detailed diagnostics */}
          {vr.diagnostics.length > 0 && (
            <Card>
              <CardHeader>
                <CardTitle className="flex items-center gap-2 text-base">
                  <AlertTriangle className="h-4 w-4 text-destructive" /> Diagnósticos Detalhados
                </CardTitle>
              </CardHeader>
              <CardContent className="p-0">
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead className="w-[60px]">Sev.</TableHead>
                      <TableHead>Camada</TableHead>
                      <TableHead>Célula</TableHead>
                      <TableHead>Campo</TableHead>
                      <TableHead>Esperado</TableHead>
                      <TableHead>Encontrado</TableHead>
                      <TableHead>Impacto</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {vr.diagnostics.map((d, i) => (
                      <TableRow key={i}>
                        <TableCell>
                          <SeverityIcon severity={d.severity} />
                        </TableCell>
                        <TableCell className="capitalize text-xs">{layerLabel(d.layer)}</TableCell>
                        <TableCell className="font-mono text-xs">{d.cellRef}</TableCell>
                        <TableCell className="text-xs">{d.field}</TableCell>
                        <TableCell className="text-xs text-muted-foreground max-w-[200px] truncate" title={d.expected}>{d.expected}</TableCell>
                        <TableCell className="text-xs text-muted-foreground max-w-[200px] truncate" title={d.actual}>{d.actual}</TableCell>
                        <TableCell className="text-xs max-w-[300px]" title={d.impact}>{d.impact}</TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </CardContent>
            </Card>
          )}

          {vr.diagnostics.length === 0 && (
            <Card>
              <CardContent className="flex items-center gap-3 py-8 justify-center text-accent">
                <CheckCircle2 className="h-6 w-6" />
                <span className="text-lg font-medium">
                  Nenhum diagnóstico — todos os dados são consistentes com a planilha de origem.
                </span>
              </CardContent>
            </Card>
          )}
        </TabsContent>
      </Tabs>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Helper components
// ---------------------------------------------------------------------------

/**
 * Compact currency for axis ticks and bar labels: no cents, thousands folded
 * ("R$ 550", "R$ 1,3 mil"), so ticks stay on one line and read consistently.
 */
function fmtCurto(v: number): string {
  const abs = Math.abs(v);
  if (abs >= 1_000_000)
    return `R$ ${(v / 1_000_000).toLocaleString("pt-BR", { maximumFractionDigits: 1 })} mi`;
  if (abs >= 1_000)
    return `R$ ${(v / 1_000).toLocaleString("pt-BR", { maximumFractionDigits: 1 })} mil`;
  return `R$ ${v.toLocaleString("pt-BR", { maximumFractionDigits: 0 })}`;
}

/** Signed percentage, or an em dash when there is no base to compare against. */
function pct(v: number | null): string {
  if (v === null) return "—";
  return `${v > 0 ? "+" : ""}${v.toFixed(2)}%`;
}

const TENDENCIA_CONFIG: Record<Tendencia, { label: string; className: string; Icon: typeof TrendingUp }> = {
  crescimento: {
    label: "Crescimento",
    className: "border-accent/30 bg-accent/10 text-accent",
    Icon: TrendingUp,
  },
  queda: {
    label: "Queda",
    className: "border-destructive/30 bg-destructive/10 text-destructive",
    Icon: TrendingDown,
  },
  estavel: {
    label: "Estável",
    className: "border-border bg-muted text-muted-foreground",
    Icon: Minus,
  },
};

/** Sales movement of a month against the one before: alta, queda or estável. */
const MOVIMENTACAO_LABEL: Record<Tendencia, string> = {
  crescimento: "Alta",
  queda: "Queda",
  estavel: "Estável",
};

function MovimentacaoBadge({ tendencia }: { tendencia: Tendencia }) {
  const { className, Icon } = TENDENCIA_CONFIG[tendencia];
  return (
    <Badge variant="outline" className={`gap-1 ${className}`}>
      <Icon className="h-3 w-3" />
      {MOVIMENTACAO_LABEL[tendencia]}
    </Badge>
  );
}

function TendenciaBadge({ tendencia }: { tendencia: Tendencia }) {
  const { label, className, Icon } = TENDENCIA_CONFIG[tendencia];
  return (
    <Badge variant="outline" className={`gap-1 ${className}`}>
      <Icon className="h-3 w-3" />
      {label}
    </Badge>
  );
}

/** Legend labels wear text ink; the colour swatch beside them carries identity. */
const legendaEmTinta = (value: string) => (
  <span className="text-xs text-muted-foreground">{value}</span>
);

/** A movement count: coloured when something moved, muted at zero. */
function ContagemCell({
  valor,
  sinal,
  tom,
}: {
  valor: number | null;
  sinal: "+" | "−";
  tom: string;
}) {
  if (valor === null) {
    return (
      <TableCell className="text-right">
        <span className="text-xs text-muted-foreground">—</span>
      </TableCell>
    );
  }
  if (valor === 0) {
    return <TableCell className="text-right text-muted-foreground">0</TableCell>;
  }
  return (
    <TableCell className={`text-right font-medium ${tom}`}>
      {sinal}
      {valor}
    </TableCell>
  );
}

/** A variation cell that shows an em dash on the first month of the window. */
function VariationCell({ value }: { value: number | null }) {
  if (value === null) return <span className="text-xs text-muted-foreground">—</span>;
  return <VariationBadge value={value} />;
}

/** One figure of the window, with where it started and where it ended. */
function EvolucaoItem({
  label,
  de,
  para,
  variacao,
}: {
  label: string;
  de: string;
  para: string;
  variacao: number | null;
}) {
  return (
    <div className="rounded-lg border p-3">
      <p className="text-xs font-medium uppercase tracking-wide text-muted-foreground">{label}</p>
      <div className="mt-1 flex flex-wrap items-center gap-1.5">
        <span className="text-sm text-muted-foreground">{de}</span>
        <ArrowRight className="h-3 w-3 shrink-0 text-muted-foreground" />
        <span className="text-lg font-bold text-foreground">{para}</span>
      </div>
      {variacao !== null && (
        <div className="mt-1">
          <VariationBadge value={variacao} />
        </div>
      )}
    </div>
  );
}

/**
 * What improved and what got worse across the window, read off the consolidated
 * figures. Everything here is derived — no number is computed a second time.
 */
function destaques(result: AnalysisResult): { positivos: string[]; negativos: string[] } {
  const e = result.evolucaoConsolidada;
  const positivos: string[] = [];
  const negativos: string[] = [];

  const empurrar = (bom: boolean, texto: string) =>
    (bom ? positivos : negativos).push(texto);

  if (e.variacaoTotal !== null && Math.abs(e.variacaoTotal) >= 1) {
    empurrar(
      e.variacaoTotal > 0,
      `Faturamento ${e.variacaoTotal > 0 ? "subiu" : "caiu"} ${pct(e.variacaoTotal)} de ${e.primeiro} (${fmt(e.faturamentoInicial)}) para ${e.ultimo} (${fmt(e.faturamentoFinal)}).`,
    );
  }

  if (e.clientesFinal !== e.clientesInicial) {
    const d = e.clientesFinal - e.clientesInicial;
    empurrar(
      d > 0,
      `Base de clientes ${d > 0 ? "cresceu" : "encolheu"} de ${e.clientesInicial} para ${e.clientesFinal} clientes no mês (${d > 0 ? "+" : ""}${d}).`,
    );
  }

  if (e.variacaoTicketMedio !== null && Math.abs(e.variacaoTicketMedio) >= 1) {
    empurrar(
      e.variacaoTicketMedio > 0,
      `Ticket médio ${e.variacaoTicketMedio > 0 ? "subiu" : "caiu"} ${pct(e.variacaoTicketMedio)} (${fmt(e.ticketMedioInicial)} → ${fmt(e.ticketMedioFinal)}).`,
    );
  }

  if (e.novosNoPeriodo > 0) {
    positivos.push(`${e.novosNoPeriodo} cliente${e.novosNoPeriodo > 1 ? "s" : ""} comprou pela primeira vez ao longo do período.`);
  }
  if (e.perdidosNoPeriodo > 0) {
    negativos.push(`${e.perdidosNoPeriodo} cliente${e.perdidosNoPeriodo > 1 ? "s" : ""} está há 2 meses ou mais sem comprar.`);
  }
  if (result.carteira.frequenciaReduzida.length > 0) {
    negativos.push(`${result.carteira.frequenciaReduzida.length} cliente${result.carteira.frequenciaReduzida.length > 1 ? "s" : ""} deixou de comprar no último mês — recuperável.`);
  }
  if (e.clientesRecorrentes > 0) {
    positivos.push(`${e.clientesRecorrentes} cliente${e.clientesRecorrentes > 1 ? "s" : ""} comprou em todos os ${e.meses} meses.`);
  }
  if (e.taxaRetencaoMedia !== null) {
    empurrar(
      e.taxaRetencaoMedia >= 60,
      `Retenção média de ${e.taxaRetencaoMedia.toFixed(0)}% dos clientes de um mês para o seguinte.`,
    );
  }
  if (e.mesesEmQueda > 0 && e.mesesEmAlta > 0) {
    (e.mesesEmAlta >= e.mesesEmQueda ? positivos : negativos).push(
      `${e.mesesEmAlta} ${e.mesesEmAlta === 1 ? "mês subiu" : "meses subiram"} e ${e.mesesEmQueda} ${e.mesesEmQueda === 1 ? "caiu" : "caíram"} dentro da janela.`,
    );
  }

  // Products moving the most, in either direction.
  const comHistorico = result.comparativoProdutos.filter((p) => p.variacaoPeriodo !== null);
  const subindo = [...comHistorico].sort((a, b) => (b.variacaoPeriodo ?? 0) - (a.variacaoPeriodo ?? 0))[0];
  const caindo = [...comHistorico].sort((a, b) => (a.variacaoPeriodo ?? 0) - (b.variacaoPeriodo ?? 0))[0];
  if (subindo && (subindo.variacaoPeriodo ?? 0) > 5) {
    positivos.push(`Produto em maior alta: ${subindo.produto} (${pct(subindo.variacaoPeriodo)} de ${e.primeiro} a ${e.ultimo}).`);
  }
  if (caindo && (caindo.variacaoPeriodo ?? 0) < -5) {
    negativos.push(`Produto em maior queda: ${caindo.produto} (${pct(caindo.variacaoPeriodo)} de ${e.primeiro} a ${e.ultimo}).`);
  }

  return { positivos, negativos };
}

/** The headline card of the summary: the whole window read as one movement. */
function EvolucaoJanela({ result }: { result: AnalysisResult }) {
  const e = result.evolucaoConsolidada;
  const { positivos, negativos } = destaques(result);

  return (
    <Card>
      <CardHeader className="pb-3">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <CardTitle className="flex items-center gap-2 text-base">
            <TrendingUp className="h-4 w-4 text-primary" />
            Evolução da janela · {e.primeiro} → {e.ultimo}
          </CardTitle>
          <TendenciaBadge tendencia={e.tendencia} />
        </div>
        <p className="text-xs text-muted-foreground">
          {e.meses} {e.meses === 1 ? "mês analisado" : "meses analisados"} em sequência.
          Os indicadores comparam o primeiro mês com o último; o total da janela aparece
          apenas como contexto.
        </p>
      </CardHeader>
      <CardContent className="space-y-5">
        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
          <EvolucaoItem
            label="Faturamento"
            de={fmt(e.faturamentoInicial)}
            para={fmt(e.faturamentoFinal)}
            variacao={e.variacaoTotal}
          />
          <EvolucaoItem
            label="Clientes no mês"
            de={String(e.clientesInicial)}
            para={String(e.clientesFinal)}
            variacao={e.variacaoClientes}
          />
          <EvolucaoItem
            label="Ticket médio"
            de={fmt(e.ticketMedioInicial)}
            para={fmt(e.ticketMedioFinal)}
            variacao={e.variacaoTicketMedio}
          />
          <div className="rounded-lg border p-3">
            <p className="text-xs font-medium uppercase tracking-wide text-muted-foreground">
              Ritmo da janela
            </p>
            <p className="mt-1 text-lg font-bold text-foreground">
              {pct(e.variacaoMediaMensal)}
              <span className="ml-1 text-xs font-normal text-muted-foreground">por mês</span>
            </p>
            <p className="mt-1 text-xs text-muted-foreground">
              {e.mesesEmAlta} {e.mesesEmAlta === 1 ? "mês em alta" : "meses em alta"} ·{" "}
              {e.mesesEmQueda} em queda
            </p>
          </div>
        </div>

        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4 text-sm">
          <div className="rounded-lg border border-dashed p-3">
            <p className="text-xs text-muted-foreground">Melhor mês</p>
            <p className="font-medium capitalize">{e.melhorMes.nome}</p>
            <p className="text-xs text-muted-foreground">{fmt(e.melhorMes.faturamento)}</p>
          </div>
          <div className="rounded-lg border border-dashed p-3">
            <p className="text-xs text-muted-foreground">Pior mês</p>
            <p className="font-medium capitalize">{e.piorMes.nome}</p>
            <p className="text-xs text-muted-foreground">{fmt(e.piorMes.faturamento)}</p>
          </div>
          <div className="rounded-lg border border-dashed p-3">
            <p className="text-xs text-muted-foreground">Clientes distintos</p>
            <p className="font-medium">{e.clientesDistintos}</p>
            <p className="text-xs text-muted-foreground">
              {e.clientesRecorrentes} em todos os meses
            </p>
          </div>
          <div className="rounded-lg border border-dashed p-3">
            <p className="text-xs text-muted-foreground">Total da janela</p>
            <p className="font-medium">{fmt(e.faturamentoTotal)}</p>
            <p className="text-xs text-muted-foreground">
              média de {fmt(e.faturamentoMedio)}/mês
            </p>
          </div>
        </div>

        <div className="grid gap-4 md:grid-cols-2">
          <div className="rounded-lg border border-accent/30 bg-accent/5 p-3">
            <p className="mb-2 flex items-center gap-1.5 text-sm font-semibold text-accent">
              <TrendingUp className="h-4 w-4" /> O que melhorou
            </p>
            {positivos.length > 0 ? (
              <ul className="space-y-1.5 text-sm text-muted-foreground">
                {positivos.map((t) => (
                  <li key={t}>• {t}</li>
                ))}
              </ul>
            ) : (
              <p className="text-sm text-muted-foreground">
                Nenhuma melhora relevante no período.
              </p>
            )}
          </div>
          <div className="rounded-lg border border-destructive/30 bg-destructive/5 p-3">
            <p className="mb-2 flex items-center gap-1.5 text-sm font-semibold text-destructive">
              <TrendingDown className="h-4 w-4" /> O que piorou
            </p>
            {negativos.length > 0 ? (
              <ul className="space-y-1.5 text-sm text-muted-foreground">
                {negativos.map((t) => (
                  <li key={t}>• {t}</li>
                ))}
              </ul>
            ) : (
              <p className="text-sm text-muted-foreground">
                Nenhuma piora relevante no período.
              </p>
            )}
          </div>
        </div>
      </CardContent>
    </Card>
  );
}

/** One month-to-month move, with the parts that explain the revenue change. */
function TransicaoCard({ t }: { t: TransicaoPeriodo }) {
  const partes = [
    { label: "Clientes novos", valor: t.ponte.novos, detalhe: `${t.novos.total} cliente(s)` },
    { label: "Clientes que retornaram", valor: t.ponte.retornados, detalhe: `${t.retornados.total} cliente(s)` },
    { label: "Mantidos que compraram mais", valor: t.ponte.expansao, detalhe: `${t.mantidos.total} mantido(s)` },
    { label: "Mantidos que compraram menos", valor: t.ponte.contracao, detalhe: `${t.mantidos.total} mantido(s)` },
    { label: "Deixaram de comprar", valor: t.ponte.perdidos, detalhe: `${t.perdidos.total} cliente(s)` },
    // Employees buy, but are not clients: their part keeps the bridge closed.
    ...(t.ponte.funcionarios !== 0
      ? [{ label: "Compras de funcionários", valor: t.ponte.funcionarios, detalhe: "fora da carteira" }]
      : []),
  ];

  return (
    <div className="rounded-lg border p-4">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <p className="flex items-center gap-2 text-sm font-semibold capitalize">
          {t.de}
          <ArrowRight className="h-3.5 w-3.5 text-muted-foreground" />
          {t.para}
        </p>
        <div className="flex flex-wrap items-center gap-2 text-sm">
          <span className="text-muted-foreground">{fmt(t.faturamentoDe)}</span>
          <ArrowRight className="h-3 w-3 text-muted-foreground" />
          <span className="font-semibold">{fmt(t.faturamentoPara)}</span>
          <VariationCell value={t.variacaoFaturamento} />
          <span
            className={
              t.diferencaFaturamento >= 0
                ? "text-xs font-medium text-accent"
                : "text-xs font-medium text-destructive"
            }
          >
            {t.diferencaFaturamento >= 0 ? "+" : "−"}
            {fmt(Math.abs(t.diferencaFaturamento))}
          </span>
        </div>
      </div>

      <div className="mt-3 grid gap-3 md:grid-cols-2">
        <div>
          <p className="mb-1.5 text-xs font-medium uppercase tracking-wide text-muted-foreground">
            Movimentação de clientes
          </p>
          <div className="grid grid-cols-2 gap-2 text-sm">
            <MovItem label="Novos" valor={t.novos.total} receita={t.novos.faturamento} tom="accent" />
            <MovItem label="Retornaram" valor={t.retornados.total} receita={t.retornados.faturamento} tom="primary" />
            <MovItem label="Mantidos" valor={t.mantidos.total} receita={t.mantidos.faturamento} tom="neutro" />
            <MovItem label="Perderam-se" valor={t.perdidos.total} receita={t.perdidos.faturamento} tom="destructive" />
          </div>
          <p className="mt-2 text-xs text-muted-foreground">
            Retenção:{" "}
            <span className="font-medium text-foreground">
              {t.taxaRetencao === null ? "—" : `${t.taxaRetencao.toFixed(0)}%`}
            </span>{" "}
            dos {t.clientesDe} clientes de {t.de} · ticket médio {fmt(t.ticketMedioDe)} →{" "}
            {fmt(t.ticketMedioPara)}
          </p>
        </div>

        <div>
          <p className="mb-1.5 text-xs font-medium uppercase tracking-wide text-muted-foreground">
            O que explica a variação
          </p>
          <div className="space-y-1 text-sm">
            {partes.map((parte) => (
              <div key={parte.label} className="flex items-center justify-between gap-2">
                <span className="text-muted-foreground">
                  {parte.label}
                  <span className="ml-1 text-xs opacity-70">({parte.detalhe})</span>
                </span>
                <span
                  className={
                    parte.valor > 0
                      ? "font-medium text-accent"
                      : parte.valor < 0
                        ? "font-medium text-destructive"
                        : "text-muted-foreground"
                  }
                >
                  {parte.valor > 0 ? "+" : parte.valor < 0 ? "−" : ""}
                  {fmt(Math.abs(parte.valor))}
                </span>
              </div>
            ))}
            <div className="flex items-center justify-between gap-2 border-t pt-1 font-semibold">
              <span>Variação do mês</span>
              <span className={t.ponte.total >= 0 ? "text-accent" : "text-destructive"}>
                {t.ponte.total >= 0 ? "+" : "−"}
                {fmt(Math.abs(t.ponte.total))}
              </span>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

/**
 * Entrada or saída of clients at every passage of month.
 *
 * Each month gets a summary line and then one row per client — name, last
 * purchase date, revenue and classification in their own columns — instead of
 * every client packed into a single cell of running text.
 */
function EntradaSaidaPorMes({
  result,
  tipo,
}: {
  result: AnalysisResult;
  tipo: "entrada" | "saida";
}) {
  if (result.transicoes.length === 0) {
    return (
      <Card>
        <CardContent className="py-8 text-center text-muted-foreground">
          É preciso pelo menos dois meses para acompanhar entrada e saída de clientes.
        </CardContent>
      </Card>
    );
  }

  return (
    <div className="space-y-4">
      {result.transicoes.map((t) => (
        <PassagemDeMes key={`${t.de}-${t.para}`} t={t} tipo={tipo} />
      ))}
    </div>
  );
}

/** One passage of month: the totals, then the clients behind them. */
function PassagemDeMes({ t, tipo }: { t: TransicaoPeriodo; tipo: "entrada" | "saida" }) {
  // Novos and retornados are listed together, each keeping its classification —
  // the operational question is "who bought this month that did not last month".
  const linhas: { cliente: ClienteMovimento; classe: "Novo" | "Retorno" | "Saiu" }[] =
    tipo === "entrada"
      ? [
          ...t.novos.clientes.map((c) => ({ cliente: c, classe: "Novo" as const })),
          ...t.retornados.clientes.map((c) => ({ cliente: c, classe: "Retorno" as const })),
        ].sort((a, b) => a.cliente.nome.localeCompare(b.cliente.nome, "pt-BR"))
      : t.perdidos.clientes.map((c) => ({ cliente: c, classe: "Saiu" as const }));

  const temContato = linhas.some(({ cliente }) => cliente.fone || cliente.celular);

  return (
    <Card>
      <CardHeader className="pb-3">
        <CardTitle className="flex items-center gap-2 text-base capitalize">
          <CalendarRange className="h-4 w-4 text-primary" />
          {t.de}
          <ArrowRight className="h-3.5 w-3.5 text-muted-foreground" />
          {t.para}
        </CardTitle>
        <div className="flex flex-wrap gap-x-6 gap-y-1 text-sm">
          {tipo === "entrada" ? (
            <>
              <Resumo rotulo="Novos" valor={t.novos.total} tom="text-accent" />
              <Resumo rotulo="Receita dos novos" valor={fmt(t.novos.faturamento)} />
              <Resumo rotulo="Retornaram" valor={t.retornados.total} tom="text-primary" />
              <Resumo rotulo="Receita dos retornos" valor={fmt(t.retornados.faturamento)} />
            </>
          ) : (
            <>
              <Resumo rotulo="Deixaram de comprar" valor={t.perdidos.total} tom="text-destructive" />
              <Resumo rotulo="Receita que saiu" valor={fmt(t.perdidos.faturamento)} />
              <Resumo rotulo="Mantidos" valor={t.mantidos.total} />
              <Resumo
                rotulo="Retenção"
                valor={t.taxaRetencao === null ? "—" : `${t.taxaRetencao.toFixed(0)}%`}
              />
            </>
          )}
        </div>
      </CardHeader>
      <CardContent className="p-0">
        {linhas.length === 0 ? (
          <p className="px-6 pb-5 text-sm text-muted-foreground">
            {tipo === "entrada"
              ? `Nenhum cliente entrou em ${t.para}.`
              : `Nenhum cliente deixou de comprar em ${t.para}.`}
          </p>
        ) : (
          <div className="overflow-x-auto">
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead>Cliente</TableHead>
                  <TableHead>Data da última compra</TableHead>
                  <TableHead>Código</TableHead>
                  <TableHead className="text-right">
                    {tipo === "entrada" ? `Receita em ${t.para}` : `Receita em ${t.de}`}
                  </TableHead>
                  {tipo === "entrada" && <TableHead>Classificação</TableHead>}
                  {temContato && <TableHead>Fone</TableHead>}
                  {temContato && <TableHead>Celular</TableHead>}
                </TableRow>
              </TableHeader>
              <TableBody>
                {linhas.map(({ cliente, classe }) => (
                  <TableRow key={cliente.chave}>
                    <TableCell className="font-medium">{cliente.nome}</TableCell>
                    <TableCell className="whitespace-nowrap">
                      {cliente.dataUltimaCompra ?? (
                        <span className="text-xs text-muted-foreground">—</span>
                      )}
                    </TableCell>
                    <TableCell className="text-muted-foreground">
                      {cliente.codigoCliente ?? "—"}
                    </TableCell>
                    <TableCell className="text-right">{fmt(cliente.faturamento)}</TableCell>
                    {tipo === "entrada" && (
                      <TableCell>
                        <Badge
                          variant="outline"
                          className={
                            classe === "Novo"
                              ? "border-accent/30 bg-accent/10 text-accent"
                              : "border-primary/30 bg-primary/10 text-primary"
                          }
                        >
                          {classe}
                        </Badge>
                      </TableCell>
                    )}
                    {temContato && (
                      <TableCell className="whitespace-nowrap text-muted-foreground">
                        {cliente.fone ?? "—"}
                      </TableCell>
                    )}
                    {temContato && (
                      <TableCell className="whitespace-nowrap text-muted-foreground">
                        {cliente.celular ?? "—"}
                      </TableCell>
                    )}
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </div>
        )}
      </CardContent>
    </Card>
  );
}

/** One labelled field of a client, kept apart from the client's name. */
function CampoCliente({ rotulo, valor }: { rotulo: string; valor: string }) {
  return (
    <span className="flex items-baseline gap-1">
      <span className="uppercase tracking-wide text-muted-foreground">{rotulo}</span>
      <span className="font-medium text-foreground">{valor}</span>
    </span>
  );
}

/** A labelled figure in the header of a passage card. */
function Resumo({
  rotulo,
  valor,
  tom,
}: {
  rotulo: string;
  valor: string | number;
  tom?: string;
}) {
  return (
    <span className="flex items-baseline gap-1.5">
      <span className="text-xs uppercase tracking-wide text-muted-foreground">{rotulo}</span>
      <span className={`font-semibold ${tom ?? "text-foreground"}`}>{valor}</span>
    </span>
  );
}

function MovItem({
  label,
  valor,
  receita,
  tom,
}: {
  label: string;
  valor: number;
  receita: number;
  tom: "accent" | "primary" | "destructive" | "neutro";
}) {
  const cor =
    tom === "accent"
      ? "text-accent"
      : tom === "primary"
        ? "text-primary"
        : tom === "destructive"
          ? "text-destructive"
          : "text-foreground";
  return (
    <div className="rounded-md border px-2 py-1.5">
      <p className="text-xs text-muted-foreground">{label}</p>
      <p className={`text-lg font-bold leading-tight ${cor}`}>{valor}</p>
      <p className="text-xs text-muted-foreground">{fmt(receita)}</p>
    </div>
  );
}

/** Older months fade toward grey; the current month gets the strong accent. */
function periodColor(index: number, total: number): string {
  if (index === total - 1) return "hsl(217 91% 50%)";
  const lightness = 82 - (index / Math.max(total - 1, 1)) * 22;
  return `hsl(217 25% ${lightness}%)`;
}

/**
 * Conferência: what each tab of the spreadsheet holds — sale rows, buyers,
 * products, units sold and revenue — read straight from the sale rows. The
 * spreadsheet is the only source: nothing is compared with anything else.
 */
function ConferenciaCard({ result }: { result: AnalysisResult }) {
  const c = result.conferencia;
  const comDescarte = c.abas.filter((a) => a.linhasDescartadas > 0);

  return (
    <Card>
      <CardHeader>
        <CardTitle className="flex items-center gap-2 text-base">
          <FileSpreadsheet className="h-4 w-4 text-primary" />
          Conferência dos dados da planilha
        </CardTitle>
        <p className="text-xs text-muted-foreground">
          Números de cada aba, apurados somente das linhas de venda da planilha. A
          quantidade vendida soma todas as colunas de quantidade dessas linhas, sem os
          totais por cliente, por unidade ou da aba.
        </p>
      </CardHeader>
      <CardContent className="space-y-4 p-0">
        <div className="overflow-x-auto">
          <Table>
            <TableHeader>
              <TableRow>
                <TableHead>Aba</TableHead>
                <TableHead className="text-right">Registros</TableHead>
                <TableHead className="text-right">Clientes</TableHead>
                <TableHead className="text-right">Funcionários</TableHead>
                <TableHead className="text-right">Produtos</TableHead>
                <TableHead className="text-right">Qtd. vendida</TableHead>
                <TableHead className="text-right">Faturamento</TableHead>
              </TableRow>
            </TableHeader>
            <TableBody>
              {c.abas.map((a) => (
                <TableRow key={a.aba}>
                  <TableCell className="font-medium">
                    {a.aba}
                    {a.periodo !== a.aba && (
                      <span className="ml-1 text-xs text-muted-foreground">({a.periodo})</span>
                    )}
                  </TableCell>
                  <TableCell className="text-right">{a.registrosVenda}</TableCell>
                  <TableCell className="text-right">{a.clientesCarteira}</TableCell>
                  <TableCell className="text-right text-muted-foreground">{a.funcionarios}</TableCell>
                  <TableCell className="text-right">{a.produtosDistintos}</TableCell>
                  <TableCell className="text-right">{fmtQtd(a.quantidadeTotal)}</TableCell>
                  <TableCell className="text-right font-medium">{fmt(a.faturamento)}</TableCell>
                </TableRow>
              ))}
              <TableRow className="font-semibold">
                <TableCell>Janela · {c.consolidado.abas} abas</TableCell>
                <TableCell className="text-right">{c.consolidado.registrosVenda}</TableCell>
                <TableCell className="text-right">{c.consolidado.clientesDistintos}</TableCell>
                <TableCell className="text-right">{c.consolidado.funcionariosDistintos}</TableCell>
                <TableCell className="text-right">{c.consolidado.produtosDistintos}</TableCell>
                <TableCell className="text-right">{fmtQtd(c.consolidado.quantidadeTotal)}</TableCell>
                <TableCell className="text-right">{fmt(c.consolidado.faturamento)}</TableCell>
              </TableRow>
            </TableBody>
          </Table>
        </div>
        <p className="px-6 pb-5 text-xs text-muted-foreground">
          Na linha da janela, clientes e produtos são os distintos em todos os meses.
          {comDescarte.length > 0 &&
            ` Linhas com valor que não são venda nem total: ${comDescarte
              .map((a) => `${a.aba} (${a.linhasDescartadas}, ${fmt(a.valorDescartado)})`)
              .join("; ")} — veja os alertas.`}
        </p>
      </CardContent>
    </Card>
  );
}

/** Employees found in the file: their purchases are revenue, not clients. */
function FuncionariosCard({ result }: { result: AnalysisResult }) {
  const periodos = result.periodos.map((p) => p.nome);
  return (
    <Card>
      <CardHeader>
        <CardTitle className="flex items-center gap-2 text-base">
          <Users className="h-4 w-4 text-primary" /> Funcionários ({result.funcionarios.length})
        </CardTitle>
        <p className="text-xs text-muted-foreground">
          Identificados pelo asterisco (*) no início do nome. As compras entram no
          faturamento, nos produtos e nos indicadores de vendas; ficam fora de clientes
          novos, perdidos, carteira, retenção e ticket médio por cliente.
        </p>
      </CardHeader>
      <CardContent className="p-0">
        {result.funcionarios.length === 0 ? (
          <p className="px-6 pb-5 text-sm text-muted-foreground">
            Nenhum funcionário (*) encontrado no arquivo.
          </p>
        ) : (
          <div className="overflow-x-auto">
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead>Funcionário</TableHead>
                  <TableHead>Código</TableHead>
                  {periodos.map((n) => (
                    <TableHead key={n} className="text-right capitalize">
                      {n}
                    </TableHead>
                  ))}
                  <TableHead className="text-right">Total</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {result.funcionarios.map((f) => (
                  <TableRow key={f.chave}>
                    <TableCell className="font-medium">{f.nome}</TableCell>
                    <TableCell className="text-muted-foreground">{f.codigoCliente ?? "—"}</TableCell>
                    {periodos.map((n) => {
                      const v = f.faturamentoPorPeriodo[n] ?? 0;
                      return (
                        <TableCell key={n} className="text-right">
                          {v !== 0 ? fmt(v) : <span className="text-xs text-muted-foreground">—</span>}
                        </TableCell>
                      );
                    })}
                    <TableCell className="text-right font-medium">{fmt(f.faturamentoTotal)}</TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </div>
        )}
      </CardContent>
    </Card>
  );
}

function origemPeriodoLabel(origem: AnalysisResult["periodos"][number]["origemPeriodo"]): string {
  switch (origem) {
    case "nome-da-aba":
      return "nome da aba";
    case "conteudo":
      return "conteúdo da aba";
    default:
      return "ordem no arquivo";
  }
}

function layerLabel(layer: string): string {
  const labels: Record<string, string> = {
    existence: "Existência",
    type: "Tipo",
    integrity: "Integridade",
    structure: "Estrutura",
    traceability: "Rastreabilidade",
    math: "Matemática",
    duplicate: "Duplicidade",
    "cross-field": "Campos Cruzados",
  };
  return labels[layer] || layer;
}

function SeverityIcon({ severity }: { severity: ValidationSeverity }) {
  switch (severity) {
    case "error":
      return <XCircle className="h-4 w-4 text-destructive" />;
    case "warning":
      return <AlertTriangle className="h-4 w-4 text-yellow-500" />;
    case "info":
      return <Info className="h-4 w-4 text-blue-500" />;
  }
}

function Row({ label, value, highlight }: { label: string; value: string; highlight?: boolean }) {
  return (
    <div className="flex items-center justify-between">
      <span className="text-muted-foreground">{label}</span>
      <span className={highlight ? "font-bold text-primary" : "font-medium text-foreground"}>
        {value}
      </span>
    </div>
  );
}

function VariationBadge({ value }: { value: number }) {
  if (value === 0) return <Badge variant="secondary">0%</Badge>;
  return (
    <Badge
      variant="outline"
      className={
        value > 0
          ? "border-accent/30 bg-accent/10 text-accent"
          : "border-destructive/30 bg-destructive/10 text-destructive"
      }
    >
      {value > 0 ? "+" : ""}
      {value.toFixed(2)}%
    </Badge>
  );
}

const STATUS_CONFIG: Record<
  StatusCarteira,
  { label: string; className: string }
> = {
  ativo: {
    label: "Ativo",
    className: "border-accent/30 bg-accent/10 text-accent",
  },
  frequencia_reduzida: {
    label: "Frequência Reduzida",
    className: "border-yellow-500/30 bg-yellow-500/10 text-yellow-600",
  },
  risco_abandono: {
    label: "Risco de Abandono",
    className: "border-orange-500/30 bg-orange-500/10 text-orange-600",
  },
  inativo: {
    label: "Inativo",
    className: "border-destructive/30 bg-destructive/10 text-destructive",
  },
};

function CarteiraAnalysis({ result }: { result: AnalysisResult }) {
  const { classificacaoCarteira, periodos } = result;

  const counts: Record<StatusCarteira, number> = {
    ativo: 0,
    frequencia_reduzida: 0,
    risco_abandono: 0,
    inativo: 0,
  };
  for (const c of classificacaoCarteira) counts[c.status]++;

  return (
    <div className="space-y-4">
      <Card>
        <CardHeader>
          <CardTitle className="text-base">
            Análise de Carteira — Histórico de {periodos.length}{" "}
            {periodos.length === 1 ? "Mês" : "Meses"}
          </CardTitle>
          <p className="text-xs text-muted-foreground">
            Períodos considerados:{" "}
            {periodos.map((p) => p.nome).join(" → ")}
          </p>
        </CardHeader>
        <CardContent>
          <div className="grid grid-cols-2 gap-3 sm:grid-cols-4">
            {(Object.entries(counts) as [StatusCarteira, number][]).map(
              ([status, count]) => (
                <div
                  key={status}
                  className="rounded-lg border p-3 text-center"
                >
                  <div className="text-2xl font-bold">{count}</div>
                  <Badge
                    variant="outline"
                    className={`mt-1 text-xs ${STATUS_CONFIG[status].className}`}
                  >
                    {STATUS_CONFIG[status].label}
                  </Badge>
                </div>
              ),
            )}
          </div>
        </CardContent>
      </Card>

      <CarteiraTable
        clientes={classificacaoCarteira}
        periodos={periodos.map((p) => p.nome)}
        emptyMessage="Nenhum cliente encontrado."
      />
    </div>
  );
}

function CarteiraTable({
  clientes,
  periodos,
  emptyMessage,
}: {
  clientes: ClienteCarteira[];
  /** Period names, oldest first — one column each */
  periodos: string[];
  emptyMessage: string;
}) {
  if (clientes.length === 0) {
    return (
      <Card>
        <CardContent className="py-12 text-center text-muted-foreground">
          {emptyMessage}
        </CardContent>
      </Card>
    );
  }

  return (
    <Card>
      <CardContent className="overflow-x-auto p-0">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>Cliente</TableHead>
              <TableHead>Data da última compra</TableHead>
              <TableHead>Código</TableHead>
              <TableHead>Status</TableHead>
              {periodos.map((n) => (
                <TableHead key={n} className="text-right capitalize">
                  {n}
                </TableHead>
              ))}
              <TableHead className="text-right">Meses</TableHead>
              <TableHead>Tendência</TableHead>
              <TableHead className="text-right">Var. período</TableHead>
              <TableHead>Último mês com compra</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {clientes.map((c) => (
              <TableRow key={c.chave}>
                <TableCell className="font-medium">{c.nome}</TableCell>
                <TableCell className="whitespace-nowrap">
                  {c.dataUltimaCompra ?? (
                    <span className="text-xs text-muted-foreground">—</span>
                  )}
                </TableCell>
                <TableCell className="text-muted-foreground">
                  {c.codigoCliente ?? "—"}
                </TableCell>
                <TableCell>
                  <Badge variant="outline" className={STATUS_CONFIG[c.status].className}>
                    {STATUS_CONFIG[c.status].label}
                  </Badge>
                </TableCell>
                {periodos.map((n) => {
                  const v = c.faturamentoPorPeriodo[n] ?? 0;
                  return (
                    <TableCell
                      key={n}
                      className={
                        v > 0
                          ? "text-right text-foreground"
                          : "text-right text-xs text-muted-foreground"
                      }
                    >
                      {v > 0 ? fmt(v) : "—"}
                    </TableCell>
                  );
                })}
                <TableCell className="text-right">
                  {c.totalPeriodosPresente}/{periodos.length}
                </TableCell>
                <TableCell>
                  <TendenciaBadge tendencia={c.tendencia} />
                </TableCell>
                <TableCell className="text-right">
                  <VariationCell value={c.variacaoPeriodo} />
                </TableCell>
                <TableCell className="text-sm capitalize text-muted-foreground">
                  {c.ultimoPeriodo}
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </CardContent>
    </Card>
  );
}

function ClientTable({
  clients,
  emptyMessage,
}: {
  clients: ClientData[];
  emptyMessage: string;
}) {
  if (clients.length === 0) {
    return (
      <Card>
        <CardContent className="py-12 text-center text-muted-foreground">
          {emptyMessage}
        </CardContent>
      </Card>
    );
  }

  return (
    <div className="space-y-4">
      {clients.map((c) => (
        <Card key={c.chave}>
          <CardHeader className="pb-2">
            <div className="flex items-center justify-between gap-2">
              <CardTitle className="text-base">{c.nome}</CardTitle>
              <Badge variant="secondary" className="text-sm font-semibold">
                {fmt(c.totalGeral)}
              </Badge>
            </div>
            {/* the report's own fields, each labelled — never appended to the name */}
            <div className="flex flex-wrap gap-x-5 gap-y-1 pt-1 text-xs">
              {c.codigoCliente && <CampoCliente rotulo="Código" valor={c.codigoCliente} />}
              {c.dataUltimaCompra && (
                <CampoCliente rotulo="Data da última compra" valor={c.dataUltimaCompra} />
              )}
              {c.fone && <CampoCliente rotulo="Fone" valor={c.fone} />}
              {c.celular && <CampoCliente rotulo="Celular" valor={c.celular} />}
            </div>
          </CardHeader>
          <CardContent className="p-0">
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead>Produto</TableHead>
                  <TableHead>Código</TableHead>
                  <TableHead className="text-right">Quantidade</TableHead>
                  <TableHead className="text-right">Valor Total</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {c.produtos.map((p, i) => (
                  <TableRow key={`${p.codigo}-${i}`}>
                    <TableCell>{p.produto}</TableCell>
                    <TableCell className="text-muted-foreground">{p.codigo}</TableCell>
                    <TableCell className="text-right">{p.quantidade}</TableCell>
                    <TableCell className="text-right">{fmt(p.valorTotal)}</TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </CardContent>
        </Card>
      ))}
    </div>
  );
}
