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
} from "recharts";
import type { AnalysisResult, StatusCarteira, ClienteCarteira } from "@/lib/excel-analyzer";
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

export function AnalysisResults({ result }: Props) {
  const [tab, setTab] = useState("resumo");

  const periodNames = result.periodos.map((p) => p.nome);

  // One series per month of the window instead of just anterior/atual.
  const chartData = result.comparativoProdutos.slice(0, 10).map((p) => ({
    nome: p.produto.length > 18 ? p.produto.slice(0, 16) + "…" : p.produto,
    ...Object.fromEntries(periodNames.map((n) => [n, p.valorPorPeriodo[n] ?? 0])),
  }));

  const evolucaoChartData = result.evolucaoPeriodos.map((p) => ({
    nome: p.nome,
    Faturamento: p.faturamento,
    Clientes: p.clientes,
  }));

  // Tabs that were read but yielded nothing — the shape a lost month takes.
  const periodosVazios = result.periodos.filter((p) => p.clientes.size === 0);

  const vr = result.validationReport;
  const hasAlerts = result.inconsistencias.length > 0 || vr.diagnostics.length > 0;
  const errorCount = vr.diagnostics.filter((d) => d.severity === "error").length;
  const warningCount = vr.diagnostics.filter((d) => d.severity === "warning").length;

  return (
    <div className="space-y-6">
      {/* KPI cards */}
      <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
        <KpiCard
          title="Faturamento Atual"
          value={fmt(result.periodoAtual.totalFaturamento)}
          icon={DollarSign}
          trend={result.variacaoFaturamento}
          color="primary"
        />
        <KpiCard
          title="Clientes Ativos"
          value={String(result.carteira.ativos.length)}
          subtitle={`${result.periodoAnterior.clientes.size} no período anterior`}
          icon={Users}
          color="primary"
        />
        <KpiCard
          title="Clientes Novos"
          value={String(result.carteira.novos.length)}
          subtitle={
            fmt(result.carteira.faturamentoNovos) +
            ` · sem compra nos ${janelaLabel(result)} anteriores`
          }
          icon={UserPlus}
          color="success"
        />
        <KpiCard
          title="Clientes Perdidos"
          value={String(result.carteira.perdidos.length)}
          subtitle={
            fmt(result.carteira.faturamentoPerdidos) + " · 2+ meses sem comprar"
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
                title={`Aba "${p.sheetName}" · ${p.linhasLidas} linhas · ${p.clientes.size} clientes · ${p.produtosLidos} produtos`}
                className={
                  i === result.periodos.length - 1
                    ? "border-primary/40 bg-primary/10 text-primary"
                    : "text-muted-foreground"
                }
              >
                {p.nome}
                {i === result.periodos.length - 1 && " (atual)"}
                <span className="ml-1.5 opacity-70">
                  {p.clientes.size} cli · {p.produtosLidos} prod
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
          {/* Month-by-month comparison across the whole window */}
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center gap-2 text-base">
                <CalendarRange className="h-4 w-4 text-primary" />
                Comparativo mês a mês · {result.periodos.length}{" "}
                {result.periodos.length === 1 ? "mês" : "meses"}
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-6">
              <div className="overflow-x-auto">
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Mês</TableHead>
                      <TableHead className="text-right">Faturamento</TableHead>
                      <TableHead className="text-right">Var. Faturamento</TableHead>
                      <TableHead className="text-right">Clientes</TableHead>
                      <TableHead className="text-right">Var. Clientes</TableHead>
                      <TableHead className="text-right">Ticket Médio</TableHead>
                      <TableHead className="text-right">Entraram</TableHead>
                      <TableHead className="text-right">Não compraram</TableHead>
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
                          {i === result.evolucaoPeriodos.length - 1 && (
                            <span className="ml-2 text-xs text-muted-foreground">(atual)</span>
                          )}
                        </TableCell>
                        <TableCell className="text-right font-medium">{fmt(p.faturamento)}</TableCell>
                        <TableCell className="text-right">
                          {p.variacaoFaturamento === null ? (
                            <span className="text-xs text-muted-foreground">—</span>
                          ) : (
                            <VariationBadge value={p.variacaoFaturamento} />
                          )}
                        </TableCell>
                        <TableCell className="text-right">{p.clientes}</TableCell>
                        <TableCell className="text-right">
                          {p.variacaoClientes === null ? (
                            <span className="text-xs text-muted-foreground">—</span>
                          ) : (
                            <VariationBadge value={p.variacaoClientes} />
                          )}
                        </TableCell>
                        <TableCell className="text-right">{fmt(p.ticketMedio)}</TableCell>
                        <TableCell className="text-right text-accent">
                          {p.entraram === null ? "—" : `+${p.entraram}`}
                        </TableCell>
                        <TableCell className="text-right text-destructive">
                          {p.sairam === null ? "—" : `−${p.sairam}`}
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </div>

              <ResponsiveContainer width="100%" height={260}>
                <BarChart data={evolucaoChartData} margin={{ top: 8 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="hsl(220 15% 88%)" />
                  <XAxis dataKey="nome" tick={{ fontSize: 12 }} />
                  <YAxis yAxisId="left" tick={{ fontSize: 11 }} />
                  <YAxis yAxisId="right" orientation="right" tick={{ fontSize: 11 }} />
                  <Tooltip
                    formatter={(v: number, n: string) => (n === "Faturamento" ? fmt(v) : v)}
                    contentStyle={{ borderRadius: "8px", border: "1px solid hsl(220 15% 88%)" }}
                  />
                  <Legend />
                  <Bar yAxisId="left" dataKey="Faturamento" fill="hsl(217 91% 50%)" radius={[4, 4, 0, 0]} />
                  <Bar yAxisId="right" dataKey="Clientes" fill="hsl(160 60% 45%)" radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </CardContent>
          </Card>

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
                  label={`Variação total (${result.evolucaoPeriodos[0].nome} → ${result.periodoAtual.nome})`}
                  value={`${variacaoTotal(result) >= 0 ? "+" : ""}${variacaoTotal(result).toFixed(2)}%`}
                  highlight
                />
                <hr className="border-border" />
                <p className="text-xs font-medium uppercase tracking-wide text-muted-foreground">
                  Carteira · janela de {result.periodos.length}{" "}
                  {result.periodos.length === 1 ? "mês" : "meses"}
                </p>
                <Row label="Clientes ativos" value={String(result.carteira.ativos.length)} />
                <Row label="Clientes novos" value={String(result.carteira.novos.length)} />
                <Row label="Clientes que retornaram" value={String(result.carteira.retornados.length)} />
                <Row label="Frequência reduzida (1 mês sem comprar)" value={String(result.carteira.frequenciaReduzida.length)} />
                <Row label="Clientes perdidos (2+ meses sem comprar)" value={String(result.carteira.perdidos.length)} />
                <Row label="Receita de novos" value={fmt(result.carteira.faturamentoNovos)} />
                <Row label="Receita perdida" value={fmt(result.carteira.faturamentoPerdidos)} />
                <hr className="border-border" />
                <p className="text-xs font-medium uppercase tracking-wide text-muted-foreground">
                  Mês a mês · {result.periodoAnterior.nome} × {result.periodoAtual.nome}
                </p>
                <Row label="Clientes mantidos" value={String(result.clientesMantidos.length)} />
                <Row label="Entraram no mês" value={String(result.clientesNovos.length)} />
                <Row label="Não compraram no mês" value={String(result.clientesPerdidos.length)} />
                <Row label="Receita de entrantes" value={fmt(result.faturamentoNovos)} />
                <Row label="Receita não recorrente" value={fmt(result.faturamentoPerdidos)} />
              </CardContent>
            </Card>

            <Card>
              <CardHeader>
                <CardTitle className="flex items-center gap-2 text-base">
                  <Package className="h-4 w-4 text-primary" /> Top 10 Produtos por Faturamento
                </CardTitle>
              </CardHeader>
              <CardContent>
                <ResponsiveContainer width="100%" height={300}>
                  <BarChart data={chartData} margin={{ bottom: 60 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="hsl(220 15% 88%)" />
                    <XAxis dataKey="nome" angle={-35} textAnchor="end" tick={{ fontSize: 11 }} />
                    <YAxis tick={{ fontSize: 11 }} />
                    <Tooltip
                      formatter={(v: number) => fmt(v)}
                      contentStyle={{
                        borderRadius: "8px",
                        border: "1px solid hsl(220 15% 88%)",
                      }}
                    />
                    <Legend />
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
            Compraram em {result.periodoAtual.nome} e em nenhum dos{" "}
            {janelaLabel(result)} anteriores da janela.
            {result.carteira.retornados.length > 0 && (
              <>
                {" "}
                Outros {result.carteira.retornados.length} clientes voltaram a comprar
                após um intervalo — listados abaixo, não contados como novos.
              </>
            )}
          </p>
          <ClientTable
            clients={result.carteira.novos}
            emptyMessage="Nenhum cliente novo nesta janela."
          />
          {result.carteira.retornados.length > 0 && (
            <>
              <h3 className="flex items-center gap-2 pt-2 text-sm font-semibold">
                <RotateCcw className="h-4 w-4 text-primary" />
                Clientes que retornaram ({result.carteira.retornados.length})
              </h3>
              <ClientTable
                clients={result.carteira.retornados}
                emptyMessage=""
              />
            </>
          )}
        </TabsContent>

        {/* Clientes Perdidos */}
        <TabsContent value="perdidos" className="space-y-4">
          <p className="text-sm text-muted-foreground">
            Sem nenhuma compra há 2 meses ou mais dentro da janela analisada. Quem
            faltou apenas no último mês aparece como frequência reduzida (
            {result.carteira.frequenciaReduzida.length}) na aba Carteira.
          </p>
          <CarteiraTable
            clientes={result.carteira.perdidos}
            totalPeriodos={result.periodos.length}
            emptyMessage="Nenhum cliente perdido nesta janela."
          />
        </TabsContent>

        {/* Comparativo Produtos */}
        <TabsContent value="produtos" className="space-y-4">
          <p className="text-sm text-muted-foreground">
            Faturamento de cada produto em todos os {result.periodos.length} meses
            analisados. As colunas de variação seguem comparando{" "}
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
                    <TableHead className="text-right">Var. Valor</TableHead>
                    <TableHead className="text-right">Qtd Anterior</TableHead>
                    <TableHead className="text-right">Qtd Atual</TableHead>
                    <TableHead className="text-right">Var. Qtd</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {result.comparativoProdutos.map((p) => (
                    <TableRow key={p.produto}>
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

        {/* Carteira 6 Meses */}
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
                        {p.clientes.size}
                      </TableCell>
                      <TableCell className="text-right">{p.produtosLidos}</TableCell>
                      <TableCell className="text-right">{fmt(p.totalFaturamento)}</TableCell>
                    </TableRow>
                  ))}
                  {result.leitura.abasIgnoradas.map((a) => (
                    <TableRow key={a.sheetName} className="text-muted-foreground">
                      <TableCell className="font-medium">{a.sheetName}</TableCell>
                      <TableCell colSpan={6}>Não lida — {a.motivo}</TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </CardContent>
          </Card>

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

/** Older months fade toward grey; the current month gets the strong accent. */
function periodColor(index: number, total: number): string {
  if (index === total - 1) return "hsl(217 91% 50%)";
  const lightness = 82 - (index / Math.max(total - 1, 1)) * 22;
  return `hsl(217 25% ${lightness}%)`;
}

/** Variation between the oldest and the most recent month of the window. */
function variacaoTotal(result: AnalysisResult): number {
  const primeiro = result.evolucaoPeriodos[0];
  const ultimo = result.evolucaoPeriodos[result.evolucaoPeriodos.length - 1];
  if (!primeiro || primeiro.faturamento === 0) return 0;
  return ((ultimo.faturamento - primeiro.faturamento) / primeiro.faturamento) * 100;
}

function janelaLabel(result: AnalysisResult): string {
  const n = result.periodos.length - 1;
  return n === 1 ? "1 mês" : `${n} meses`;
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
        totalPeriodos={periodos.length}
        emptyMessage="Nenhum cliente encontrado."
      />
    </div>
  );
}

function CarteiraTable({
  clientes,
  totalPeriodos,
  emptyMessage,
}: {
  clientes: ClienteCarteira[];
  totalPeriodos: number;
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
      <CardContent className="p-0">
        <Table>
          <TableHeader>
            <TableRow>
              <TableHead>Cliente</TableHead>
              <TableHead>Status</TableHead>
              <TableHead>Última Compra</TableHead>
              <TableHead className="text-right">Meses Presente</TableHead>
              <TableHead>Histórico</TableHead>
              <TableHead className="text-right">Fat. Última Compra</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {clientes.map((c) => (
              <TableRow key={c.nome}>
                <TableCell className="font-medium">{c.nome}</TableCell>
                <TableCell>
                  <Badge variant="outline" className={STATUS_CONFIG[c.status].className}>
                    {STATUS_CONFIG[c.status].label}
                  </Badge>
                </TableCell>
                <TableCell className="text-sm text-muted-foreground">
                  {c.ultimoPeriodo}
                </TableCell>
                <TableCell className="text-right">
                  {c.totalPeriodosPresente}/{totalPeriodos}
                </TableCell>
                <TableCell className="text-xs text-muted-foreground">
                  {c.periodosPresente.join(", ")}
                </TableCell>
                <TableCell className="text-right">
                  {fmt(c.faturamentoUltimoPeriodo)}
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
  clients: { nome: string; codigoCliente?: string; produtos: { produto: string; codigo: string; unidade?: string; quantidade: number; valorTotal: number }[]; totalGeral: number }[];
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
        <Card key={c.nome}>
          <CardHeader className="pb-2">
            <div className="flex items-center justify-between">
              <CardTitle className="text-base">{c.nome}</CardTitle>
              <Badge variant="secondary" className="text-sm font-semibold">
                {fmt(c.totalGeral)}
              </Badge>
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
                {c.produtos.map((p) => (
                  <TableRow key={p.codigo}>
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
