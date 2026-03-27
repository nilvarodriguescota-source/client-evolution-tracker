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
import type { AnalysisResult } from "@/lib/excel-analyzer";
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

  const chartData = result.comparativoProdutos.slice(0, 10).map((p) => ({
    nome: p.produto.length > 18 ? p.produto.slice(0, 16) + "…" : p.produto,
    [result.periodoAnterior.nome]: p.valorAnterior,
    [result.periodoAtual.nome]: p.valorAtual,
  }));

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
          value={String(result.periodoAtual.clientes.size)}
          subtitle={`${result.periodoAnterior.clientes.size} no período anterior`}
          icon={Users}
          color="primary"
        />
        <KpiCard
          title="Clientes Novos"
          value={String(result.clientesNovos.length)}
          subtitle={fmt(result.faturamentoNovos) + " em receita"}
          icon={UserPlus}
          color="success"
        />
        <KpiCard
          title="Clientes Perdidos"
          value={String(result.clientesPerdidos.length)}
          subtitle={fmt(result.faturamentoPerdidos) + " em receita perdida"}
          icon={UserMinus}
          color="destructive"
        />
      </div>

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
            Clientes Novos ({result.clientesNovos.length})
          </TabsTrigger>
          <TabsTrigger value="perdidos">
            Clientes Perdidos ({result.clientesPerdidos.length})
          </TabsTrigger>
          <TabsTrigger value="produtos">Comparativo Produtos</TabsTrigger>
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
          <div className="grid gap-6 lg:grid-cols-2">
            <Card>
              <CardHeader>
                <CardTitle className="flex items-center gap-2 text-base">
                  <TrendingUp className="h-4 w-4 text-accent" /> Resumo Executivo
                </CardTitle>
              </CardHeader>
              <CardContent className="space-y-3 text-sm">
                <Row label={`Faturamento ${result.periodoAnterior.nome}`} value={fmt(result.periodoAnterior.totalFaturamento)} />
                <Row label={`Faturamento ${result.periodoAtual.nome}`} value={fmt(result.periodoAtual.totalFaturamento)} />
                <Row label="Variação" value={`${result.variacaoFaturamento >= 0 ? "+" : ""}${result.variacaoFaturamento.toFixed(2)}%`} highlight />
                <hr className="border-border" />
                <Row label="Clientes mantidos" value={String(result.clientesMantidos.length)} />
                <Row label="Clientes novos" value={String(result.clientesNovos.length)} />
                <Row label="Clientes perdidos" value={String(result.clientesPerdidos.length)} />
                <Row label="Receita de novos" value={fmt(result.faturamentoNovos)} />
                <Row label="Receita perdida" value={fmt(result.faturamentoPerdidos)} />
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
                    <Bar dataKey={result.periodoAnterior.nome} fill="hsl(220 15% 75%)" radius={[4, 4, 0, 0]} />
                    <Bar dataKey={result.periodoAtual.nome} fill="hsl(217 91% 50%)" radius={[4, 4, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </CardContent>
            </Card>
          </div>
        </TabsContent>

        {/* Clientes Novos */}
        <TabsContent value="novos">
          <ClientTable
            clients={result.clientesNovos}
            emptyMessage="Nenhum cliente novo neste período."
          />
        </TabsContent>

        {/* Clientes Perdidos */}
        <TabsContent value="perdidos">
          <ClientTable
            clients={result.clientesPerdidos}
            emptyMessage="Nenhum cliente perdido neste período."
          />
        </TabsContent>

        {/* Comparativo Produtos */}
        <TabsContent value="produtos">
          <Card>
            <CardContent className="p-0">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>Produto</TableHead>
                    <TableHead>Código</TableHead>
                    <TableHead className="text-right">Qtd Anterior</TableHead>
                    <TableHead className="text-right">Qtd Atual</TableHead>
                    <TableHead className="text-right">Var. Qtd</TableHead>
                    <TableHead className="text-right">Valor Anterior</TableHead>
                    <TableHead className="text-right">Valor Atual</TableHead>
                    <TableHead className="text-right">Var. Valor</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {result.comparativoProdutos.map((p) => (
                    <TableRow key={p.produto}>
                      <TableCell className="font-medium">{p.produto}</TableCell>
                      <TableCell className="text-muted-foreground">{p.codigo}</TableCell>
                      <TableCell className="text-right">{p.qtdAnterior}</TableCell>
                      <TableCell className="text-right">{p.qtdAtual}</TableCell>
                      <TableCell className="text-right">
                        <VariationBadge value={p.variacaoQtd} />
                      </TableCell>
                      <TableCell className="text-right">{fmt(p.valorAnterior)}</TableCell>
                      <TableCell className="text-right">{fmt(p.valorAtual)}</TableCell>
                      <TableCell className="text-right">
                        <VariationBadge value={p.variacaoValor} />
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </CardContent>
          </Card>
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
