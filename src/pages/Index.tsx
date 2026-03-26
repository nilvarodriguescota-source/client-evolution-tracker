import { useState } from "react";
import { FileUpload } from "@/components/FileUpload";
import { AnalysisResults } from "@/components/AnalysisResults";
import { analyzeWorkbook, type AnalysisResult } from "@/lib/excel-analyzer";
import { BarChart3, AlertCircle } from "lucide-react";
import { Card, CardContent } from "@/components/ui/card";

const Index = () => {
  const [result, setResult] = useState<AnalysisResult | null>(null);
  const [error, setError] = useState<string | null>(null);

  const handleFile = (buffer: ArrayBuffer, _fileName: string) => {
    try {
      setError(null);
      const analysis = analyzeWorkbook(buffer);
      setResult(analysis);
    } catch (e: any) {
      setError(e.message || "Erro ao processar planilha.");
      setResult(null);
    }
  };

  return (
    <div className="min-h-screen bg-background">
      {/* Header */}
      <header className="sticky top-0 z-10 border-b border-border bg-card/80 backdrop-blur-md">
        <div className="mx-auto flex max-w-7xl items-center gap-3 px-6 py-4">
          <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-primary text-primary-foreground">
            <BarChart3 className="h-5 w-5" />
          </div>
          <div>
            <h1 className="text-lg font-bold leading-tight text-foreground">
              Análise de Carteira
            </h1>
            <p className="text-xs text-muted-foreground">
              Comparativo de períodos • Clientes &amp; Produtos
            </p>
          </div>
        </div>
      </header>

      <main className="mx-auto max-w-7xl space-y-8 px-6 py-8">
        <FileUpload onFileLoaded={handleFile} />

        {error && (
          <Card className="border-destructive/50 bg-destructive/5">
            <CardContent className="flex items-center gap-3 py-4">
              <AlertCircle className="h-5 w-5 text-destructive" />
              <p className="text-sm text-destructive">{error}</p>
            </CardContent>
          </Card>
        )}

        {result && <AnalysisResults result={result} />}

        {!result && !error && (
          <div className="py-16 text-center">
            <p className="text-lg text-muted-foreground">
              Carregue uma planilha Excel com pelo menos duas abas para começar a análise.
            </p>
            <p className="mt-2 text-sm text-muted-foreground">
              Colunas esperadas: Cliente, Produto, Código do Produto, Valor Unitário, Quantidade, Valor Total por Produto, Total Geral
            </p>
          </div>
        )}
      </main>
    </div>
  );
};

export default Index;
