import { describe, it, expect, vi } from "vitest";
import { render, screen, within, fireEvent } from "@testing-library/react";
import { cloneElement, type ReactElement } from "react";
import { AnalysisResults } from "../components/AnalysisResults";
import { analyzeWorkbook } from "../lib/excel-analyzer";
import { montarArquivo, type AbaFixture, type ClienteFixture } from "./relatorio-fixture";

// Recharts measures its container, which jsdom reports as 0×0.
vi.mock("recharts", async () => {
  const actual = await vi.importActual<typeof import("recharts")>("recharts");
  return {
    ...actual,
    ResponsiveContainer: ({ children }: { children: ReactElement }) =>
      cloneElement(children, { width: 800, height: 400 }),
  };
});

const cli = (codigo: string, nome: string, valor: number, produto = "PRODUTO A"): ClienteFixture => ({
  codigo,
  nome,
  itens: [{ codigo: produto === "PRODUTO A" ? "5001" : "5002", produto, qtd: 1, preco: valor }],
});

/**
 * Seis meses com valores em milhares e centavos, e dois funcionários:
 * *MARIA SILVA compra de janeiro a março; *JOÃO PEREIRA só em junho.
 */
const MESES: Record<string, ClienteFixture[]> = {
  janeiro: [cli("1", "ALFA", 1234.56), cli("2", "BETA", 2000), cli("9", "*MARIA SILVA", 45.9, "CAFE")],
  fevereiro: [cli("1", "ALFA", 1300), cli("2", "BETA", 2100.1), cli("9", "*MARIA SILVA", 45.9, "CAFE")],
  marco: [cli("1", "ALFA", 1400), cli("9", "*MARIA SILVA", 45.9, "CAFE")],
  abril: [cli("1", "ALFA", 1500), cli("3", "GAMA", 750.25)],
  maio: [cli("1", "ALFA", 1600), cli("3", "GAMA", 800)],
  junho: [cli("1", "ALFA", 1700), cli("3", "GAMA", 820), cli("8", "*JOÃO PEREIRA", 99.99, "CAFE")],
};

function renderAnalise() {
  const abas: Record<string, AbaFixture> = Object.fromEntries(
    Object.entries(MESES).map(([m, clientes]) => [m, { clientes }]),
  );
  const result = analyzeWorkbook(montarArquivo(abas));
  return { result, ...render(<AnalysisResults result={result} />) };
}

function abrirAba(nome: RegExp) {
  const trigger = screen.getByRole("tab", { name: nome });
  fireEvent.mouseDown(trigger);
  fireEvent.click(trigger);
}

const fmt = (v: number) =>
  v.toLocaleString("pt-BR", { style: "currency", currency: "BRL", minimumFractionDigits: 2 });

describe("tela · seis meses com funcionários", () => {
  it("mostra no comparativo mês a mês o faturamento apurado de cada aba, com funcionários", () => {
    renderAnalise();
    const esperado = Object.values(MESES).map((cs) =>
      Math.round(cs.reduce((s, c) => s + c.itens[0].preco * 100, 0)) / 100,
    );
    for (const [i, mes] of Object.keys(MESES).entries()) {
      const linha = screen
        .getAllByRole("row")
        .find((r) => r.querySelector("td")?.textContent?.startsWith(mes));
      expect(linha, mes).toBeDefined();
      expect(linha!.textContent).toContain(fmt(esperado[i]));
    }
  });

  it("não lista funcionários como clientes novos nem perdidos", () => {
    renderAnalise();
    abrirAba(/Clientes Novos/);
    expect(screen.queryByText("*JOÃO PEREIRA")).not.toBeInTheDocument();
    expect(screen.getAllByText("GAMA").length).toBeGreaterThan(0);

    abrirAba(/Clientes Perdidos/);
    expect(screen.queryByText("*MARIA SILVA")).not.toBeInTheDocument();
    expect(screen.getAllByText("BETA").length).toBeGreaterThan(0);

    abrirAba(/Carteira/);
    expect(screen.queryByText("*MARIA SILVA")).not.toBeInTheDocument();
    expect(screen.queryByText("*JOÃO PEREIRA")).not.toBeInTheDocument();
  });

  it("explica a variação com a parcela das compras de funcionários", () => {
    renderAnalise();
    // março → abril: *MARIA stopped buying (R$ 45,90)
    expect(screen.getAllByText("Compras de funcionários").length).toBeGreaterThan(0);
  });

  it("tela de conferência sem subtotais impressos, total impresso e status", () => {
    const { result } = renderAnalise();
    abrirAba(/Validação/);

    const titulo = screen.getByText("Conferência dos dados da planilha");
    const cartao = titulo.closest(".rounded-lg") as HTMLElement;
    const cabecalhos = within(cartao)
      .getAllByRole("columnheader")
      .map((h) => h.textContent);
    expect(cabecalhos).toEqual([
      "Aba",
      "Registros",
      "Clientes",
      "Funcionários",
      "Produtos",
      "Qtd. vendida",
      "Faturamento",
    ]);
    expect(screen.queryByText(/Subtotais impressos/i)).not.toBeInTheDocument();
    expect(screen.queryByText(/Total impresso/i)).not.toBeInTheDocument();
    expect(screen.queryByText(/planilha × sistema/i)).not.toBeInTheDocument();
    expect(within(cartao).queryByText("Status")).not.toBeInTheDocument();
    expect(within(cartao).queryByText("Confere")).not.toBeInTheDocument();
    expect(within(cartao).queryByText("Divergente")).not.toBeInTheDocument();

    // one unit per sale in this file: 3 + 3 + 2 + 2 + 2 + 3 sale rows
    expect(result.conferencia.abas.map((a) => a.quantidadeTotal)).toEqual([3, 3, 2, 2, 2, 3]);
    const linhaJan = within(cartao).getAllByRole("row").find((r) => r.textContent?.startsWith("janeiro"))!;
    expect(linhaJan.textContent).toContain(fmt(3280.46));

    expect(screen.getByText("Funcionários (2)")).toBeInTheDocument();
    expect(screen.getAllByText("*MARIA SILVA").length).toBeGreaterThan(0);
  });

  it("mostra a movimentação de cada mês (alta, queda, estável) no comparativo", () => {
    const { result } = renderAnalise();
    const esperado = result.evolucaoPeriodos.map((p) => p.movimentacao);
    // jan → fev +6,4% alta · fev → mar −53% queda · mar → abr +56% alta · abr → mai +7% alta · mai → jun +9% alta
    expect(esperado).toEqual([null, "crescimento", "queda", "crescimento", "crescimento", "crescimento"]);
    const tabela = screen.getByRole("columnheader", { name: "Movimentação" }).closest("table")!;
    const linhas = within(tabela).getAllByRole("row").slice(1);
    const lido = linhas.map((r) => {
      const t = r.textContent ?? "";
      return t.includes("Alta") ? "Alta" : t.includes("Queda") ? "Queda" : t.includes("Estável") ? "Estável" : "—";
    });
    expect(lido).toEqual(["—", "Alta", "Queda", "Alta", "Alta", "Alta"]);
    // units sold are shown in the same table
    expect(within(tabela).getByRole("columnheader", { name: "Qtd. Vendida" })).toBeInTheDocument();
  });
});
