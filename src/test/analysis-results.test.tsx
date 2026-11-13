import { describe, it, expect, vi } from "vitest";
import { render, screen, within, fireEvent } from "@testing-library/react";
import { cloneElement, type ReactElement } from "react";
import { AnalysisResults } from "../components/AnalysisResults";
import { analyzeWorkbook } from "../lib/excel-analyzer";
import { montarArquivo, type AbaFixture } from "./relatorio-fixture";

// Recharts measures its container, which jsdom reports as 0×0 and then renders
// nothing. Hand the chart a fixed size so its bars, ticks and legend really exist.
vi.mock("recharts", async () => {
  const actual = await vi.importActual<typeof import("recharts")>("recharts");
  return {
    ...actual,
    ResponsiveContainer: ({ children }: { children: ReactElement }) =>
      cloneElement(children, { width: 800, height: 400 }),
  };
});

const junho: AbaFixture = {
  intervalo: "01/06/2025 a 30/06/2025",
  clientes: [
    { codigo: "1001", nome: "CLIENTE ALFA", itens: [{ codigo: "5001", produto: "PARAFUSO", qtd: 10, preco: 2.5 }] },
    { codigo: "1002", nome: "CLIENTE BETA", itens: [{ codigo: "5004", produto: "GRAXA", qtd: 3, preco: 10 }] },
  ],
};
const julho: AbaFixture = {
  intervalo: "01/07/2025 a 31/07/2025",
  clientes: [
    { codigo: "1001", nome: "CLIENTE ALFA", itens: [{ codigo: "5001", produto: "PARAFUSO", qtd: 12, preco: 2.5 }] },
  ],
};
const agosto: AbaFixture = {
  intervalo: "01/08/2025 a 31/08/2025",
  clientes: [
    { codigo: "1001", nome: "CLIENTE ALFA", itens: [{ codigo: "5001", produto: "PARAFUSO", qtd: 20, preco: 2.5 }] },
    { codigo: "1002", nome: "CLIENTE BETA", itens: [{ codigo: "5003", produto: "ARRUELA", qtd: 4, preco: 3 }] },
  ],
};

function renderAnalise() {
  const result = analyzeWorkbook(montarArquivo({ junho, julho, agosto }));
  return { result, ...render(<AnalysisResults result={result} />) };
}

/** Radix tabs switch on mouse down, not on a bare click event. */
function abrirAba(nome: RegExp) {
  const trigger = screen.getByRole("tab", { name: nome });
  fireEvent.mouseDown(trigger);
  fireEvent.click(trigger);
}

/** Matches text that the markup splits across several elements. */
const textoContinuo = (trecho: string) => (_: string, el: Element | null) =>
  el?.textContent?.replace(/\s+/g, " ").includes(trecho) ?? false;

describe("AnalysisResults — os três meses chegam à tela", () => {
  it("mostra a janela de análise com junho, julho e agosto", () => {
    renderAnalise();

    expect(screen.getByText(/Janela de análise da carteira: 3 meses/)).toBeInTheDocument();
    for (const mes of ["junho", "julho", "agosto"]) {
      expect(screen.getAllByText(new RegExp(mes)).length).toBeGreaterThan(0);
    }
  });

  it("mostra junho na tabela de comparativo mês a mês", () => {
    renderAnalise();

    const linhaJunho = screen.getByRole("row", { name: /junho/ });
    // junho: 2 clientes × R$ 25 + R$ 30 = R$ 55
    expect(within(linhaJunho).getByText("R$ 55,00")).toBeInTheDocument();
    expect(within(linhaJunho).getByText("2")).toBeInTheDocument();
    // first month of the window: no previous month to compare against
    expect(within(linhaJunho).getAllByText("—").length).toBeGreaterThan(0);
  });

  it("plota junho nos gráficos de evolução e de produtos", () => {
    const { container } = renderAnalise();

    // month-by-month bar chart: one tick per month
    const ticks = [...container.querySelectorAll("text")].map((t) => t.textContent);
    expect(ticks).toContain("junho");
    expect(ticks).toContain("julho");
    expect(ticks).toContain("agosto");

    // the product chart draws one series per month of the window
    const legendas = [...container.querySelectorAll(".recharts-legend-item-text")].map(
      (l) => l.textContent,
    );
    expect(legendas).toEqual(expect.arrayContaining(["junho", "julho", "agosto"]));
  });

  it("mostra na aba Produtos um item vendido só em junho", () => {
    renderAnalise();
    abrirAba(/Comparativo Produtos/);

    const linhaGraxa = screen.getByRole("row", { name: /GRAXA/ });
    expect(within(linhaGraxa).getByText("R$ 30,00")).toBeInTheDocument();
  });

  it("mostra na aba Validação a conferência por mês, sem o quadro de leitura nem o resumo por camada", () => {
    renderAnalise();
    abrirAba(/Validação/);

    expect(screen.queryByText(textoContinuo("Leitura do arquivo"))).not.toBeInTheDocument();
    expect(screen.queryByText("Resumo de Validação por Camada")).not.toBeInTheDocument();

    const linhaJunho = screen.getAllByRole("row").find((r) => r.querySelector("td")?.textContent === "junho")!;
    // junho: 2 clients, 2 products, 10 + 3 units, R$ 55,00
    expect(within(linhaJunho).getAllByText("2").length).toBeGreaterThanOrEqual(2);
    expect(within(linhaJunho).getByText("13")).toBeInTheDocument();
  });
});
