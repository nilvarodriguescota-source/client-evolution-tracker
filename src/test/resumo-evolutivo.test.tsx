import { describe, it, expect, vi } from "vitest";
import { render, screen, within, fireEvent } from "@testing-library/react";
import { cloneElement, type ReactElement } from "react";
import { AnalysisResults } from "../components/AnalysisResults";
import { analyzeWorkbook } from "../lib/excel-analyzer";
import { montarArquivo, type AbaFixture, type ClienteFixture } from "./relatorio-fixture";

vi.mock("recharts", async () => {
  const actual = await vi.importActual<typeof import("recharts")>("recharts");
  return {
    ...actual,
    ResponsiveContainer: ({ children }: { children: ReactElement }) =>
      cloneElement(children, { width: 800, height: 400 }),
  };
});

const cli = (nome: string, valor: number): ClienteFixture => ({
  codigo: nome,
  nome,
  itens: [{ codigo: "5001", produto: "PRODUTO A", qtd: 1, preco: valor }],
});

const abas = (meses: Record<string, ClienteFixture[]>): Record<string, AbaFixture> =>
  Object.fromEntries(Object.entries(meses).map(([m, clientes]) => [m, { clientes }]));

/** Radix tabs switch on mouse down, not on a bare click event. */
function abrirAba(nome: RegExp) {
  const trigger = screen.getByRole("tab", { name: nome });
  fireEvent.mouseDown(trigger);
  fireEvent.click(trigger);
}

const textoContinuo = (trecho: string) => (_: string, el: Element | null) =>
  el?.textContent?.replace(/\s+/g, " ").includes(trecho) ?? false;

const TRES = abas({
  junho: [cli("ALFA", 100), cli("BETA", 200)],
  julho: [cli("ALFA", 120), cli("GAMA", 50)],
  agosto: [cli("ALFA", 90), cli("BETA", 210), cli("GAMA", 50), cli("DELTA", 70)],
});

const SEIS = abas({
  janeiro: [cli("A", 100), cli("B", 100)],
  fevereiro: [cli("A", 110), cli("B", 90)],
  marco: [cli("A", 120), cli("C", 60)],
  abril: [cli("A", 130), cli("C", 70), cli("D", 40)],
  maio: [cli("A", 140), cli("D", 50)],
  junho: [cli("A", 150), cli("B", 120), cli("D", 60)],
});

function renderAnalise(fixtures: Record<string, AbaFixture>) {
  const result = analyzeWorkbook(montarArquivo(fixtures));
  return { result, ...render(<AnalysisResults result={result} />) };
}

describe("Resumo evolutivo · 3 meses", () => {
  it("abre com a evolução da janela do primeiro ao último mês", () => {
    renderAnalise(TRES);

    expect(screen.getAllByText(textoContinuo("Evolução da janela · junho → agosto")).length)
      .toBeGreaterThan(0);
    // 300 (junho) → 420 (agosto) = +40%
    expect(screen.getAllByText("R$ 300,00").length).toBeGreaterThan(0);
    expect(screen.getAllByText("R$ 420,00").length).toBeGreaterThan(0);
    expect(screen.getAllByText("+40.00%").length).toBeGreaterThan(0);
    expect(screen.getAllByText("Crescimento").length).toBeGreaterThan(0);
  });

  it("aponta o que melhorou e o que piorou", () => {
    renderAnalise(TRES);

    const melhorou = screen.getByText("O que melhorou").closest("div")!;
    const piorou = screen.getByText("O que piorou").closest("div")!;

    expect(within(melhorou).getByText(/Faturamento subiu/)).toBeInTheDocument();
    expect(within(melhorou).getByText(/comprou pela primeira vez/)).toBeInTheDocument();
    // julho caiu contra junho: a janela teve mês de alta e mês de queda
    expect(piorou.textContent).toMatch(/ticket médio|Ticket médio|caiu|queda/i);
  });

  it("mostra a movimentação de clientes em cada mês da tabela", () => {
    renderAnalise(TRES);

    const julho = screen.getByRole("row", { name: /^julho/ });
    // julho: 1 novo (GAMA), 1 mantido (ALFA), 1 perdido (BETA), retenção 50%
    expect(within(julho).getByText("R$ 170,00")).toBeInTheDocument();
    expect(within(julho).getByText("+1")).toBeInTheDocument();
    expect(within(julho).getByText("−1")).toBeInTheDocument();
    expect(within(julho).getByText("50%")).toBeInTheDocument();

    const agosto = screen.getByRole("row", { name: /^agosto/ });
    expect(within(agosto).getByText("R$ 420,00")).toBeInTheDocument();
    expect(within(agosto).getByText("100%")).toBeInTheDocument();
  });

  it("explica cada transição pela ponte de receita", () => {
    renderAnalise(TRES);

    expect(screen.getAllByText(textoContinuo("junho → julho")).length).toBeGreaterThan(0);
    expect(screen.getAllByText(textoContinuo("julho → agosto")).length).toBeGreaterThan(0);

    // junho → julho: BETA levou R$ 200 embora
    expect(screen.getAllByText("−R$ 200,00").length).toBeGreaterThan(0);
    // julho → agosto: BETA voltou com R$ 210
    expect(screen.getAllByText("+R$ 210,00").length).toBeGreaterThan(0);
  });

  it("lista entrada e saída de clientes mês a mês na aba Clientes Novos", () => {
    renderAnalise(TRES);
    abrirAba(/Clientes Novos/);

    const linhaJulho = screen.getByRole("row", { name: /junho → julho/ });
    expect(within(linhaJulho).getByText("GAMA")).toBeInTheDocument();

    const linhaAgosto = screen.getByRole("row", { name: /julho → agosto/ });
    // DELTA entrou como novo e BETA voltou
    expect(within(linhaAgosto).getByText(/DELTA/)).toBeInTheDocument();
    expect(within(linhaAgosto).getByText(/BETA/)).toBeInTheDocument();
  });

  it("lista quem saiu em cada mês na aba Clientes Perdidos", () => {
    renderAnalise(TRES);
    abrirAba(/Clientes Perdidos/);

    const linhaJulho = screen.getByRole("row", { name: /junho → julho/ });
    expect(within(linhaJulho).getByText("BETA")).toBeInTheDocument();
    expect(within(linhaJulho).getByText("R$ 200,00")).toBeInTheDocument();
  });
});

describe("Resumo evolutivo · 6 meses", () => {
  it("aplica a mesma leitura às seis abas, sem mudar nada no código", () => {
    const { result } = renderAnalise(SEIS);

    expect(result.periodos).toHaveLength(6);
    expect(screen.getAllByText(textoContinuo("Evolução da janela · janeiro → junho")).length)
      .toBeGreaterThan(0);
    expect(screen.getAllByText(textoContinuo("Comparativo mês a mês · 6 meses")).length)
      .toBeGreaterThan(0);

    // uma linha por mês na tabela de evolução
    for (const mes of ["janeiro", "fevereiro", "marco", "abril", "maio", "junho"]) {
      expect(screen.getByRole("row", { name: new RegExp(`^${mes}`) })).toBeInTheDocument();
    }
  });

  it("mostra as cinco transições entre os seis meses", () => {
    renderAnalise(SEIS);

    for (const par of [
      "janeiro → fevereiro",
      "fevereiro → marco",
      "marco → abril",
      "abril → maio",
      "maio → junho",
    ]) {
      expect(screen.getAllByText(textoContinuo(par)).length).toBeGreaterThan(0);
    }
  });

  it("mostra o faturamento de cada cliente em cada um dos seis meses na Carteira", () => {
    renderAnalise(SEIS);
    abrirAba(/Carteira/);

    const linhaC = screen.getByRole("row", { name: /^C / });
    // C comprou só em março (60) e abril (70)
    expect(within(linhaC).getByText("R$ 60,00")).toBeInTheDocument();
    expect(within(linhaC).getByText("R$ 70,00")).toBeInTheDocument();
    expect(within(linhaC).getByText("2/6")).toBeInTheDocument();
  });

  it("traz tendência do período no comparativo de produtos", () => {
    renderAnalise(SEIS);
    abrirAba(/Comparativo Produtos/);

    const linha = screen.getByRole("row", { name: /PRODUTO A/ });
    expect(within(linha).getByText("6/6")).toBeInTheDocument();
    expect(within(linha).getByText("Crescimento")).toBeInTheDocument();
  });
});
