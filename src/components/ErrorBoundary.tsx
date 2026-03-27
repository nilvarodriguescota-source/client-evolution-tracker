import { Component, type ErrorInfo, type ReactNode } from "react";

interface Props {
  children: ReactNode;
}

interface State {
  hasError: boolean;
  error: Error | null;
}

export class ErrorBoundary extends Component<Props, State> {
  constructor(props: Props) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error: Error): State {
    return { hasError: true, error };
  }

  componentDidCatch(error: Error, info: ErrorInfo) {
    console.error("[ErrorBoundary]", error, info.componentStack);
  }

  render() {
    if (this.state.hasError) {
      return (
        <div style={{ padding: "2rem", fontFamily: "system-ui, sans-serif", maxWidth: 600, margin: "4rem auto" }}>
          <h1 style={{ color: "#dc2626", fontSize: "1.5rem", marginBottom: "1rem" }}>
            Erro inesperado na aplicação
          </h1>
          <p style={{ color: "#64748b", marginBottom: "1rem" }}>
            Ocorreu um erro ao renderizar a página. Tente recarregar.
          </p>
          <pre style={{
            background: "#f1f5f9",
            padding: "1rem",
            borderRadius: "0.5rem",
            fontSize: "0.85rem",
            overflow: "auto",
            color: "#334155",
          }}>
            {this.state.error?.message}
          </pre>
          <button
            onClick={() => window.location.reload()}
            style={{
              marginTop: "1rem",
              padding: "0.5rem 1.5rem",
              background: "#2563eb",
              color: "#fff",
              border: "none",
              borderRadius: "0.375rem",
              cursor: "pointer",
              fontSize: "0.9rem",
            }}
          >
            Recarregar página
          </button>
        </div>
      );
    }

    return this.props.children;
  }
}
