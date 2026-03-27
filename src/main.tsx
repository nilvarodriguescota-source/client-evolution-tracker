import { createRoot } from "react-dom/client";
import { ErrorBoundary } from "./components/ErrorBoundary";
import App from "./App.tsx";
import "./index.css";

const rootElement = document.getElementById("root");

if (!rootElement) {
  document.body.innerHTML =
    '<div style="padding:2rem;font-family:system-ui;color:#dc2626">' +
    "<h1>Erro crítico</h1>" +
    '<p>Elemento root não encontrado no DOM. Verifique o index.html.</p>' +
    "</div>";
} else {
  createRoot(rootElement).render(
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  );
}
