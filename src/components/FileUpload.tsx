import { useCallback, useState } from "react";
import { Upload, FileSpreadsheet } from "lucide-react";
import { Card } from "@/components/ui/card";

interface FileUploadProps {
  onFileLoaded: (buffer: ArrayBuffer, fileName: string) => void;
}

export function FileUpload({ onFileLoaded }: FileUploadProps) {
  const [dragOver, setDragOver] = useState(false);
  const [fileName, setFileName] = useState<string | null>(null);

  const handleFile = useCallback(
    (file: File) => {
      setFileName(file.name);
      const reader = new FileReader();
      reader.onload = (e) => {
        onFileLoaded(e.target!.result as ArrayBuffer, file.name);
      };
      reader.readAsArrayBuffer(file);
    },
    [onFileLoaded]
  );

  return (
    <Card
      className={`relative cursor-pointer border-2 border-dashed p-12 text-center transition-all ${
        dragOver
          ? "border-primary bg-primary/5 scale-[1.01]"
          : "border-border hover:border-primary/50 hover:bg-muted/50"
      }`}
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragOver(false);
        const file = e.dataTransfer.files[0];
        if (file) handleFile(file);
      }}
      onClick={() => {
        const input = document.createElement("input");
        input.type = "file";
        input.accept = ".xlsx,.xls";
        input.onchange = () => {
          const file = input.files?.[0];
          if (file) handleFile(file);
        };
        input.click();
      }}
    >
      {fileName ? (
        <div className="flex flex-col items-center gap-3">
          <FileSpreadsheet className="h-12 w-12 text-primary" />
          <p className="text-lg font-semibold text-foreground">{fileName}</p>
          <p className="text-sm text-muted-foreground">Clique para trocar o arquivo</p>
        </div>
      ) : (
        <div className="flex flex-col items-center gap-3">
          <Upload className="h-12 w-12 text-muted-foreground" />
          <p className="text-lg font-semibold text-foreground">
            Arraste sua planilha Excel aqui
          </p>
          <p className="text-sm text-muted-foreground">
            ou clique para selecionar um arquivo .xlsx
          </p>
        </div>
      )}
    </Card>
  );
}
