"use client";

import { useRef, useState } from "react";
import { FileSpreadsheet, FileScan } from "lucide-react";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";

type FileDropProps = {
  title: string;
  hint: string;
  accept: string;
  fileName: string | null;
  meta?: string | null;
  disabled?: boolean;
  onFile: (file: File) => void;
};

export function FileDrop({
  title,
  hint,
  accept,
  fileName,
  meta,
  disabled,
  onFile,
}: FileDropProps) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [over, setOver] = useState(false);
  const Icon = accept.includes("pdf") ? FileScan : FileSpreadsheet;

  return (
    <div
      className={cn(
        "rounded-2xl border border-dashed bg-card px-5 py-5 transition-colors",
        over ? "border-primary bg-primary/5" : "border-border",
      )}
      onDragOver={(event) => {
        event.preventDefault();
        if (!disabled) setOver(true);
      }}
      onDragLeave={() => setOver(false)}
      onDrop={(event) => {
        event.preventDefault();
        setOver(false);
        const file = event.dataTransfer.files?.[0];
        if (file && !disabled) onFile(file);
      }}
    >
      <div className="flex items-start gap-3">
        <span className="grid size-10 shrink-0 place-items-center rounded-xl bg-muted text-primary">
          <Icon className="size-5" />
        </span>
        <div className="min-w-0">
          <h2 className="text-base font-semibold">{title}</h2>
          <p className="mt-1 text-sm leading-6 wrap-break-word text-muted-foreground">{hint}</p>
        </div>
      </div>
      {fileName ? (
        <p className="mt-4 truncate text-sm font-medium" title={fileName ?? undefined}>
          {fileName}
          {meta ? <span className="font-normal text-muted-foreground"> · {meta}</span> : null}
        </p>
      ) : (
        <p className="mt-4 text-sm text-muted-foreground">Файл ещё не выбран</p>
      )}
      <Button
        type="button"
        variant="outline"
        className="mt-4"
        disabled={disabled}
        onClick={() => inputRef.current?.click()}
      >
        Выбрать файл
      </Button>
      <input
        ref={inputRef}
        className="sr-only"
        type="file"
        accept={accept}
        aria-label={title}
        disabled={disabled}
        onChange={(event) => {
          const file = event.target.files?.[0];
          event.target.value = "";
          if (file) onFile(file);
        }}
      />
    </div>
  );
}
