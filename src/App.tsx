import { useState } from "react";
import ExcelJS from "exceljs";
import { AlertCircle, Check, Download } from "lucide-react";

import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Badge } from "@/components/ui/badge";
import {
  Card,
  CardContent,
  CardDescription,
  CardFooter,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import {
  Field,
  FieldDescription,
  FieldGroup,
  FieldLabel,
} from "@/components/ui/field";
import { Textarea } from "@/components/ui/textarea";
import { Button } from "@/components/ui/button";
import { createDownloadFilename, parseStudentName } from "@/lib/parser";

const TEMPLATE_PATH = "/Ficha_Alumno_TEMPLATE.xlsx";

export default function App() {
  const [text, setText] = useState("");
  const [isDownloading, setIsDownloading] = useState(false);
  const [downloadError, setDownloadError] = useState<string | null>(null);
  const parsedName = parseStudentName(text);
  const isValid = parsedName.status === "valid";
  const hasParseError = parsedName.status !== "empty" && !isValid;

  async function handleDownload() {
    if (!isValid) {
      return;
    }

    setIsDownloading(true);
    setDownloadError(null);

    try {
      const response = await fetch(TEMPLATE_PATH);
      if (!response.ok) {
        throw new Error("No se pudo cargar la plantilla.");
      }

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await response.arrayBuffer());
      const worksheet =
        workbook.getWorksheet("FICHA") ?? workbook.worksheets[0];

      if (!worksheet) {
        throw new Error("La plantilla no contiene una hoja de trabajo.");
      }

      worksheet.getCell("B4").value = parsedName.name;
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer as BlobPart], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const objectUrl = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = objectUrl;
      link.download = createDownloadFilename(parsedName.name);
      link.click();
      URL.revokeObjectURL(objectUrl);
    } catch {
      setDownloadError(
        "No pudimos preparar el archivo. Verificá que la plantilla esté disponible e intentá de nuevo.",
      );
    } finally {
      setIsDownloading(false);
    }
  }

  function handleTextChange(value: string) {
    setText(value);
    setDownloadError(null);
  }

  return (
    <main className="flex min-h-svh items-center justify-center px-4 py-10 sm:px-6">
      <div className="flex w-full max-w-2xl flex-col gap-6">
        <header className="flex flex-col gap-3">
          <Badge
            variant="outline"
            className="w-fit uppercase tracking-[0.18em]"
          >
            Ficha de alumno
          </Badge>
          <div className="flex flex-col gap-2">
            <h1 className="text-3xl font-semibold tracking-tight sm:text-4xl">
              Prepará la ficha en segundos.
            </h1>
            <p className="max-w-xl text-base leading-7 text-muted-foreground">
              Pegá el texto de tu registro y detectamos el nombre completo para
              completar la plantilla de Excel. Todo se procesa en este
              navegador.
            </p>
          </div>
        </header>

        <Card>
          <CardHeader>
            <CardTitle>Texto del registro</CardTitle>
            <CardDescription>
              Por ahora usamos únicamente el valor que aparece debajo de «Nombre
              completo».
            </CardDescription>
          </CardHeader>

          <CardContent className="flex flex-col gap-6">
            <FieldGroup>
              <Field data-invalid={hasParseError || undefined}>
                <FieldLabel htmlFor="student-record">
                  Pegá el contenido de la ficha
                </FieldLabel>
                <Textarea
                  id="student-record"
                  value={text}
                  onChange={(event) => handleTextChange(event.target.value)}
                  placeholder={"Nombre completo:\nNombre Apellido\n\nDNI:\n..."}
                  aria-invalid={hasParseError || undefined}
                  className="h-96 max-h-96 resize-none overflow-y-auto"
                />
                <FieldDescription>
                  La etiqueta puede tener dos puntos y cualquier combinación de
                  mayúsculas o tildes.
                </FieldDescription>
              </Field>
            </FieldGroup>

            <section
              aria-live="polite"
              className="flex flex-col gap-3 rounded-lg border bg-muted/30 p-4"
            >
              <div className="flex flex-wrap items-center justify-between gap-3">
                <p className="text-sm font-medium">Estado de detección</p>
                <Badge variant={isValid ? "default" : "secondary"}>
                  {isValid ? "Listo para descargar" : "Pendiente"}
                </Badge>
              </div>

              {isValid ? (
                <div className="flex items-start gap-3">
                  <Check aria-hidden="true" className="mt-0.5 text-primary" />
                  <div className="flex min-w-0 flex-col gap-1">
                    <p className="text-sm text-muted-foreground">
                      Nombre detectado
                    </p>
                    <p className="break-words text-lg font-medium">
                      {parsedName.name}
                    </p>
                  </div>
                </div>
              ) : (
                <Alert variant={hasParseError ? "destructive" : "default"}>
                  {hasParseError && <AlertCircle aria-hidden="true" />}
                  <AlertTitle>
                    {parsedName.status === "empty"
                      ? "Esperando el texto"
                      : "No hay un nombre válido todavía"}
                  </AlertTitle>
                  <AlertDescription>{parsedName.message}</AlertDescription>
                </Alert>
              )}
            </section>

            {downloadError && (
              <Alert variant="destructive">
                <AlertCircle aria-hidden="true" />
                <AlertTitle>No se pudo descargar</AlertTitle>
                <AlertDescription>{downloadError}</AlertDescription>
              </Alert>
            )}
          </CardContent>

          <CardFooter className="flex-col items-stretch gap-3 sm:flex-row sm:items-center sm:justify-between">
            <p className="text-sm text-muted-foreground">
              Se completa la celda B4 de la hoja FICHA.
            </p>
            <Button
              type="button"
              size="lg"
              onClick={handleDownload}
              disabled={!isValid || isDownloading}
            >
              <Download data-icon="inline-start" />
              {isDownloading ? "Preparando…" : "Descargar ficha"}
            </Button>
          </CardFooter>
        </Card>
      </div>
    </main>
  );
}
