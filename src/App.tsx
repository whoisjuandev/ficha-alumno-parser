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
import {
  createDownloadFilename,
  parseStudentRecord,
  type StudentData,
} from "@/lib/parser";

const TEMPLATE_PATH = "/Ficha_Alumno_TEMPLATE.xlsx";

function joinValues(...values: Array<string | null>) {
  const presentValues = values.filter((value): value is string =>
    Boolean(value),
  );
  return presentValues.length > 0 ? presentValues.join(" / ") : null;
}

function writeRecordToWorksheet(
  worksheet: ExcelJS.Worksheet,
  record: StudentData,
) {
  const values: Record<string, string | null> = {
    B4: record.name,
    B5: record.username,
    D5: record.email,
    B6: record.document,
    D6: record.cuil,
    B7: record.birthProvince,
    D7: record.birthLocality,
    B8: record.birthDate,
    D8: record.nationality,
    B9: record.access,
    D9: record.account,
    B12: record.address,
    B13: record.locality,
    D13: record.postalCode,
    B14: record.homePhone,
    D14: record.mobilePhone,
    B17: record.enrollment,
    D17: record.withRecord,
    B18: joinValues(record.bookMatrix, record.folio),
    D18: record.lastRegeneration,
    B19: record.grade,
    B21: record.withdrawal1Name,
    D21: joinValues(record.withdrawal1Dni, record.withdrawal1Relationship),
    B22: record.withdrawal2Name,
    D22: joinValues(record.withdrawal2Dni, record.withdrawal2Relationship),
    B25: record.emergency1Contact,
    B26: joinValues(record.emergency1Phone, record.emergency1Relationship),
    B27: record.emergency2Contact,
    B28: joinValues(record.emergency2Phone, record.emergency2Relationship),
    B31: record.healthInsurance,
    D31: record.bloodType,
    B32: record.allergies,
    D32: record.assistance,
    B33: record.personalImage,
    D33: record.trips,
    B34: record.medicalAttention,
    D34: record.psychologist,
    B37: record.twoFactor,
    D37: record.role,
    B38: record.lastLogin,
    D38: record.digitalSignature,
    B39: record.floor,
    D39: record.apartment,
    B40: record.tower,
    D40: record.studentRecord,
    B41: record.bookMatrix,
    D41: record.folio,
    B42: record.behavior,
    D42: record.previousSchool,
    B43: record.tutor,
    B44: record.taxTreatment,
    B45: record.cuit,
    B48: record.preEnrollment,
    B49: record.admission,
    B50: record.previousSchoolRecord,
  };

  for (const [cell, value] of Object.entries(values)) {
    worksheet.getCell(cell).value = value;
  }
}

export default function App() {
  const [text, setText] = useState("");
  const [isDownloading, setIsDownloading] = useState(false);
  const [downloadError, setDownloadError] = useState<string | null>(null);
  const parsedRecord = parseStudentRecord(text);
  const isValid = parsedRecord.status === "valid";
  const hasParseError = parsedRecord.status !== "empty" && !isValid;

  async function handleDownload() {
    if (parsedRecord.status !== "valid") {
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

      writeRecordToWorksheet(worksheet, parsedRecord.record);
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer as BlobPart], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const objectUrl = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = objectUrl;
      link.download = createDownloadFilename(parsedRecord.name);
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
              Pegá el texto de tu registro y detectamos el nombre completo y el
              grado para completar la plantilla de Excel. Todo se procesa en
              este navegador.
            </p>
          </div>
        </header>

        <Card>
          <CardHeader>
            <CardTitle>Texto del registro</CardTitle>
            <CardDescription>
              Detectamos todos los campos disponibles desde el texto pegado para
              completar la plantilla de Excel.
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
                  Se completan todos los campos reconocidos; el nombre y el
                  grado son obligatorios para descargar la ficha.
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
                <div className="grid gap-4 sm:grid-cols-2">
                  <div className="flex items-start gap-3">
                    <Check aria-hidden="true" className="mt-0.5 text-primary" />
                    <div className="flex min-w-0 flex-col gap-1">
                      <p className="text-sm text-muted-foreground">
                        Nombre detectado
                      </p>
                      <p className="break-words text-lg font-medium">
                        {parsedRecord.name}
                      </p>
                    </div>
                  </div>
                  <div className="flex items-start gap-3">
                    <Check aria-hidden="true" className="mt-0.5 text-primary" />
                    <div className="flex min-w-0 flex-col gap-1">
                      <p className="text-sm text-muted-foreground">
                        Grado detectado
                      </p>
                      <p className="break-words text-lg font-medium">
                        {parsedRecord.grade}
                      </p>
                    </div>
                  </div>
                </div>
              ) : (
                <Alert variant={hasParseError ? "destructive" : "default"}>
                  {hasParseError && <AlertCircle aria-hidden="true" />}
                  <AlertTitle>
                    {parsedRecord.status === "empty"
                      ? "Esperando el texto"
                      : parsedRecord.status === "missing-grade"
                        ? "Falta el grado"
                        : "No hay datos válidos todavía"}
                  </AlertTitle>
                  <AlertDescription>{parsedRecord.message}</AlertDescription>
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
              Se completan todos los campos reconocidos de la hoja FICHA.
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
