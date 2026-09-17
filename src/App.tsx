import { useState } from "react";
import ExcelJS from "exceljs";
import { AlertCircle, Check, Download } from "lucide-react";

import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import {
  Field,
  FieldDescription,
  FieldGroup,
  FieldLabel,
} from "@/components/ui/field";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import {
  createDownloadFilename,
  parseStudentRecord,
  type StudentData,
} from "@/lib/parser";

const TEMPLATE_PATH = "/Ficha_Alumno_TEMPLATE.xlsx";

type PreviewField = {
  key: keyof StudentData;
  label: string;
  wide?: boolean;
};

type PreviewSection = {
  title: string;
  fields: PreviewField[];
};

const previewSections: PreviewSection[] = [
  {
    title: "Datos personales y usuario",
    fields: [
      { key: "name", label: "Nombre completo" },
      { key: "username", label: "Usuario" },
      { key: "email", label: "Email" },
      { key: "access", label: "Acceso" },
      { key: "twoFactor", label: "Autenticación en dos pasos" },
      { key: "account", label: "Cuenta corriente" },
      { key: "role", label: "Función" },
      { key: "lastLogin", label: "Última vez" },
      { key: "lastRegeneration", label: "Última regeneración" },
    ],
  },
  {
    title: "Datos personales",
    fields: [
      { key: "document", label: "Documento (DNI)" },
      { key: "digitalSignature", label: "Firma digital" },
      { key: "birthProvince", label: "Nacido en provincia" },
      { key: "birthLocality", label: "Nacido en localidad" },
      { key: "cuil", label: "CUIL" },
      { key: "birthDate", label: "Fecha de nacimiento" },
      { key: "nationality", label: "Nacionalidad" },
    ],
  },
  {
    title: "Datos de contacto y domicilio",
    fields: [
      { key: "homePhone", label: "Tel. particular" },
      { key: "mobilePhone", label: "Tel. celular" },
      { key: "address", label: "Domicilio", wide: true },
      { key: "floor", label: "Piso" },
      { key: "apartment", label: "Departamento" },
      { key: "tower", label: "Torre" },
      { key: "locality", label: "Localidad" },
      { key: "postalCode", label: "Código postal" },
    ],
  },
  {
    title: "Datos del alumno",
    fields: [
      { key: "grade", label: "Grado" },
      { key: "enrollment", label: "Matrícula" },
      { key: "studentRecord", label: "Legajo" },
      { key: "bookMatrix", label: "Libro matriz" },
      { key: "folio", label: "Folio" },
      { key: "behavior", label: "Conducta" },
      { key: "previousSchool", label: "Colegio de procedencia", wide: true },
      { key: "tutor", label: "Tutor principal", wide: true },
      { key: "withRecord", label: "Con legajo" },
    ],
  },
  {
    title: "Datos de facturación",
    fields: [
      { key: "taxTreatment", label: "Tratamiento impositivo", wide: true },
      { key: "cuit", label: "CUIT", wide: true },
    ],
  },
  {
    title: "Autorizaciones de retiro",
    fields: [
      { key: "withdrawal1Name", label: "Retira 1" },
      { key: "withdrawal1Dni", label: "DNI retira 1" },
      { key: "withdrawal1Relationship", label: "Parentesco retira 1" },
      { key: "withdrawal2Name", label: "Retira 2" },
      { key: "withdrawal2Dni", label: "DNI retira 2" },
      { key: "withdrawal2Relationship", label: "Parentesco retira 2" },
    ],
  },
  {
    title: "Contactos de emergencia",
    fields: [
      { key: "emergency1Contact", label: "Contacto emergencia 1" },
      { key: "emergency1Phone", label: "Teléfono emerg. 1" },
      { key: "emergency1Relationship", label: "Parentesco emerg. 1" },
      { key: "emergency2Contact", label: "Contacto emergencia 2" },
      { key: "emergency2Phone", label: "Teléfono emerg. 2" },
      { key: "emergency2Relationship", label: "Parentesco emerg. 2" },
    ],
  },
  {
    title: "Asistencia y permisos",
    fields: [
      { key: "personalImage", label: "Imagen personal" },
      { key: "trips", label: "Excursiones" },
      { key: "medicalAttention", label: "Atención médica" },
      { key: "assistance", label: "Ayuda asistencia" },
      { key: "psychologist", label: "Psicopedagogo" },
    ],
  },
  {
    title: "Salud",
    fields: [
      { key: "healthInsurance", label: "Obra social" },
      { key: "bloodType", label: "Grupo sanguíneo" },
      { key: "allergies", label: "Alergias" },
    ],
  },
  {
    title: "Legajo",
    fields: [
      { key: "preEnrollment", label: "Cita preinscripción", wide: true },
      { key: "admission", label: "Admisión", wide: true },
      { key: "previousSchoolRecord", label: "Escuela anterior", wide: true },
    ],
  },
];

const TOTAL_PREVIEW_FIELDS = previewSections.reduce(
  (total, section) => total + section.fields.length,
  0,
);

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

type EditableFieldProps = {
  field: PreviewField;
  value: string | null;
  onChange: (key: keyof StudentData, value: string) => void;
};

function EditableField({ field, value, onChange }: EditableFieldProps) {
  const inputId = `preview-${String(field.key)}`;

  return (
    <Field className={field.wide ? "sm:col-span-2 xl:col-span-3" : undefined}>
      <FieldLabel htmlFor={inputId}>{field.label}</FieldLabel>
      <Input
        id={inputId}
        value={value ?? ""}
        onChange={(event) => onChange(field.key, event.target.value)}
      />
    </Field>
  );
}

export default function App() {
  const [text, setText] = useState("");
  const [editedRecord, setEditedRecord] = useState<StudentData | null>(null);
  const [isDownloading, setIsDownloading] = useState(false);
  const [downloadError, setDownloadError] = useState<string | null>(null);
  const parsedRecord = parseStudentRecord(text);
  const isParsed = parsedRecord.status === "valid" && editedRecord !== null;
  const requiredFieldsPresent = Boolean(
    editedRecord?.name.trim() && editedRecord?.grade.trim(),
  );
  const isReady = isParsed && requiredFieldsPresent;
  const hasParseError = parsedRecord.status !== "empty" && !isParsed;
  const detectedFieldCount = editedRecord
    ? Object.values(editedRecord).filter(Boolean).length
    : 0;

  function handleTextChange(value: string) {
    const nextRecord = parseStudentRecord(value);
    setText(value);
    setEditedRecord(nextRecord.status === "valid" ? nextRecord.record : null);
    setDownloadError(null);
  }

  function handleFieldChange(key: keyof StudentData, value: string) {
    setEditedRecord((current) =>
      current ? { ...current, [key]: value } : current,
    );
    setDownloadError(null);
  }

  async function handleDownload() {
    if (!isReady || !editedRecord) {
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

      writeRecordToWorksheet(worksheet, editedRecord);
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer as BlobPart], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const objectUrl = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = objectUrl;
      link.download = createDownloadFilename(editedRecord.name);
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

  return (
    <main className="min-h-svh bg-background px-6 py-8 text-foreground sm:px-8 lg:h-svh lg:overflow-hidden lg:px-12">
      <div className="mx-auto flex min-h-[calc(100svh-4rem)] max-w-[1600px] flex-col gap-8 lg:h-[calc(100svh-4rem)] lg:min-h-0">
        <header className="flex flex-wrap items-end justify-between gap-4 border-b pb-6">
          <div className="flex flex-col gap-3">
            <p className="text-xs font-medium uppercase tracking-[0.18em] text-muted-foreground">
              Ficha de alumno
            </p>
            <div className="flex flex-col gap-1">
              <h1 className="text-2xl font-semibold tracking-tight sm:text-3xl">
                Revisá y ajustá la ficha.
              </h1>
              <p className="max-w-3xl text-sm text-muted-foreground sm:text-base">
                Pegá el registro, revisá todos los campos y corregí lo que haga
                falta antes de descargar el Excel.
              </p>
            </div>
          </div>
          <Badge variant={isReady ? "default" : "secondary"}>
            {isReady ? "Lista para descargar" : "Esperando datos"}
          </Badge>
        </header>

        <div className="grid min-h-0 flex-1 gap-8 lg:grid-cols-[minmax(280px,0.36fr)_minmax(0,1fr)]">
          <section className="flex min-h-0 flex-col gap-4 lg:border-r lg:border-border lg:pr-8">
            <div className="flex flex-col gap-1">
              <h2 className="text-base font-semibold">Texto de entrada</h2>
              <p className="text-sm text-muted-foreground">
                El parser ignora el ruido del perfil y detecta los campos
                reconocidos.
              </p>
            </div>

            <FieldGroup className="min-h-0 flex-1">
                  <Field
                    className="min-h-0 flex-1"
                    data-invalid={hasParseError || undefined}
                  >
                <FieldLabel htmlFor="student-record">
                  Pegá el contenido de la ficha
                </FieldLabel>
                <Textarea
                  id="student-record"
                  value={text}
                  onChange={(event) => handleTextChange(event.target.value)}
                  placeholder={
                    "Perfil de Nombre Apellido de 6º Grado\n...\n\nNombre completo\nNombre Apellido"
                  }
                  aria-invalid={hasParseError || undefined}
                  className="h-[min(60svh,620px)] max-h-[620px] min-h-0 resize-none overflow-y-auto lg:h-auto lg:max-h-none lg:flex-1"
                />
                <FieldDescription>
                  Se procesa todo localmente en este navegador.
                </FieldDescription>
              </Field>
            </FieldGroup>

            <div aria-live="polite" className="shrink-0 border-t pt-4">
              {parsedRecord.status === "empty" ? (
                <p className="text-sm text-muted-foreground">
                  Pegá un registro para cargar la ficha editable.
                </p>
              ) : isParsed ? (
                <div className="flex items-start gap-3 text-sm">
                  <Check aria-hidden="true" className="mt-0.5 text-primary" />
                  <div className="flex min-w-0 flex-col gap-1">
                    <p className="font-medium">Datos detectados</p>
                    <p className="text-muted-foreground">
                      Podés editar cualquier campo en la previsualización.
                    </p>
                    {!requiredFieldsPresent && (
                      <p className="text-destructive">
                        El nombre y el grado son obligatorios para descargar.
                      </p>
                    )}
                  </div>
                </div>
              ) : (
                <Alert variant="destructive">
                  <AlertCircle aria-hidden="true" />
                  <AlertTitle>No se pudo cargar la ficha</AlertTitle>
                  <AlertDescription>{parsedRecord.message}</AlertDescription>
                </Alert>
              )}
            </div>
          </section>

          <section className="flex min-h-0 min-w-0 flex-col gap-4 overflow-hidden lg:h-full">
            <div className="flex flex-wrap items-baseline justify-between gap-3 border-b pb-4">
              <div className="flex flex-col gap-1">
                <h2 className="text-base font-semibold">
                  Previsualización editable
                </h2>
                <p className="text-sm text-muted-foreground">
                  Los cambios se aplican al archivo que descargues.
                </p>
              </div>
              <p className="text-xs text-muted-foreground">
                {editedRecord
                  ? `${detectedFieldCount}/${TOTAL_PREVIEW_FIELDS} campos`
                  : "Sin datos"}
              </p>
            </div>

                <div className="scroll-fade-y scroll-fade-24 min-h-0 flex-1 overflow-y-auto px-1 pt-1 pb-6">
                  {editedRecord ? (
                    <div className="flex flex-col gap-6">
                      {previewSections.map((section) => (
                        <Card
                          key={section.title}
                          className="gap-0 border border-border/70 py-0 ring-0 shadow-[0_0_0_1px_oklch(0_0_0_/_0.04),0_1px_2px_-1px_oklch(0_0_0_/_0.06),0_2px_4px_0_oklch(0_0_0_/_0.04)] dark:shadow-[0_0_0_1px_oklch(1_0_0_/_0.08)]"
                        >
                          <CardHeader className="flex flex-row items-center justify-between gap-3 border-b bg-muted/20 py-3">
                            <CardTitle className="text-sm font-semibold uppercase tracking-wide">
                              {section.title}
                            </CardTitle>
                            <span className="text-xs text-muted-foreground">
                              {section.fields.length} campos
                            </span>
                          </CardHeader>
                          <CardContent className="pt-5 pb-6">
                            <FieldGroup className="grid gap-x-6 gap-y-4 sm:grid-cols-2 xl:grid-cols-3">
                              {section.fields.map((field) => (
                                <EditableField
                                  key={field.key}
                                  field={field}
                                  value={editedRecord[field.key]}
                                  onChange={handleFieldChange}
                                />
                              ))}
                            </FieldGroup>
                          </CardContent>
                        </Card>
                      ))}
                    </div>
              ) : (
                <p className="pt-2 text-sm text-muted-foreground">
                  Cuando pegues un registro válido, acá vas a poder revisar y
                  editar sus {TOTAL_PREVIEW_FIELDS} campos.
                </p>
              )}
            </div>
          </section>
        </div>

        <footer className="sticky bottom-0 z-10 flex shrink-0 flex-col gap-4 border-t bg-background pt-5 sm:flex-row sm:items-center sm:justify-between">
          <div className="flex flex-col gap-1">
            <p className="text-sm font-medium">Campos detectados</p>
            <p className="text-sm text-muted-foreground">
              {editedRecord
                ? `${detectedFieldCount} de ${TOTAL_PREVIEW_FIELDS} campos listos para revisar.`
                : "Todavía no hay datos para revisar."}
            </p>
          </div>
          <Button
            type="button"
            size="lg"
            onClick={handleDownload}
            disabled={!isReady || isDownloading}
          >
            <Download data-icon="inline-start" />
            {isDownloading ? "Preparando…" : "Descargar ficha"}
          </Button>
        </footer>

        {downloadError && (
          <Alert variant="destructive">
            <AlertCircle aria-hidden="true" />
            <AlertTitle>No se pudo descargar</AlertTitle>
            <AlertDescription>{downloadError}</AlertDescription>
          </Alert>
        )}
      </div>
    </main>
  );
}
