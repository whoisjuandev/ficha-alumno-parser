export type NameParseResult =
  | { status: "empty"; name: null; message: string }
  | { status: "missing-label"; name: null; message: string }
  | { status: "missing-value"; name: null; message: string }
  | { status: "invalid-value"; name: null; message: string }
  | { status: "valid"; name: string; message: string };

function normalize(value: string) {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function looksLikeLabel(value: string) {
  const normalized = normalize(value);

  if (normalized.includes(":")) {
    return true;
  }

  return /^(nombre|apellido|dni|documento|fecha|domicilio|telefono|email|correo|curso|grado|turno|escuela|institucion|legajo|cuil|localidad|provincia|nacionalidad|sexo|edad|responsable|contacto)\b/.test(
    normalized,
  );
}

export function parseStudentName(text: string): NameParseResult {
  if (!text.trim()) {
    return {
      status: "empty",
      name: null,
      message: "Pegá el texto de la ficha para detectar el nombre.",
    };
  }

  const lines = text.split(/\r?\n/);
  const labelIndex = lines.findIndex((line) =>
    /^nombre completo\s*:?\s*$/.test(normalize(line)),
  );

  if (labelIndex === -1) {
    return {
      status: "missing-label",
      name: null,
      message:
        "No encontramos la etiqueta «Nombre completo». Revisá el texto pegado.",
    };
  }

  const value = lines[labelIndex + 1]?.trim() ?? "";

  if (!value) {
    return {
      status: "missing-value",
      name: null,
      message:
        "Encontramos «Nombre completo», pero falta el valor en la línea siguiente.",
    };
  }

  if (looksLikeLabel(value)) {
    return {
      status: "invalid-value",
      name: null,
      message:
        "La línea siguiente parece otra etiqueta, no un nombre. Pegá el valor debajo de «Nombre completo».",
    };
  }

  return {
    status: "valid",
    name: value,
    message: "Nombre detectado correctamente.",
  };
}

export function createDownloadFilename(name: string) {
  const safeName = name
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 80);

  return `Ficha_Alumno_${safeName || "Alumno"}.xlsx`;
}
