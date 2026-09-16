export type StudentData = {
  name: string;
  grade: string;
  username: string | null;
  email: string | null;
  access: string | null;
  twoFactor: string | null;
  account: string | null;
  role: string | null;
  lastLogin: string | null;
  lastRegeneration: string | null;
  document: string | null;
  digitalSignature: string | null;
  birthProvince: string | null;
  birthLocality: string | null;
  cuil: string | null;
  birthDate: string | null;
  nationality: string | null;
  homePhone: string | null;
  mobilePhone: string | null;
  address: string | null;
  floor: string | null;
  apartment: string | null;
  tower: string | null;
  locality: string | null;
  postalCode: string | null;
  enrollment: string | null;
  studentRecord: string | null;
  bookMatrix: string | null;
  folio: string | null;
  behavior: string | null;
  previousSchool: string | null;
  tutor: string | null;
  taxTreatment: string | null;
  cuit: string | null;
  withdrawal1Name: string | null;
  withdrawal1Dni: string | null;
  withdrawal1Relationship: string | null;
  withdrawal2Name: string | null;
  withdrawal2Dni: string | null;
  withdrawal2Relationship: string | null;
  emergency1Contact: string | null;
  emergency1Phone: string | null;
  emergency1Relationship: string | null;
  emergency2Contact: string | null;
  emergency2Phone: string | null;
  emergency2Relationship: string | null;
  personalImage: string | null;
  trips: string | null;
  medicalAttention: string | null;
  assistance: string | null;
  psychologist: string | null;
  healthInsurance: string | null;
  bloodType: string | null;
  allergies: string | null;
  preEnrollment: string | null;
  admission: string | null;
  previousSchoolRecord: string | null;
  withRecord: string | null;
};

type InvalidResult = {
  status:
    | "empty"
    | "missing-label"
    | "missing-value"
    | "invalid-value"
    | "missing-grade";
  name: string | null;
  grade: string | null;
  record: null;
  message: string;
};

export type StudentParseResult =
  | InvalidResult
  | {
      status: "valid";
      name: string;
      grade: string;
      record: StudentData;
      message: string;
    };

function normalize(value: string) {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeLabel(value: string) {
  return normalize(value).replace(/\s*:\s*$/, "");
}

const FIELD_LABELS = new Set(
  [
    "nombre completo",
    "usuario",
    "email",
    "acceso",
    "autenticacion en dos pasos (?)",
    "cuenta corriente",
    "funcion",
    "ultima vez",
    "ultima regeneracion",
    "documento",
    "firma digital",
    "nacido en provincia",
    "nacido en localidad",
    "cuil",
    "fecha de nacimiento",
    "nacionalidad",
    "tel. particular",
    "tel. celular",
    "domicilio",
    "piso",
    "departamento",
    "torre",
    "localidad",
    "codigo postal",
    "matricula",
    "legajo",
    "libro matriz",
    "folio",
    "conducta",
    "colegio de procedencia",
    "tutor principal",
    "tratamiento impositivo",
    "cuit",
    "retira 1",
    "dni retira 1",
    "parentesco retira 1",
    "retira 2",
    "dni retira 2",
    "parentesco retira 2",
    "contacto emerg 1",
    "telefono emerg 1",
    "parentesco emerg 1",
    "contacto emerg 2",
    "telefono emerg 2",
    "parentesco emerg 2",
    "imagen personal",
    "excursiones",
    "atencion medica",
    "ayuda asistencia",
    "psicopedagogo",
    "obra social",
    "grupo sanguineo",
    "alergias",
    "cita preinscripcion",
    "admision",
    "escuela anterior",
    "con legajo",
  ].map(normalizeLabel),
);

const SECTION_LABELS = new Set(
  [
    "datos personales",
    "datos personales y usuario",
    "datos de contacto",
    "datos de contacto y domicilio",
    "datos del alumno",
    "datos de facturacion",
    "autorizacion retiro",
    "contacto emergencia",
    "asistencia/permisos",
    "salud",
    "legajo",
  ].map(normalizeLabel),
);

const UI_NOISE = new Set(
  [
    "enviar un mensaje regenerar contraseña / carta",
    "administrar funciones",
  ].map(normalizeLabel),
);

function isLabel(value: string) {
  const normalized = normalizeLabel(value);
  return (
    /:\s*$/.test(value) ||
    FIELD_LABELS.has(normalized) ||
    SECTION_LABELS.has(normalized) ||
    UI_NOISE.has(normalized)
  );
}

function findLabelIndex(lines: string[], label: string) {
  const normalizedLabel = normalizeLabel(label);

  return lines.findIndex((line) => {
    if (normalizeLabel(line) !== normalizedLabel) {
      return false;
    }

    // The source has a section heading `LEGAJO` and a field label `Legajo`.
    return !(normalizedLabel === "legajo" && line.trim() === "LEGAJO");
  });
}

function readValueAfter(lines: string[], index: number) {
  const value = lines[index + 1]?.trim() ?? "";
  return value && !isLabel(value) ? value : null;
}

function readField(lines: string[], label: string) {
  const index = findLabelIndex(lines, label);
  return index === -1 ? null : readValueAfter(lines, index);
}

function extractGrade(lines: string[], nameLabelIndex: number) {
  const profileLines = lines.slice(
    0,
    nameLabelIndex === -1 ? Math.min(lines.length, 8) : nameLabelIndex,
  );
  const gradePattern =
    /\bde\s+(\d{1,2})\s*(º|°|ª|\.º|to|ro|do|ero)?\s+grado\b/i;

  for (const line of profileLines) {
    const match = line.match(gradePattern);
    if (match) {
      return `${match[1]}º Grado`;
    }
  }

  return null;
}

function stripParentheses(value: string | null) {
  return value?.replace(/^\((.*)\)$/, "$1").trim() ?? null;
}

function cleanDocument(value: string | null) {
  return value?.replace(/^documento\s+nacional\s+/i, "").trim() ?? null;
}

function formatCuil(value: string | null) {
  if (!value) {
    return null;
  }

  const digits = value.replace(/\D/g, "");
  return digits.length === 11
    ? `${digits.slice(0, 2)}-${digits.slice(2, 10)}-${digits.slice(10)}`
    : value;
}

export function parseStudentRecord(text: string): StudentParseResult {
  if (!text.trim()) {
    return {
      status: "empty",
      name: null,
      grade: null,
      record: null,
      message: "Pegá el texto de la ficha para detectar todos los datos.",
    };
  }

  const lines = text.split(/\r?\n/);
  const nameLabelIndex = findLabelIndex(lines, "Nombre completo");
  const grade = extractGrade(lines, nameLabelIndex);

  if (nameLabelIndex === -1) {
    return {
      status: "missing-label",
      name: null,
      grade,
      record: null,
      message:
        "No encontramos la etiqueta «Nombre completo». Revisá el texto pegado.",
    };
  }

  const name = readValueAfter(lines, nameLabelIndex);

  if (!name) {
    return {
      status: "missing-value",
      name: null,
      grade,
      record: null,
      message:
        "Encontramos «Nombre completo», pero falta el valor en la línea siguiente.",
    };
  }

  if (isLabel(name)) {
    return {
      status: "invalid-value",
      name: null,
      grade,
      record: null,
      message:
        "La línea siguiente parece otra etiqueta, no un nombre. Pegá el valor debajo de «Nombre completo».",
    };
  }

  if (!grade) {
    return {
      status: "missing-grade",
      name,
      grade: null,
      record: null,
      message:
        "Detectamos el nombre, pero falta el grado en el encabezado (por ejemplo: «de 6º Grado»).",
    };
  }

  const record: StudentData = {
    name,
    grade,
    username: readField(lines, "Usuario"),
    email: readField(lines, "Email"),
    access: stripParentheses(readField(lines, "Acceso")),
    twoFactor: readField(lines, "Autenticación en dos pasos (?)"),
    account: readField(lines, "Cuenta Corriente"),
    role: readField(lines, "Función"),
    lastLogin: readField(lines, "Última vez"),
    lastRegeneration: readField(lines, "Última regeneración"),
    document: cleanDocument(readField(lines, "Documento")),
    digitalSignature: readField(lines, "Firma Digital"),
    birthProvince: readField(lines, "Nacido en Provincia"),
    birthLocality: readField(lines, "Nacido en Localidad"),
    cuil: formatCuil(readField(lines, "CUIL")),
    birthDate: readField(lines, "Fecha de nacimiento"),
    nationality: readField(lines, "Nacionalidad"),
    homePhone: readField(lines, "Tel. Particular"),
    mobilePhone: readField(lines, "Tel. Celular"),
    address: readField(lines, "Domicilio"),
    floor: readField(lines, "Piso"),
    apartment: readField(lines, "Departamento"),
    tower: readField(lines, "Torre"),
    locality: readField(lines, "Localidad"),
    postalCode: readField(lines, "Código postal"),
    enrollment: readField(lines, "Matrícula"),
    studentRecord: readField(lines, "Legajo"),
    bookMatrix: readField(lines, "Libro Matriz"),
    folio: readField(lines, "Folio"),
    behavior: readField(lines, "Conducta"),
    previousSchool: readField(lines, "Colegio de procedencia"),
    tutor: readField(lines, "Tutor principal"),
    taxTreatment: readField(lines, "Tratamiento Impositivo"),
    cuit: readField(lines, "CUIT"),
    withdrawal1Name: readField(lines, "RETIRA 1"),
    withdrawal1Dni: readField(lines, "DNI RETIRA 1"),
    withdrawal1Relationship: readField(lines, "PARENTESCO RETIRA 1"),
    withdrawal2Name: readField(lines, "RETIRA 2"),
    withdrawal2Dni: readField(lines, "DNI RETIRA 2"),
    withdrawal2Relationship: readField(lines, "PARENTESCO RETIRA 2"),
    emergency1Contact: readField(lines, "CONTACTO EMERG 1"),
    emergency1Phone: readField(lines, "TELEFONO EMERG 1"),
    emergency1Relationship: readField(lines, "PARENTESCO EMERG 1"),
    emergency2Contact: readField(lines, "CONTACTO EMERG 2"),
    emergency2Phone: readField(lines, "TELEFONO EMERG 2"),
    emergency2Relationship: readField(lines, "PARENTESCO EMERG 2"),
    personalImage: readField(lines, "IMAGEN PERSONAL"),
    trips: readField(lines, "EXCURSIONES"),
    medicalAttention: readField(lines, "ATENCION MEDICA"),
    assistance: readField(lines, "AYUDA ASISTENCIA"),
    psychologist: readField(lines, "PSICOPEDAGOGO"),
    healthInsurance: readField(lines, "OBRA SOCIAL"),
    bloodType: readField(lines, "GRUPO SANGUÍNEO"),
    allergies: readField(lines, "ALERGIAS"),
    preEnrollment: readField(lines, "CITA PREINSCRIPCION"),
    admission: readField(lines, "ADMISION"),
    previousSchoolRecord: readField(lines, "ESCUELA ANTERIOR"),
    withRecord: readField(lines, "CON LEGAJO"),
  };

  return {
    status: "valid",
    name,
    grade,
    record,
    message: "Datos del alumno detectados correctamente.",
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
