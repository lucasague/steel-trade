export function fieldsOf(record) {
  return record?.fields || record?.cellValuesByFieldId || {};
}

export function recordIds(value) {
  if (!value) {
    return [];
  }
  if (Array.isArray(value)) {
    return value
      .map((entry) => {
        if (typeof entry === "string") return entry;
        if (entry && typeof entry === "object") return entry.id;
        return null;
      })
      .filter(Boolean);
  }
  if (Array.isArray(value.linkedRecordIds)) {
    return value.linkedRecordIds;
  }
  return [];
}

export function textValue(value, separator = " / ") {
  const values = flatten(value)
    .map((entry) => {
      if (entry === null || entry === undefined) return "";
      if (typeof entry === "object" && "name" in entry) return entry.name;
      if (typeof entry === "object" && "label" in entry) return entry.label;
      if (typeof entry === "object" && "error" in entry) return "";
      return String(entry);
    })
    .map((entry) => entry.trim())
    .filter((entry) => entry && !isAirtableRecordId(entry));

  return [...new Set(values)].join(separator);
}

export function firstText(value) {
  return textValue(value, " / ");
}

export function numberValue(value) {
  const entry = flatten(value).find((item) => item !== null && item !== undefined);
  if (typeof entry === "number") return entry;
  if (typeof entry === "string") {
    const normalized = entry.replace(/\./g, "").replace(",", ".");
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : undefined;
  }
  return undefined;
}

export function dateValue(value) {
  const text = firstText(value);
  if (!text) return undefined;
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (match) {
    const [, day, month, year] = match;
    return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
  }
  return text;
}

export function formatDateEs(dateLike) {
  const date = dateValue(dateLike) || dateLike;
  if (!date) return "";
  const match = String(date).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return String(date);
  return `${match[3]}/${match[2]}/${match[1]}`;
}

function flatten(value) {
  if (value === null || value === undefined) {
    return [];
  }
  if (Array.isArray(value)) {
    return value.flatMap(flatten);
  }
  if (value && typeof value === "object" && value.valuesByLinkedRecordId) {
    return Object.values(value.valuesByLinkedRecordId).flatMap(flatten);
  }
  if (value && typeof value === "object" && "name" in value) {
    return [value.name];
  }
  return [value];
}

export function isAirtableRecordId(value) {
  return /^rec[A-Za-z0-9]{14}$/.test(String(value || ""));
}
