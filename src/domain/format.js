export function formatNumber(value, digits = 3) {
  return new Intl.NumberFormat("es-ES", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
    useGrouping: true
  }).format(value || 0);
}

export function formatCurrency(value) {
  return `${new Intl.NumberFormat("es-ES", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
    useGrouping: true
  }).format(value || 0)} \u20ac`;
}

export function formatPrice(value) {
  return `${new Intl.NumberFormat("es-ES", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
    useGrouping: true
  }).format(value || 0)} \u20ac`;
}

export function roundMoney(value) {
  return Math.round((value || 0) * 100) / 100;
}

export function buildMeasure({ thickness, width, length, fallback }) {
  if (fallback) return fallback;
  const parts = [thickness, width, length].filter(
    (part) => part !== undefined && part !== null && part !== ""
  );
  return parts.map((part) => String(part).replace(".", ",")).join(" x ");
}
