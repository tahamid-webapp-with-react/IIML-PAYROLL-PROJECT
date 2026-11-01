// Formatting utilities for text, numbers, dates, and export/print operations

/**
 * Formats a number with comma separators and NO decimal places (rounded to integer).
 * Example: 1250.75 -> "1,251", 1000 -> "1,000", null -> "0"
 */
export function formatNumber(val: number | string | null | undefined): string {
  if (val === null || val === undefined || val === "") return "0";
  const num = typeof val === "number" ? val : parseFloat(String(val).replace(/,/g, ""));
  if (isNaN(num)) return String(val);
  return Math.round(num).toLocaleString("en-US");
}

/**
 * Formats a currency value with comma separators and NO decimal places.
 */
export function formatCurrency(val: number | string | null | undefined, currency = ""): string {
  const formatted = formatNumber(val);
  return currency ? `${currency} ${formatted}` : formatted;
}

/**
 * Clean date and time format
 */
export function formatDate(value: string | null | undefined): string {
  if (!value) return "—";
  const date = new Date(value.length === 10 ? `${value}T00:00:00` : value);
  return Number.isNaN(date.getTime())
    ? value
    : new Intl.DateTimeFormat("en-US", { year: "numeric", month: "short", day: "numeric" }).format(date);
}

export function formatDateTime(value: string | null | undefined): string {
  if (!value) return "—";
  const date = new Date(value);
  return Number.isNaN(date.getTime())
    ? value
    : new Intl.DateTimeFormat("en-US", {
        year: "numeric",
        month: "short",
        day: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: true,
      }).format(date);
}

export function humanize(value: string): string {
  return value
    .replace(/_/g, " ")
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .replace(/\b\w/g, (c) => c.toUpperCase());
}

/**
 * Trigger clean print/PDF generation with attractive ready-to-print layout
 */
export function printDocument(title: string, subtitle: string, headers: string[], rows: (string | number)[][], meta?: Record<string, string>) {
  const printWindow = window.open("", "_blank");
  if (!printWindow) {
    alert("Please allow popups to generate print / PDF preview.");
    return;
  }

  const exportDate = new Date().toLocaleDateString("en-US", { year: "numeric", month: "short", day: "numeric" });
  const exportTime = new Date().toLocaleTimeString("en-US", { hour: "2-digit", minute: "2-digit", hour12: true });

  const metaHtml = meta
    ? Object.entries(meta)
        .map(([k, v]) => `<div><strong>${k}:</strong> ${v}</div>`)
        .join("")
    : "";

  const tableHeaderHtml = headers.map((h) => `<th>${h}</th>`).join("");
  const tableRowsHtml = rows
    .map(
      (row, idx) =>
        `<tr class="${idx % 2 === 1 ? "zebra" : ""}">` +
        row
          .map((cell) => {
            const isNum = typeof cell === "number" || (!isNaN(Number(cell)) && cell !== "" && !isNaN(parseFloat(String(cell))));
            const displayVal = isNum ? formatNumber(cell) : String(cell ?? "");
            return `<td class="${isNum ? "num" : "text"}">${displayVal}</td>`;
          })
          .join("") +
        `</tr>`
    )
    .join("");

  printWindow.document.write(`
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="utf-8" />
        <title>${title}</title>
        <style>
          @page { size: A4 landscape; margin: 12mm 15mm; }
          body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; color: #17221f; margin: 0; padding: 20px; font-size: 11px; }
          .header { border-bottom: 2px solid #0f5b4a; padding-bottom: 12px; margin-bottom: 16px; display: flex; justify-content: space-between; align-items: flex-start; }
          .title h1 { margin: 0 0 4px; font-size: 20px; color: #073b33; }
          .title p { margin: 0; font-size: 12px; color: #5d6e68; }
          .meta { font-size: 11px; color: #5d6e68; text-align: right; display: grid; gap: 3px; }
          .report-meta { background: #f3f7f5; border: 1px solid #cfdcd7; border-radius: 6px; padding: 8px 12px; margin-bottom: 16px; display: flex; flex-wrap: wrap; gap: 18px; font-size: 11px; }
          table { width: 100%; border-collapse: collapse; margin-top: 8px; border: 1px solid #cfdcd7; }
          th { background: #073b33; color: #ffffff; font-weight: 700; text-align: left; padding: 7px 10px; border: 1px solid #073b33; font-size: 10.5px; }
          td { padding: 6px 10px; border: 1px solid #cfdcd7; }
          tr.zebra { background-color: #f7faf8; }
          td.num { text-align: right; font-variant-numeric: tabular-nums; padding-right: 12px; font-weight: 500; }
          td.text { text-align: left; padding-left: 12px; }
          .footer { margin-top: 20px; font-size: 10px; color: #5d6e68; display: flex; justify-content: space-between; border-top: 1px solid #cfdcd7; padding-top: 8px; }
          @media print {
            body { padding: 0; }
            button { display: none; }
          }
        </style>
      </head>
      <body>
        <div class="header">
          <div class="title">
            <h1>${title}</h1>
            <p>${subtitle}</p>
          </div>
          <div class="meta">
            <div><strong>Export Date:</strong> ${exportDate}</div>
            <div><strong>Export Time:</strong> ${exportTime}</div>
          </div>
        </div>
        ${metaHtml ? `<div class="report-meta">${metaHtml}</div>` : ""}
        <table>
          <thead><tr>${tableHeaderHtml}</tr></thead>
          <tbody>${tableRowsHtml}</tbody>
        </table>
        <div class="footer">
          <span>Official Payroll System Export Document</span>
          <span>Printed on ${exportDate} ${exportTime}</span>
        </div>
        <script>
          window.onload = function() {
            setTimeout(function() { window.print(); }, 400);
          };
        </script>
      </body>
    </html>
  `);
  printWindow.document.close();
}
