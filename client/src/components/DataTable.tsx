import { useState, useMemo, type ReactNode } from "react";
import * as XLSX from "xlsx";
import { formatNumber, printDocument } from "../utils/formatters";

export interface Column<T> {
  key: string;
  header: string;
  render?: (row: T) => ReactNode;
  isNumeric?: boolean;
  isImportant?: boolean;
  sortable?: boolean;
  width?: string;
}

export interface DataTableProps<T> {
  data: T[];
  columns: Column<T>[];
  title?: string;
  subtitle?: string;
  searchPlaceholder?: string;
  searchFilter?: (row: T, term: string) => boolean;
  extraFilters?: ReactNode;
  onView?: (row: T) => void;
  onEdit?: (row: T) => void;
  onDelete?: (row: T) => void;
  onImportClick?: () => void;
  onDownloadTemplate?: () => void;
  totals?: Record<string, string | number>;
  emptyMessage?: string;
  customActions?: (row: T) => ReactNode;
  keyExtractor?: (row: T, index: number) => string | number;
  initialPageSize?: number;
  hideToolbar?: boolean;
  rowClassName?: (row: T) => string;
}

export function DataTable<T extends Record<string, any>>({
  data,
  columns,
  title = "Data Records",
  subtitle = "Generated system table",
  searchPlaceholder = "Search records...",
  searchFilter,
  extraFilters,
  onView,
  onEdit,
  onDelete,
  onImportClick,
  onDownloadTemplate,
  totals,
  emptyMessage = "No records found.",
  customActions,
  keyExtractor,
  initialPageSize = 25,
  hideToolbar = false,
  rowClassName,
}: DataTableProps<T>) {
  const [searchTerm, setSearchTerm] = useState("");
  const [pageSize, setPageSize] = useState(initialPageSize);
  const [currentPage, setCurrentPage] = useState(1);
  const [density, setDensity] = useState<"compact" | "normal" | "comfortable">("normal");
  const [printIndication, setPrintIndication] = useState(false);

  // Filtered data based on search term
  const filteredData = useMemo(() => {
    if (!searchTerm.trim()) return data;
    const term = searchTerm.toLowerCase().trim();
    if (searchFilter) {
      return data.filter((row) => searchFilter(row, term));
    }
    return data.filter((row) =>
      Object.values(row).some((val) =>
        String(val ?? "").toLowerCase().includes(term)
      )
    );
  }, [data, searchTerm, searchFilter]);

  // Pagination calculation
  const totalRows = filteredData.length;
  const totalPages = Math.max(1, Math.ceil(totalRows / pageSize));
  const effectivePage = Math.min(currentPage, totalPages);
  const startIndex = (effectivePage - 1) * pageSize;
  const endIndex = Math.min(startIndex + pageSize, totalRows);
  const pageRows = filteredData.slice(startIndex, endIndex);

  // Export to Excel
  const exportExcel = () => {
    const exportCols = columns.filter((c) => c.key !== "select" && c.key !== "actions");
    const headers = exportCols.map((c) => c.header);
    const rows = filteredData.map((row) =>
      exportCols.map((col) => {
        const val = row[col.key];
        if (col.isNumeric && val !== undefined && val !== null) {
          const num = typeof val === "number" ? val : parseFloat(String(val).replace(/,/g, ""));
          return isNaN(num) ? val : Math.round(num);
        }
        return val ?? "";
      })
    );
    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Records");
    const fileName = `${title.toLowerCase().replace(/\s+/g, "_")}_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(workbook, fileName);
  };

  // Export to PDF / Print
  const handlePrint = () => {
    setPrintIndication(true);
    setTimeout(() => {
      setPrintIndication(false);
      const headers = columns.map((c) => c.header);
      const rows = filteredData.map((row) =>
        columns.map((col) => {
          const val = row[col.key];
          if (col.isNumeric) return formatNumber(val);
          return String(val ?? "");
        })
      );
      printDocument(title, subtitle, headers, rows, {
        "Total Records": String(totalRows),
        "Filter Query": searchTerm || "All",
      });
    }, 250);
  };

  const scrollToTop = () => {
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  const hasActions = Boolean(onView || onEdit || onDelete || customActions);

  return (
    <div className={`data-table-container density-${density}`}>
      {/* Table Toolbar */}
      {!hideToolbar && (
        <div className="table-toolbar">
          <div className="table-search-box">
            <span className="search-icon" aria-hidden="true">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
            </span>
            <input
              type="search"
              className="table-search-input"
              placeholder={searchPlaceholder}
              value={searchTerm}
              onChange={(e) => {
                setSearchTerm(e.target.value);
                setCurrentPage(1);
              }}
            />
            {searchTerm && (
              <button
                type="button"
                className="clear-search-btn"
                onClick={() => setSearchTerm("")}
                title="Clear search"
              >
                ✕
              </button>
            )}
          </div>

          {extraFilters && <div className="table-extra-filters">{extraFilters}</div>}

          <div className="table-toolbar-actions">
            {/* Density Selector */}
            <div className="density-toggle" title="Table Density">
              <button
                type="button"
                className={`density-btn ${density === "compact" ? "active" : ""}`}
                onClick={() => setDensity("compact")}
                title="Compact View"
              >
                C
              </button>
              <button
                type="button"
                className={`density-btn ${density === "normal" ? "active" : ""}`}
                onClick={() => setDensity("normal")}
                title="Normal View"
              >
                N
              </button>
              <button
                type="button"
                className={`density-btn ${density === "comfortable" ? "active" : ""}`}
                onClick={() => setDensity("comfortable")}
                title="Comfort View"
              >
                +
              </button>
            </div>

            {/* Print Button */}
            <button
              type="button"
              className="tool-btn print-btn"
              onClick={handlePrint}
              title="Print Table Document"
            >
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M6 9V2h12v7"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect width="12" height="8" x="6" y="14"/></svg>
              <span>{printIndication ? "Printing..." : "Print"}</span>
            </button>

            {/* Export Excel Button */}
            <button
              type="button"
              className="tool-btn excel-btn"
              onClick={exportExcel}
              title="Export as Excel XLSX"
            >
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><path d="M8 13h8M8 17h8M10 9h4"/></svg>
              <span>Excel</span>
            </button>

            {/* Export PDF Button */}
            <button
              type="button"
              className="tool-btn pdf-btn"
              onClick={handlePrint}
              title="Export as Ready-to-Print PDF"
            >
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" x2="12" y1="15" y2="3"/></svg>
              <span>PDF</span>
            </button>

            {/* Download Blank Template */}
            {onDownloadTemplate && (
              <button
                type="button"
                className="tool-btn template-btn"
                onClick={onDownloadTemplate}
                title="Download Blank Template for Upload"
              >
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" x2="12" y1="15" y2="3"/></svg>
                <span>Template</span>
              </button>
            )}

            {/* Bulk Upload / Import Button */}
            {onImportClick && (
              <button
                type="button"
                className="tool-btn import-btn"
                onClick={onImportClick}
                title="Bulk Import Data"
              >
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" x2="12" y1="3" y2="15"/></svg>
                <span>Import</span>
              </button>
            )}
          </div>
        </div>
      )}

      {/* Main Data Table */}
      <div className="table-responsive-wrapper">
        <table className="enterprise-data-table">
          <thead>
            <tr>
              <th className="row-number-th">#</th>
              {columns.map((col) => (
                <th
                  key={col.key}
                  style={col.width ? { width: col.width } : undefined}
                  className={col.isNumeric ? "text-right" : "text-left"}
                >
                  {col.header}
                </th>
              ))}
              {hasActions && <th className="actions-th text-center">Actions</th>}
            </tr>
          </thead>
          <tbody>
            {pageRows.length === 0 ? (
              <tr>
                <td
                  colSpan={columns.length + (hasActions ? 2 : 1)}
                  className="table-empty-td"
                >
                  {emptyMessage}
                </td>
              </tr>
            ) : (
              pageRows.map((row, index) => {
                const key = keyExtractor ? keyExtractor(row, index) : row.id ?? index;
                const rowNum = startIndex + index + 1;
                return (
                  <tr
                    key={key}
                    className={`${index % 2 === 1 ? "zebra-row" : "normal-row"} ${rowClassName ? rowClassName(row) : ""}`.trim()}
                  >
                    <td className="row-number-td">{rowNum}</td>
                    {columns.map((col) => {
                      const rawVal = row[col.key];
                      const rendered = col.render ? col.render(row) : (
                        col.isNumeric ? formatNumber(rawVal) : String(rawVal ?? "—")
                      );
                      return (
                        <td
                          key={col.key}
                          className={`${col.isNumeric ? "cell-numeric" : "cell-text"} ${col.isImportant ? "cell-important" : ""}`}
                        >
                          {rendered}
                        </td>
                      );
                    })}
                    {hasActions && (
                      <td className="actions-td text-center">
                        <div className="table-actions-group">
                          {onView && (
                            <button
                              type="button"
                              className="action-icon-btn view-btn"
                              onClick={() => onView(row)}
                              title="View details"
                              aria-label="View details"
                            >
                              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M2 12s3-7 10-7 10 7 10 7-3 7-10 7-10-7-10-7Z"/><circle cx="12" cy="12" r="3"/></svg>
                            </button>
                          )}
                          {onEdit && (
                            <button
                              type="button"
                              className="action-icon-btn edit-btn"
                              onClick={() => onEdit(row)}
                              title="Edit record"
                              aria-label="Edit record"
                            >
                              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                            </button>
                          )}
                          {onDelete && (
                            <button
                              type="button"
                              className="action-icon-btn delete-btn"
                              onClick={() => onDelete(row)}
                              title="Delete record"
                              aria-label="Delete record"
                            >
                              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                            </button>
                          )}
                          {customActions && customActions(row)}
                        </div>
                      </td>
                    )}
                  </tr>
                );
              })
            )}
          </tbody>
          {totals && Object.keys(totals).length > 0 && (
            <tfoot>
              <tr className="table-total-row">
                <td className="row-number-td font-bold">Total</td>
                {columns.map((col) => (
                  <td
                    key={col.key}
                    className={col.isNumeric ? "cell-numeric font-bold" : "cell-text font-bold"}
                  >
                    {totals[col.key] !== undefined
                      ? (col.isNumeric ? formatNumber(totals[col.key]) : totals[col.key])
                      : ""}
                  </td>
                ))}
                {hasActions && <td />}
              </tr>
            </tfoot>
          )}
        </table>
      </div>

      {/* Pagination & Footer Info */}
      <div className="table-footer-pagination">
        <div className="table-info-section">
          <span>
            Showing <strong>{totalRows > 0 ? startIndex + 1 : 0}</strong> to{" "}
            <strong>{endIndex}</strong> of <strong>{totalRows}</strong> records
          </span>
          <div className="page-size-selector">
            <label htmlFor="pageSizeSelect">Show:</label>
            <select
              id="pageSizeSelect"
              value={pageSize}
              onChange={(e) => {
                setPageSize(Number(e.target.value));
                setCurrentPage(1);
              }}
            >
              <option value={10}>10</option>
              <option value={25}>25</option>
              <option value={50}>50</option>
              <option value={100}>100</option>
            </select>
          </div>
        </div>

        <div className="pagination-controls">
          {/* First Page */}
          <button
            type="button"
            className="pagination-icon-btn"
            disabled={effectivePage <= 1}
            onClick={() => setCurrentPage(1)}
            title="First page"
            aria-label="First page"
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="11 17 6 12 11 7"/><polyline points="18 17 13 12 18 7"/></svg>
          </button>

          {/* Previous Page */}
          <button
            type="button"
            className="pagination-icon-btn"
            disabled={effectivePage <= 1}
            onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
            title="Previous page"
            aria-label="Previous page"
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="15 18 9 12 15 6"/></svg>
          </button>

          <span className="pagination-current-page">
            {effectivePage} / {totalPages}
          </span>

          {/* Next Page */}
          <button
            type="button"
            className="pagination-icon-btn"
            disabled={effectivePage >= totalPages}
            onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
            title="Next page"
            aria-label="Next page"
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="9 18 15 12 9 6"/></svg>
          </button>

          {/* Last Page */}
          <button
            type="button"
            className="pagination-icon-btn"
            disabled={effectivePage >= totalPages}
            onClick={() => setCurrentPage(totalPages)}
            title="Last page"
            aria-label="Last page"
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="13 17 18 12 13 7"/><polyline points="6 17 11 12 6 7"/></svg>
          </button>

          {/* Back to top button */}
          <button
            type="button"
            className="back-to-top-btn"
            onClick={scrollToTop}
            title="Back to top"
            aria-label="Back to top"
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="12" y1="19" x2="12" y2="5"/><polyline points="5 12 12 5 19 12"/></svg>
          </button>
        </div>
      </div>
    </div>
  );
}
