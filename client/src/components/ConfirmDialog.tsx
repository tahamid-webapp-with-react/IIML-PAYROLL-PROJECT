export interface ConfirmDialogProps {
  isOpen: boolean;
  title: string;
  message: string;
  confirmLabel?: string;
  confirmText?: string;
  cancelLabel?: string;
  cancelText?: string;
  isDangerous?: boolean;
  danger?: boolean;
  onConfirm: () => void;
  onCancel: () => void;
  isLoading?: boolean;
}

export function ConfirmDialog({
  isOpen,
  title,
  message,
  confirmLabel,
  confirmText,
  cancelLabel,
  cancelText,
  isDangerous,
  danger,
  onConfirm,
  onCancel,
  isLoading = false,
}: ConfirmDialogProps) {
  const resolvedConfirm = confirmText ?? confirmLabel ?? "Confirm";
  const resolvedCancel = cancelText ?? cancelLabel ?? "Cancel";
  const resolvedDangerous = danger ?? isDangerous ?? true;

  if (!isOpen) return null;

  return (
    <div className="modal-backdrop" role="presentation" onMouseDown={(e) => { if (e.target === e.currentTarget && !isLoading) onCancel(); }}>
      <div className="confirm-dialog-card" role="dialog" aria-modal="true" aria-labelledby="confirm-dialog-title">
        <div className={`confirm-dialog-icon ${resolvedDangerous ? "danger" : "info"}`}>
          {resolvedDangerous ? (
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
          ) : (
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>
          )}
        </div>
        <div className="confirm-dialog-content">
          <h3 id="confirm-dialog-title">{title}</h3>
          <p>{message}</p>
        </div>
        <div className="confirm-dialog-actions">
          <button
            type="button"
            className="ghost-button"
            onClick={onCancel}
            disabled={isLoading}
          >
            {resolvedCancel}
          </button>
          <button
            type="button"
            className={resolvedDangerous ? "danger-button" : "primary-button"}
            onClick={onConfirm}
            disabled={isLoading}
          >
            {isLoading ? "Processing..." : resolvedConfirm}
          </button>
        </div>
      </div>
    </div>
  );
}
