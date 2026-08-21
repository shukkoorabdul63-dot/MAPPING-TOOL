import React from "react";

export function StatRow({ label, value, mono = true }) {
  return (
    <div className="stat-row">
      <span className="stat-label">{label}</span>
      <span className={mono ? "stat-value mono" : "stat-value"}>{value}</span>
    </div>
  );
}

export function Toggle({ checked, onChange, label, description }) {
  return (
    <div className="toggle-row">
      <div className="toggle-copy">
        <span className="toggle-label">{label}</span>
        {description && <span className="toggle-desc">{description}</span>}
      </div>
      <button
        type="button"
        role="switch"
        aria-checked={checked}
        className={`switch${checked ? " on" : ""}`}
        onClick={() => onChange(!checked)}
      >
        <span className="switch-thumb" />
      </button>
    </div>
  );
}

export function Button({ variant = "primary", size = "md", children, ...rest }) {
  const cls = `btn btn-${variant}${size === "sm" ? " btn-sm" : ""}`;
  return (
    <button className={cls} {...rest}>
      {children}
    </button>
  );
}

export function EmptyStateIcon({ done }) {
  return (
    <span className="empty-icon">
      {done ? (
        <svg viewBox="0 0 24 24" fill="none">
          <path d="M4 12l5 5L19 6" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" />
        </svg>
      ) : (
        <svg viewBox="0 0 24 24" fill="none">
          <circle cx="12" cy="12" r="7" stroke="currentColor" strokeWidth="1.8" />
        </svg>
      )}
    </span>
  );
}

export function StatusRow({ status, labels }) {
  if (status === "idle") return null;
  return (
    <div className="status-row">
      <span className={`dot ${status}`} />
      <span>{labels[status]}</span>
    </div>
  );
}

export function Dropzone({
  onFile,
  fileName,
  emptyTitle,
  emptyHint,
  filledHint = "Click to choose a different file",
  accept = ".xlsx,.xls",
}) {
  const [dragOver, setDragOver] = React.useState(false);
  const inputRef = React.useRef(null);
  return (
    <div
      className={`dropzone${dragOver ? " drag" : ""}`}
      onClick={() => inputRef.current?.click()}
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragOver(false);
        const file = e.dataTransfer.files?.[0];
        if (file) onFile(file);
      }}
    >
      <p className="dropzone-title">{fileName || emptyTitle}</p>
      <p className="dropzone-hint">{fileName ? filledHint : emptyHint}</p>
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) onFile(file);
        }}
      />
    </div>
  );
}
