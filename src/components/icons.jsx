import React from "react";

const base = {
  viewBox: "0 0 24 24",
  fill: "none",
  stroke: "currentColor",
  strokeWidth: 1.8,
  strokeLinecap: "round",
  strokeLinejoin: "round",
};

export function HomeIcon() {
  return (
    <svg {...base}>
      <path d="M4 11.5 12 4l8 7.5" />
      <path d="M6 10v9h12v-9" />
      <path d="M10 19v-5h4v5" />
    </svg>
  );
}

export function ReceiptsIcon() {
  return (
    <svg {...base}>
      <path d="M6 3h12v18l-2.5-1.6L13 21l-1-1.6-1 1.6-2.5-1.6L6 21z" />
      <path d="M9 8h6M9 12h6" />
    </svg>
  );
}

export function BillsIcon() {
  return (
    <svg {...base}>
      <path d="M14 3H7a1 1 0 0 0-1 1v16a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V8z" />
      <path d="M14 3v5h4" />
      <path d="M9 13h6M9 17h6" />
    </svg>
  );
}

export function IpCreditIcon() {
  return (
    <svg {...base}>
      <rect x="3" y="6" width="18" height="13" rx="2" />
      <path d="M3 10.5h18" />
      <path d="M7 14.5h4" />
    </svg>
  );
}

export function DialysisIcon() {
  return (
    <svg {...base}>
      <path d="M12 3s6 6.5 6 11a6 6 0 0 1-12 0c0-4.5 6-11 6-11z" />
    </svg>
  );
}

export function ReconcileIcon() {
  return (
    <svg {...base}>
      <path d="M17 3 20.5 6.5 17 10" />
      <path d="M3.5 6.5H20.5" />
      <path d="M7 21l-3.5-3.5L7 14" />
      <path d="M20.5 17.5H3.5" />
    </svg>
  );
}

export function MasterIcon() {
  return (
    <svg {...base}>
      <path d="M4 6h16M4 12h11M4 18h7" />
    </svg>
  );
}

export function WorkingsIcon() {
  return (
    <svg {...base}>
      <path d="M4 7h10M17 7h3" />
      <path d="M4 12h3M10 12h10" />
      <path d="M4 17h10M17 17h3" />
      <circle cx="14" cy="7" r="2" />
      <circle cx="7" cy="12" r="2" />
      <circle cx="14" cy="17" r="2" />
    </svg>
  );
}
