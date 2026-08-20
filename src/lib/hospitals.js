import { DEFAULT_LEDGER_NAMES } from "./tokens.js";

// Per-hospital config profiles — ledger/voucher names AND the bill
// categorization strategy (see billMappingMode) can differ per hospital.
// Persisted client-side only; there is no backend, so this is a config
// switcher, not authentication (see the "Hospital Profile Switcher" plan).
const STORAGE_KEY = "finova-mapping-tool:hospitalProfiles";
const LEGACY_LEDGER_NAMES_KEY = "finova-mapping-tool:ledgerNames";

// "fixed": today's behavior — three configurable buckets (Others / Pharmacy
//   / Discharge), voucher type + discount ledger come from ledgerNames.
// "per-bill-name": each BILLNAME gets its own bucket; voucher type is the
//   bill name itself, discount ledger is "Discount - {bill name}". See
//   mapBills() in bills.js.
export const BILL_MAPPING_MODES = { FIXED: "fixed", PER_BILL_NAME: "per-bill-name" };

function makeHospital(id, name, shortCode, billMappingMode) {
  return {
    id,
    name,
    shortCode,
    billMappingMode,
    ledgerNames: { ...DEFAULT_LEDGER_NAMES },
    updatedAt: new Date().toISOString(),
  };
}

export const DEFAULT_HOSPITALS = {
  version: 2,
  activeHospitalId: "hospA",
  hospitals: {
    hospA: makeHospital("hospA", "Hospital A", "HOSPA", BILL_MAPPING_MODES.FIXED),
    hospB: makeHospital("hospB", "Hospital B", "HOSPB", BILL_MAPPING_MODES.PER_BILL_NAME),
  },
};

export function loadHospitalProfiles() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (parsed && parsed.hospitals && Object.keys(parsed.hospitals).length > 0) {
        // Top-up: earlier migrations from the legacy single-profile key
        // only created hospA, leaving upgrading users without a Hospital B
        // pill to click. If B still isn't there, seed it — never touch A.
        if (!parsed.hospitals.hospB) {
          parsed.hospitals.hospB = makeHospital(
            "hospB",
            "Hospital B",
            "HOSPB",
            BILL_MAPPING_MODES.PER_BILL_NAME
          );
        }
        return parsed;
      }
    }

    // Migrate the single-profile key from before hospital profiles existed.
    // The legacy key is intentionally left in place, untouched, as a safety net.
    // Hospital B is seeded alongside so the switcher UI always has both
    // pills to pick from — otherwise upgrading users would only see one.
    const legacyRaw = localStorage.getItem(LEGACY_LEDGER_NAMES_KEY);
    if (legacyRaw) {
      const legacyLedgerNames = JSON.parse(legacyRaw);
      return {
        version: 2,
        activeHospitalId: "hospA",
        hospitals: {
          hospA: {
            ...makeHospital("hospA", "Hospital A", "HOSPA", BILL_MAPPING_MODES.FIXED),
            ledgerNames: { ...DEFAULT_LEDGER_NAMES, ...legacyLedgerNames },
          },
          hospB: makeHospital("hospB", "Hospital B", "HOSPB", BILL_MAPPING_MODES.PER_BILL_NAME),
        },
      };
    }

    // Fresh install — pre-seed both hospitals so there's always an active one.
    return DEFAULT_HOSPITALS;
  } catch {
    return DEFAULT_HOSPITALS;
  }
}

export function saveHospitalProfiles(profiles) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(profiles));
  } catch {
    // localStorage unavailable (private browsing, quota, etc.) — settings
    // just won't persist across tab closes.
  }
}

export function getActiveHospital(profiles) {
  return profiles.hospitals[profiles.activeHospitalId] ?? null;
}

export function withUpdatedLedgerNames(profiles, hospitalId, ledgerNames) {
  const hospital = profiles.hospitals[hospitalId];
  if (!hospital) return profiles;
  return {
    ...profiles,
    hospitals: {
      ...profiles.hospitals,
      [hospitalId]: { ...hospital, ledgerNames, updatedAt: new Date().toISOString() },
    },
  };
}

export function withActiveHospital(profiles, hospitalId) {
  if (!profiles.hospitals[hospitalId]) return profiles;
  return { ...profiles, activeHospitalId: hospitalId };
}

