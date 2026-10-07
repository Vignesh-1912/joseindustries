import { useState } from "react";
import { Input, message } from "antd";
import { api } from "../api/client";
import { Customer, Vendor } from "../types";

export interface GstinLookupResult {
  customer: Customer;
  /** "database" = already one of our customers; "gstin_api" = just fetched from the GST portal and saved as a new customer. */
  source: "database" | "gstin_api";
  /** GST portal registration status (only when freshly fetched), e.g. "Active" / "Cancelled". */
  gstStatus?: string | null;
}

export interface VendorGstinLookupResult {
  vendor: Vendor;
  source: "database" | "gstin_api";
  gstStatus?: string | null;
}

interface GstinLookupProps<R> {
  onResult: (result: R) => void;
  placeholder?: string;
  autoFocus?: boolean;
}

const GSTIN_PATTERN = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/;

interface EntityConfig<R> {
  label: "customer" | "vendor";
  path: string;
  nameOf: (r: R) => string;
}

/**
 * "Find by GSTIN": a record we already have is returned straight from our own
 * database; a GSTIN we have never seen is fetched from the GST portal (via the
 * server) and saved as a new record, so the next lookup is a plain database
 * hit. Either way the caller gets the record back.
 */
function GstinLookupBase<R extends { source: "database" | "gstin_api"; gstStatus?: string | null }>({
  onResult,
  placeholder,
  autoFocus,
  config,
}: GstinLookupProps<R> & { config: EntityConfig<R> }) {
  const [value, setValue] = useState("");
  const [loading, setLoading] = useState(false);

  async function search(raw: string) {
    const gstin = raw.replace(/\s+/g, "").toUpperCase();
    if (!gstin) return;
    if (!GSTIN_PATTERN.test(gstin)) {
      message.warning("That is not a valid 15-character GSTIN");
      return;
    }
    setLoading(true);
    try {
      const res = await api.get<R>(`${config.path}/${gstin}`);
      if (res.source === "database") {
        message.success(`Found existing ${config.label}: ${config.nameOf(res)}`);
      } else {
        message.success(`Fetched from GST portal and added: ${config.nameOf(res)}`);
        if (res.gstStatus && res.gstStatus.toLowerCase() !== "active") {
          message.warning(`Note: this GSTIN is ${res.gstStatus} on the GST portal`);
        }
      }
      setValue("");
      onResult(res);
    } catch (err) {
      message.error(err instanceof Error ? err.message : "GSTIN lookup failed");
    } finally {
      setLoading(false);
    }
  }

  return (
    <Input.Search
      value={value}
      onChange={(e) => setValue(e.target.value.toUpperCase())}
      onSearch={search}
      enterButton="Find"
      loading={loading}
      maxLength={15}
      placeholder={placeholder ?? `Enter GSTIN to find or add a ${config.label}`}
      autoFocus={autoFocus}
      allowClear
    />
  );
}

const CUSTOMER_CONFIG: EntityConfig<GstinLookupResult> = {
  label: "customer",
  path: "/customers/lookup-gstin",
  nameOf: (r) => r.customer.name,
};

const VENDOR_CONFIG: EntityConfig<VendorGstinLookupResult> = {
  label: "vendor",
  path: "/vendors/lookup-gstin",
  nameOf: (r) => r.vendor.name,
};

export function GstinLookup(props: GstinLookupProps<GstinLookupResult>) {
  return <GstinLookupBase {...props} config={CUSTOMER_CONFIG} />;
}

export function VendorGstinLookup(props: GstinLookupProps<VendorGstinLookupResult>) {
  return <GstinLookupBase {...props} config={VENDOR_CONFIG} />;
}
