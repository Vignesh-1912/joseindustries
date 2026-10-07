import { Router } from "express";
import { pool } from "../config/db";
import { requireAuth } from "../middleware/auth";
import { requireModuleAccess, getModuleAccess } from "../utils/permissions";
import { GSTIN_PATTERN, GstinLookupError, fetchGstinDetails, normalizeGstin } from "../services/gstinLookup";
import { asyncHandler } from "../utils/asyncHandler";
import { Vendor } from "../types";

export const vendorsRouter = Router();
const MODULE = "contacts.vendors";

vendorsRouter.use(requireAuth);

async function findVendorById(id: number): Promise<Vendor | undefined> {
  const [rows] = await pool.query<any[]>("SELECT * FROM vendors WHERE id = ? LIMIT 1", [id]);
  return rows[0] as Vendor | undefined;
}

vendorsRouter.get(
  "/",
  requireModuleAccess(MODULE, "view"),
  asyncHandler(async (req, res) => {
    const page = Math.max(Number(req.query.page) || 1, 1);
    const perPage = Math.min(Math.max(Number(req.query.perPage) || 10, 1), 100);
    const search = typeof req.query.search === "string" ? req.query.search.trim() : "";
    const offset = (page - 1) * perPage;

    const searchClause = search ? "WHERE name LIKE ? OR phone LIKE ? OR email LIKE ?" : "";
    const searchParams = search ? [`%${search}%`, `%${search}%`, `%${search}%`] : [];

    const [rows] = await pool.query<any[]>(
      `SELECT * FROM vendors ${searchClause} ORDER BY name ASC LIMIT ? OFFSET ?`,
      [...searchParams, perPage, offset]
    );
    const [countRows] = await pool.query<any[]>(
      `SELECT COUNT(*) as total FROM vendors ${searchClause}`,
      searchParams
    );

    res.json({ data: rows as Vendor[], meta: { page, perPage, total: countRows[0].total as number } });
  })
);

// Find a vendor by GSTIN - same flow as customers: our own database first, the
// paid GSTIN API only for a GSTIN we have never seen, and the result is saved
// straight away. Registered before "/:id".
vendorsRouter.get(
  "/lookup-gstin/:gstin",
  requireModuleAccess(MODULE, "view"),
  asyncHandler(async (req, res) => {
    const gstin = normalizeGstin(String(req.params.gstin));
    if (!GSTIN_PATTERN.test(gstin)) {
      return res.status(400).json({ message: "That is not a valid 15-character GSTIN" });
    }

    const [rows] = await pool.query<any[]>("SELECT * FROM vendors WHERE gstin = ? ORDER BY id ASC LIMIT 1", [gstin]);
    if (rows[0]) return res.json({ source: "database", vendor: rows[0] as Vendor });

    const access = await getModuleAccess(req.user!.sub, req.user!.role, MODULE);
    if (!access.can_create) {
      return res.status(403).json({ message: "This GSTIN is not in your vendors yet, and you do not have permission to add vendors" });
    }

    try {
      const details = await fetchGstinDetails(gstin);
      const [result] = await pool.query<any>(
        `INSERT INTO vendors (name, gstin, address, state) VALUES (?, ?, ?, ?)`,
        [details.name, gstin, details.billing_address, details.state]
      );
      const created = await findVendorById(result.insertId);
      res.status(201).json({ source: "gstin_api", vendor: created, gstStatus: details.gstStatus });
    } catch (err) {
      if (err instanceof GstinLookupError) return res.status(err.status).json({ message: err.message });
      throw err;
    }
  })
);

vendorsRouter.get(
  "/:id",
  requireModuleAccess(MODULE, "view"),
  asyncHandler(async (req, res) => {
    const vendor = await findVendorById(Number(req.params.id));
    if (!vendor) return res.status(404).json({ message: "Vendor not found" });
    res.json({ vendor });
  })
);

vendorsRouter.post(
  "/",
  requireModuleAccess(MODULE, "create"),
  asyncHandler(async (req, res) => {
    const { name, gstin, phone, email, address, state } = req.body ?? {};
    if (!name) return res.status(400).json({ message: "Name is required" });

    const [result] = await pool.query<any>(
      `INSERT INTO vendors (name, gstin, phone, email, address, state) VALUES (?, ?, ?, ?, ?, ?)`,
      [name, gstin || null, phone || null, email || null, address || null, state || null]
    );
    const created = await findVendorById(result.insertId);
    res.status(201).json({ vendor: created });
  })
);

vendorsRouter.put(
  "/:id",
  requireModuleAccess(MODULE, "edit"),
  asyncHandler(async (req, res) => {
    const id = Number(req.params.id);
    const existing = await findVendorById(id);
    if (!existing) return res.status(404).json({ message: "Vendor not found" });

    const { name, gstin, phone, email, address, state } = req.body ?? {};
    if (!name) return res.status(400).json({ message: "Name is required" });

    await pool.query(
      `UPDATE vendors SET name = ?, gstin = ?, phone = ?, email = ?, address = ?, state = ? WHERE id = ?`,
      [name, gstin || null, phone || null, email || null, address || null, state || null, id]
    );
    const updated = await findVendorById(id);
    res.json({ vendor: updated });
  })
);

vendorsRouter.delete(
  "/:id",
  requireModuleAccess(MODULE, "delete"),
  asyncHandler(async (req, res) => {
    const id = Number(req.params.id);
    const existing = await findVendorById(id);
    if (!existing) return res.status(404).json({ message: "Vendor not found" });

    await pool.query("DELETE FROM vendors WHERE id = ?", [id]);
    res.json({ message: "Vendor deleted" });
  })
);
