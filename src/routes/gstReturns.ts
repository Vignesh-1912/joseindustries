import { Router } from "express";
import { pool } from "../config/db";
import { requireAuth } from "../middleware/auth";
import { requireModuleAccess } from "../utils/permissions";
import { asyncHandler } from "../utils/asyncHandler";
import { getGstr1, getGstr3b } from "../services/gstReturns";
import {
  createGstr1Workbook,
  createGstr3bWorkbook,
  streamGstr1Pdf,
  streamGstr3bPdf,
} from "../services/gstReturnExports";

// GST Returns preparation (Phase 11B) - GSTR-1/GSTR-3B preparation data
// only. This is an internal preparation/reporting tool, never a filing or
// submission mechanism - there is no GST portal API integration anywhere
// in this codebase. Gated by the existing "reports.reports" permission,
// same as every report since Phase 8 - no new permission introduced.
export const gstReturnsRouter = Router();
gstReturnsRouter.use(requireAuth, requireModuleAccess("reports.reports", "view"));

function dateRange(query: any): { from: string; to: string } {
  const to = typeof query.to === "string" && query.to ? query.to : new Date().toISOString().slice(0, 10);
  const from = typeof query.from === "string" && query.from ? query.from : "1900-01-01";
  return { from, to };
}

async function exportContext(req: import("express").Request, res: import("express").Response) {
  const companyId = Number(req.query.company_id);
  if (!Number.isInteger(companyId) || companyId <= 0) {
    res.status(400).json({ message: "A valid company_id is required" });
    return null;
  }
  const { from, to } = dateRange(req.query);
  const isValidDate = (value: string) =>
    /^\d{4}-\d{2}-\d{2}$/.test(value) &&
    !Number.isNaN(Date.parse(value)) &&
    new Date(`${value}T00:00:00.000Z`).toISOString().slice(0, 10) === value;
  if (!isValidDate(from) || !isValidDate(to) || from > to) {
    res.status(400).json({ message: "Provide a valid date range with the start date on or before the end date" });
    return null;
  }
  const [companyRows] = await pool.query<any[]>(
    "SELECT name, code FROM companies WHERE id = ? LIMIT 1",
    [companyId]
  );
  if (!companyRows[0]) {
    res.status(404).json({ message: "Company not found" });
    return null;
  }
  return { companyId, from, to, company: companyRows[0] as { name: string; code: string } };
}

function filePart(value: string): string {
  return value.replace(/[^a-zA-Z0-9_-]+/g, "-").replace(/^-+|-+$/g, "") || "company";
}

// GSTR-1 - source-document-derived (documents + document_items), never the
// ledger - see getGstr1's own doc comment for the full reasoning.
gstReturnsRouter.get(
  "/gstr1/excel",
  asyncHandler(async (req, res) => {
    const context = await exportContext(req, res);
    if (!context) return;
    const data = await getGstr1(context.companyId, context.from, context.to);
    const workbook = createGstr1Workbook(data, context.company);
    const content = await workbook.xlsx.writeBuffer();
    const filename = `GSTR-1-${filePart(context.company.code)}-${context.from}-to-${context.to}.xlsx`;
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.end(content);
  })
);

gstReturnsRouter.get(
  "/gstr1/pdf",
  asyncHandler(async (req, res) => {
    const context = await exportContext(req, res);
    if (!context) return;
    const data = await getGstr1(context.companyId, context.from, context.to);
    streamGstr1Pdf(res, data, context.company);
  })
);

gstReturnsRouter.get(
  "/gstr1",
  asyncHandler(async (req, res) => {
    const companyId = req.query.company_id ? Number(req.query.company_id) : null;
    if (!companyId) return res.status(400).json({ message: "company_id is required" });
    const { from, to } = dateRange(req.query);

    const result = await getGstr1(companyId, from, to);
    res.json(result);
  })
);

// GSTR-3B - a reshaping of getGstSummary/getProfitAndLoss's existing,
// unmodified ledger figures - see getGstr3b's own doc comment.
gstReturnsRouter.get(
  "/gstr3b/excel",
  asyncHandler(async (req, res) => {
    const context = await exportContext(req, res);
    if (!context) return;
    const data = await getGstr3b(context.companyId, context.from, context.to);
    const workbook = createGstr3bWorkbook(data, context.company);
    const content = await workbook.xlsx.writeBuffer();
    const filename = `GSTR-3B-${filePart(context.company.code)}-${context.from}-to-${context.to}.xlsx`;
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.end(content);
  })
);

gstReturnsRouter.get(
  "/gstr3b/pdf",
  asyncHandler(async (req, res) => {
    const context = await exportContext(req, res);
    if (!context) return;
    const data = await getGstr3b(context.companyId, context.from, context.to);
    streamGstr3bPdf(res, data, context.company);
  })
);

gstReturnsRouter.get(
  "/gstr3b",
  asyncHandler(async (req, res) => {
    const companyId = req.query.company_id ? Number(req.query.company_id) : null;
    if (!companyId) return res.status(400).json({ message: "company_id is required" });
    const { from, to } = dateRange(req.query);

    const result = await getGstr3b(companyId, from, to);
    res.json(result);
  })
);
