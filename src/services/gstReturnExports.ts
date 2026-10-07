import ExcelJS from "exceljs";
import PDFDocument from "pdfkit";
import { Response } from "express";
import { Gstr1Result, Gstr3bResult } from "./gstReturns";

interface ExportCompany {
  name: string;
  code: string;
}

type ExportRow = Array<string | number>;

function baseWorkbook(title: string, company: ExportCompany, from: string, to: string): ExcelJS.Workbook {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Jose Industries";
  workbook.subject = `${title} preparation report`;
  workbook.created = new Date();
  workbook.properties.date1904 = false;
  workbook.company = company.name;
  workbook.title = title;
  workbook.description = `${company.name} (${company.code}), ${from} to ${to}`;
  return workbook;
}

function applyNumberFormats(sheet: ExcelJS.Worksheet, headers: string[]): void {
  headers.forEach((header, index) => {
    if (/taxable|cgst|sgst|igst|tax amount|invoice total|amount|payable|refundable/i.test(header)) {
      sheet.getColumn(index + 1).numFmt = "#,##0.00;[Red]-#,##0.00";
    } else if (/quantity|^qty$/i.test(header)) {
      sheet.getColumn(index + 1).numFmt = "#,##0.####";
    } else if (/rate %/i.test(header)) {
      sheet.getColumn(index + 1).numFmt = "0.##";
    }
  });
}

function addSheet(
  workbook: ExcelJS.Workbook,
  name: string,
  title: string,
  company: ExportCompany,
  from: string,
  to: string,
  headers: string[],
  rows: ExportRow[],
  widths?: number[]
): ExcelJS.Worksheet {
  const sheet = workbook.addWorksheet(name);
  sheet.addRow([title]);
  sheet.addRow([company.name, `(${company.code})`]);
  sheet.addRow([`Period: ${from} to ${to}`]);
  sheet.addRow([]);
  const header = sheet.addRow(headers);
  header.font = { bold: true, color: { argb: "FFFFFFFF" } };
  header.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1B7A4D" } };
  header.alignment = { vertical: "middle", wrapText: true };
  sheet.views = [{ state: "frozen", ySplit: 5 }];
  for (const row of rows) sheet.addRow(row);
  if (widths) widths.forEach((width, index) => { sheet.getColumn(index + 1).width = width; });
  else sheet.columns.forEach((column) => { column.width = 20; });
  applyNumberFormats(sheet, headers);
  sheet.getRow(1).font = { bold: true, size: 16 };
  sheet.getRow(2).font = { bold: true, size: 12 };
  sheet.autoFilter = { from: { row: 5, column: 1 }, to: { row: 5, column: headers.length } };
  return sheet;
}

function invoiceRows(rows: Gstr1Result["b2b"]["invoices"]): ExportRow[] {
  return rows.map((row) => [
    row.docNumber,
    row.issueDate,
    row.status,
    row.customerName,
    row.customerGstin || "",
    row.taxableValue,
    row.cgst,
    row.sgst,
    row.igst,
    row.taxTotal,
    row.grandTotal,
  ]);
}

export function createGstr1Workbook(
  data: Gstr1Result,
  company: ExportCompany
): ExcelJS.Workbook {
  const workbook = baseWorkbook("GSTR-1", company, data.from, data.to);
  const invoiceHeaders = [
    "Invoice No.", "Date", "Status", "Customer", "GSTIN", "Taxable Value",
    "CGST", "SGST", "IGST", "Tax Amount", "Invoice Total",
  ];
  const invoices = addSheet(
    workbook, "B2B Invoices", "GSTR-1 - B2B Invoices", company, data.from, data.to,
    invoiceHeaders, invoiceRows(data.b2b.invoices),
    [22, 14, 14, 28, 20, 18, 16, 16, 16, 16, 18]
  );
  if (data.b2b.invoices.length) {
    const total = invoices.addRow([
      "Total", "", "", "", "",
      data.b2b.totalTaxableValue, data.b2b.totalCgst, data.b2b.totalSgst, data.b2b.totalIgst,
      data.b2b.totalCgst + data.b2b.totalSgst + data.b2b.totalIgst,
      data.b2b.invoices.reduce((sum, row) => sum + row.grandTotal, 0),
    ]);
    total.font = { bold: true };
  }

  addSheet(
    workbook, "B2C Summary", "GSTR-1 - B2C (Others)", company, data.from, data.to,
    ["Invoice Count", "Taxable Value", "CGST", "SGST", "IGST"],
    [[data.b2c.invoiceCount, data.b2c.totalTaxableValue, data.b2c.totalCgst, data.b2c.totalSgst, data.b2c.totalIgst]],
    [18, 20, 18, 18, 18]
  );
  addSheet(
    workbook, "HSN Summary", "GSTR-1 - HSN Summary", company, data.from, data.to,
    ["HSN/SAC", "Tax Rate %", "Quantity", "Taxable Value", "CGST", "SGST", "IGST", "Tax Amount"],
    data.hsnSummary.map((row) => [
      row.hsnCode, row.taxRate, row.qty, row.taxableValue, row.cgst, row.sgst, row.igst, row.taxTotal,
    ]),
    [20, 16, 16, 20, 18, 18, 18, 18]
  );
  addSheet(
    workbook, "Cancelled Invoices", "GSTR-1 - Cancelled Invoices (Excluded from totals)", company, data.from, data.to,
    invoiceHeaders, invoiceRows(data.cancelledInvoices),
    [22, 14, 14, 28, 20, 18, 16, 16, 16, 16, 18]
  );
  addSheet(
    workbook, "Notes", "GSTR-1 - Preparation Notes", company, data.from, data.to,
    ["Note"], [
      [`Draft invoices included in totals: ${data.draftCount}`],
      ...data.unsupported.map((note) => [note]),
    ],
    [120]
  );
  return workbook;
}

export function createGstr3bWorkbook(
  data: Gstr3bResult,
  company: ExportCompany
): ExcelJS.Workbook {
  const workbook = baseWorkbook("GSTR-3B", company, data.from, data.to);
  const headers = ["Section", "Taxable Value", "IGST", "CGST", "SGST", "Status / Note"];
  const toRows = (sections: Gstr3bResult["outwardSupplies"]) =>
    sections.map((row) => [
      row.label,
      row.taxableValue ?? "",
      row.igst ?? "",
      row.cgst ?? "",
      row.sgst ?? "",
      [row.notTracked ? "Not tracked" : "", row.note || ""].filter(Boolean).join(" - "),
    ]);

  addSheet(
    workbook, "Outward Supplies", "GSTR-3B - Table 3.1 Outward Supplies", company, data.from, data.to,
    headers, toRows(data.outwardSupplies), [58, 20, 18, 18, 18, 76]
  );
  addSheet(
    workbook, "Eligible ITC", "GSTR-3B - Table 4 Eligible ITC", company, data.from, data.to,
    headers, toRows(data.itc), [58, 20, 18, 18, 18, 76]
  );
  addSheet(
    workbook, "Summary", "GSTR-3B - Net GST Position", company, data.from, data.to,
    ["Metric", "Amount"], [
      ["Net GST Payable", data.netGstPayable],
      ["Net GST Refundable", data.netGstRefundable],
    ],
    [32, 22]
  );
  addSheet(
    workbook, "Notes", "GSTR-3B - Preparation Notes", company, data.from, data.to,
    ["Note"], data.unsupported.map((note) => [note]), [120]
  );
  return workbook;
}

interface PdfSection {
  title: string;
  headers: string[];
  widths: number[];
  rows: ExportRow[];
}

function streamReturnPdf(
  res: Response,
  title: string,
  company: ExportCompany,
  from: string,
  to: string,
  sections: PdfSection[]
): void {
  const pageWidth = 841.89;
  const pageHeight = 595.28;
  const margin = 32;
  const bottom = pageHeight - margin;
  const doc = new PDFDocument({ size: "A4", layout: "landscape", margin, bufferPages: true });
  res.setHeader("Content-Type", "application/pdf");
  const safeCompanyCode = company.code.replace(/[^a-zA-Z0-9_-]+/g, "-");
  const safeTitle = title.replace(/[^a-zA-Z0-9_-]+/g, "-");
  res.setHeader(
    "Content-Disposition",
    `inline; filename="${safeTitle}-${safeCompanyCode}-${from}-to-${to}.pdf"`
  );
  doc.pipe(res);

  const drawPageHeading = () => {
    doc.font("Helvetica-Bold").fontSize(16).fillColor("#181818").text(title, margin, margin);
    doc.font("Helvetica-Bold").fontSize(10).text(company.name, margin, margin + 23);
    doc.font("Helvetica").fontSize(9).text(`Company code: ${company.code}   |   Period: ${from} to ${to}`, margin, margin + 38);
    doc.moveTo(margin, margin + 54).lineTo(pageWidth - margin, margin + 54).strokeColor("#1B7A4D").lineWidth(1).stroke();
    doc.y = margin + 66;
  };
  const newPage = () => {
    doc.addPage();
    drawPageHeading();
  };
  drawPageHeading();

  for (const section of sections) {
    if (doc.y + 38 > bottom) newPage();
    doc.font("Helvetica-Bold").fontSize(11).fillColor("#181818").text(section.title, margin, doc.y);
    doc.moveDown(0.5);

    const drawHeader = () => {
      const y = doc.y;
      doc.save().fillColor("#1B7A4D").rect(margin, y, section.widths.reduce((sum, width) => sum + width, 0), 20).fill().restore();
      doc.font("Helvetica-Bold").fontSize(7).fillColor("#FFFFFF");
      let x = margin;
      section.headers.forEach((header, index) => {
        doc.text(header, x + 3, y + 6, { width: section.widths[index] - 6, height: 12, ellipsis: true });
        x += section.widths[index];
      });
      doc.y = y + 20;
    };

    const drawRow = (row: ExportRow) => {
      const fontSize = 7;
      doc.font("Helvetica").fontSize(fontSize);
      const heights = row.map((value, index) => doc.heightOfString(String(value), {
        width: section.widths[index] - 6,
        lineBreak: true,
      }));
      const height = Math.max(18, Math.min(48, Math.max(...heights) + 8));
      if (doc.y + height > bottom) {
        newPage();
        doc.font("Helvetica-Bold").fontSize(8).fillColor("#181818").text(section.title, margin, doc.y);
        doc.moveDown(0.5);
        drawHeader();
      }
      const y = doc.y;
      doc.save().strokeColor("#D8DED9").lineWidth(0.5)
        .rect(margin, y, section.widths.reduce((sum, width) => sum + width, 0), height).stroke().restore();
      let x = margin;
      row.forEach((value, index) => {
        doc.font("Helvetica").fontSize(fontSize).fillColor("#181818")
          .text(String(value), x + 3, y + 4, {
            width: section.widths[index] - 6,
            height: height - 6,
            ellipsis: true,
          });
        x += section.widths[index];
      });
      doc.y = y + height;
    };

    drawHeader();
    if (section.rows.length) section.rows.forEach(drawRow);
    else {
      doc.font("Helvetica-Oblique").fontSize(8).fillColor("#666666").text("No data for this period.", margin + 4, doc.y + 5);
      doc.y += 22;
    }
    doc.moveDown(1);
  }

  const pageRange = doc.bufferedPageRange();
  for (let index = 0; index < pageRange.count; index++) {
    doc.switchToPage(pageRange.start + index);
    doc.font("Helvetica").fontSize(7).fillColor("#666666")
      .text("Internal GST return preparation report - verify figures before filing.", margin, pageHeight - 20, {
        width: pageWidth - margin * 2,
        align: "center",
      });
  }
  doc.end();
}

export function streamGstr1Pdf(res: Response, data: Gstr1Result, company: ExportCompany): void {
  const invoiceHeaders = ["Invoice No.", "Date", "Status", "Customer", "GSTIN", "Taxable", "CGST", "SGST", "IGST", "Tax", "Total"];
  const invoiceWidths = [76, 51, 44, 112, 80, 67, 55, 55, 55, 60, 63];
  const invoiceRowsForPdf = (rows: Gstr1Result["b2b"]["invoices"]): ExportRow[] =>
    rows.map((row) => [
      row.docNumber, row.issueDate, row.status, row.customerName, row.customerGstin || "-",
      row.taxableValue.toFixed(2), row.cgst.toFixed(2), row.sgst.toFixed(2), row.igst.toFixed(2),
      row.taxTotal.toFixed(2), row.grandTotal.toFixed(2),
    ]);
  const sections: PdfSection[] = [
    {
      title: "B2B Invoices",
      headers: invoiceHeaders,
      widths: invoiceWidths,
      rows: invoiceRowsForPdf(data.b2b.invoices),
    },
    {
      title: "B2B Totals",
      headers: ["Taxable Value", "CGST", "SGST", "IGST"],
      widths: [194, 194, 194, 194],
      rows: [[data.b2b.totalTaxableValue.toFixed(2), data.b2b.totalCgst.toFixed(2), data.b2b.totalSgst.toFixed(2), data.b2b.totalIgst.toFixed(2)]],
    },
    {
      title: "B2C (Others) Summary",
      headers: ["Invoice Count", "Taxable Value", "CGST", "SGST", "IGST"],
      widths: [155, 155, 155, 155, 156],
      rows: [[data.b2c.invoiceCount, data.b2c.totalTaxableValue.toFixed(2), data.b2c.totalCgst.toFixed(2), data.b2c.totalSgst.toFixed(2), data.b2c.totalIgst.toFixed(2)]],
    },
    {
      title: "HSN Summary",
      headers: ["HSN/SAC", "Rate %", "Quantity", "Taxable Value", "CGST", "SGST", "IGST", "Tax Amount"],
      widths: [108, 82, 82, 112, 92, 92, 92, 92],
      rows: data.hsnSummary.map((row) => [
        row.hsnCode, row.taxRate, row.qty, row.taxableValue.toFixed(2),
        row.cgst.toFixed(2), row.sgst.toFixed(2), row.igst.toFixed(2), row.taxTotal.toFixed(2),
      ]),
    },
    {
      title: "Cancelled Invoices (excluded from totals)",
      headers: invoiceHeaders,
      widths: invoiceWidths,
      rows: invoiceRowsForPdf(data.cancelledInvoices),
    },
    {
      title: "Preparation Notes",
      headers: ["Note"],
      widths: [777],
      rows: [[`Draft invoices included in totals: ${data.draftCount}`], ...data.unsupported.map((note) => [note])],
    },
  ];
  streamReturnPdf(res, "GSTR-1 Preparation Report", company, data.from, data.to, sections);
}

export function streamGstr3bPdf(res: Response, data: Gstr3bResult, company: ExportCompany): void {
  const sections = (rows: Gstr3bResult["outwardSupplies"]): ExportRow[] =>
    rows.map((row) => [
      row.label,
      row.taxableValue === null ? "-" : row.taxableValue.toFixed(2),
      row.igst === null ? "-" : row.igst.toFixed(2),
      row.cgst === null ? "-" : row.cgst.toFixed(2),
      row.sgst === null ? "-" : row.sgst.toFixed(2),
      [row.notTracked ? "Not tracked" : "", row.note || ""].filter(Boolean).join(" - "),
    ]);
  streamReturnPdf(res, "GSTR-3B Preparation Report", company, data.from, data.to, [
    {
      title: "Table 3.1 - Outward Supplies",
      headers: ["Section", "Taxable Value", "IGST", "CGST", "SGST", "Status / Note"],
      widths: [140, 95, 70, 70, 70, 332],
      rows: sections(data.outwardSupplies),
    },
    {
      title: "Table 4 - Eligible ITC",
      headers: ["Section", "Taxable Value", "IGST", "CGST", "SGST", "Status / Note"],
      widths: [140, 95, 70, 70, 70, 332],
      rows: sections(data.itc),
    },
    {
      title: "Net GST Position",
      headers: ["Net GST Payable", "Net GST Refundable"],
      widths: [388.5, 388.5],
      rows: [[data.netGstPayable.toFixed(2), data.netGstRefundable.toFixed(2)]],
    },
    {
      title: "Preparation Notes",
      headers: ["Note"],
      widths: [777],
      rows: data.unsupported.map((note) => [note]),
    },
  ]);
}
