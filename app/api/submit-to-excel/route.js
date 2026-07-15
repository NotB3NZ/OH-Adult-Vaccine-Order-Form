import { NextResponse } from "next/server";

/**
 * POST /api/submit-to-excel
 *
 * Receives a flat array of row objects from the frontend and
 * demonstrates how to forward them to Excel Online.
 *
 * Expected body: { rows: [ { OrderID, Timestamp, AssociateNumber, ... }, ... ] }
 *
 * In production, replace the mock logic below with either:
 *   A) Microsoft Graph API calls to append rows to an Excel workbook
 *   B) A Power Automate webhook trigger
 */

// --- Column order matching the Excel spreadsheet ---
const EXCEL_COLUMNS = [
  "OrderID",
  "Timestamp",
  "AssociateNumber",
  "OrderType",
  "DPT_Qty", "DPT_Site",
  "HepA_Qty", "HepA_Site",
  "HepB_Qty", "HepB_Site",
  "HepAB_Qty", "HepAB_Site",
  "HPV_Qty", "HPV_Site",
  "MMR_Qty", "MMR_Site",
  "Pneumo13_Qty", "Pneumo13_Site",
  "Pneumo20_Qty", "Pneumo20_Site",
  "Pneumo23_Qty", "Pneumo23_Site",
  "Shingles_Qty", "Shingles_Site",
  "Varicella_Qty", "Varicella_Site",
  "RabiesSpeeda_Qty", "RabiesSpeeda_Site",
  "RabiesVerorab_Qty", "RabiesVerorab_Site",
  "TotalCost",
  "ConsentAgreed",
];

export async function POST(request) {
  try {
    const body = await request.json();
    const { rows } = body;

    // --- Validation ---
    if (!rows || !Array.isArray(rows) || rows.length === 0) {
      return NextResponse.json(
        { error: "No order rows provided." },
        { status: 400 }
      );
    }

    if (rows.length > 2) {
      return NextResponse.json(
        { error: "Maximum 2 rows allowed (Associate + Dependent)." },
        { status: 400 }
      );
    }

    // Validate each row has the required columns
    for (const row of rows) {
      const missingCols = EXCEL_COLUMNS.filter((col) => !(col in row));
      if (missingCols.length > 0) {
        return NextResponse.json(
          { error: `Missing columns: ${missingCols.join(", ")}` },
          { status: 400 }
        );
      }

      if (!row.AssociateNumber) {
        return NextResponse.json(
          { error: "AssociateNumber is required." },
          { status: 400 }
        );
      }

      if (!["Associate", "Dependent"].includes(row.OrderType)) {
        return NextResponse.json(
          { error: `Invalid OrderType: ${row.OrderType}` },
          { status: 400 }
        );
      }
    }

    // --- Transform rows into 2D arrays for Excel ---
    // Microsoft Graph API expects values as a 2D array: [[cell1, cell2, ...], ...]
    const excelValues = rows.map((row) =>
      EXCEL_COLUMNS.map((col) => row[col] ?? "")
    );

    // --- Log for development ---
    console.log("=== VACCINE ORDER RECEIVED ===");
    console.log(`Order ID: ${rows[0].OrderID}`);
    console.log(`Associate: ${rows[0].AssociateNumber}`);
    console.log(`Rows to write: ${rows.length}`);
    console.log("Flat row data:", JSON.stringify(rows, null, 2));
    console.log("Excel-ready 2D values:", JSON.stringify(excelValues, null, 2));
    console.log("==============================");

    // ===============================================================
    // OPTION A: Microsoft Graph API
    // ===============================================================
    // Uncomment and configure the following to write directly to Excel
    // via Microsoft Graph API. You'll need:
    //   1. An Azure AD app registration with Files.ReadWrite scope
    //   2. The workbook's driveItem ID or path in OneDrive/SharePoint
    //   3. An access token (use MSAL or client credentials flow)
    //
    // const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
    // const DRIVE_ITEM_PATH = "/me/drive/items/{item-id}"; // or /sites/{site-id}/drive/items/{item-id}
    // const TABLE_NAME = "VaccineOrders"; // Name of the Excel Table
    // const ACCESS_TOKEN = process.env.MICROSOFT_GRAPH_TOKEN;
    //
    // for (const rowValues of excelValues) {
    //   const response = await fetch(
    //     `${GRAPH_BASE}${DRIVE_ITEM_PATH}/workbook/tables/${TABLE_NAME}/rows/add`,
    //     {
    //       method: "POST",
    //       headers: {
    //         Authorization: `Bearer ${ACCESS_TOKEN}`,
    //         "Content-Type": "application/json",
    //       },
    //       body: JSON.stringify({
    //         values: [rowValues],
    //       }),
    //     }
    //   );
    //
    //   if (!response.ok) {
    //     const errData = await response.json();
    //     console.error("Graph API error:", errData);
    //     throw new Error(`Graph API error: ${response.status}`);
    //   }
    // }

    // ===============================================================
    // OPTION B: Power Automate Webhook
    // ===============================================================
    // If you prefer Power Automate, create an "When an HTTP request
    // is received" trigger flow, then add an "Add a row into a table"
    // Excel action for each row.
    //
    // const WEBHOOK_URL = process.env.POWER_AUTOMATE_WEBHOOK_URL;
    //
    // const response = await fetch(WEBHOOK_URL, {
    //   method: "POST",
    //   headers: { "Content-Type": "application/json" },
    //   body: JSON.stringify({ rows }),
    // });
    //
    // if (!response.ok) {
    //   throw new Error(`Power Automate error: ${response.status}`);
    // }

    // ===============================================================
    // MOCK RESPONSE (remove in production)
    // ===============================================================
    // Simulate a short delay to mimic network request
    await new Promise((resolve) => setTimeout(resolve, 800));

    return NextResponse.json({
      success: true,
      message: `Successfully queued ${rows.length} row(s) for Excel.`,
      orderId: rows[0].OrderID,
      rowsWritten: rows.length,
    });
  } catch (error) {
    console.error("Submit to Excel error:", error);
    return NextResponse.json(
      { error: "Internal server error. Please try again." },
      { status: 500 }
    );
  }
}
