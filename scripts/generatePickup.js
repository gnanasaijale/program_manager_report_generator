// const ExcelJS = require("exceljs");
// const { supabase } = require("../supabaseClient");
// const { uploadReport } = require("../storageUtils");
// const fs = require("fs");

// async function generatePickupReport(programId) {
//   console.log("📦 Generating pickup report for program:", programId);

//   const ordersRes = await supabase
//     .from("orders")
//     .select("id")
//     .eq("program_id", programId);

//   if (ordersRes.error) {
//     console.error("❌ Error fetching orders:", ordersRes.error);
//     throw new Error(ordersRes.error.message);
//   }

//   const orderIds = ordersRes.data.map((o) => o.id);
//   console.log("📦 Order IDs:", orderIds);

//   if (!orderIds.length) {
//     throw new Error("No orders found for program");
//   }

//   const registrationsRes = await supabase
//     .from("registrations")
//     .select(
//       `
//       arrival,
//       service_no_ongoing,
//       group_id,
//       participants(first_name, last_name, primary_phone_number, city),
//       products(name)
//     `
//     )
//     .in("order_id", orderIds);

//   if (registrationsRes.error) {
//     console.error("❌ Error fetching registrations:", registrationsRes.error);
//     throw new Error(registrationsRes.error.message);
//   }

//   const data = registrationsRes.data || [];
//   console.log("📦 Registrations fetched:", data.length);

//   if (!data.length) {
//     throw new Error("No registrations found");
//   }

//   console.log("🔍 Sample registration:", data[0]);
//   console.log("🔧 Starting row transformation...");

//   const rows = data
//     .map((r, idx) => {
//       try {
//         return {
//           arrival_date: r.arrival?.split("T")[0],
//           service_no: r.service_no_ongoing || "",
//           group_id: r.group_id || "",
//           Name: `${r.participants?.first_name || ""} ${
//             r.participants?.last_name || ""
//           }`.trim(),
//           Phone: r.participants?.primary_phone_number || "",
//           City: r.participants?.city || "",
//           Product: r.products?.name || "",
//         };
//       } catch (e) {
//         console.error(
//           `❌ Failed transforming row at index ${idx}:`,
//           r,
//           e.message
//         );
//         return null;
//       }
//     })
//     .filter(Boolean);

//   console.log("✅ Rows created:", rows.length);

//   const workbook = new ExcelJS.Workbook();
//   const sheet = workbook.addWorksheet("Pickup Report");

//   sheet.columns = [
//     { header: "Arrival Date", key: "arrival_date", width: 15 },
//     { header: "Service No", key: "service_no", width: 15 },
//     { header: "Group ID", key: "group_id", width: 15 },
//     { header: "Name", key: "Name", width: 25 },
//     { header: "Phone", key: "Phone", width: 15 },
//     { header: "City", key: "City", width: 15 },
//     { header: "Product", key: "Product", width: 20 },
//   ];
//   sheet.getRow(1).font = { bold: true };

//   // Add and sort rows
//   rows.sort((a, b) => {
//     return (
//       a.arrival_date.localeCompare(b.arrival_date) ||
//       a.group_id.localeCompare(b.group_id) ||
//       a.Name.localeCompare(b.Name)
//     );
//   });

//   sheet.addRows(rows);

//   // Get column indices
//   const arrivalCol = sheet.getColumn("arrival_date").number;
//   const groupCol = sheet.getColumn("group_id").number;

//   // Merge helper
//   function mergeIfSame(startRow, colIdx) {
//     let endRow = startRow;
//     const value = sheet.getCell(startRow, colIdx).value;

//     for (let i = startRow + 1; i <= sheet.rowCount; i++) {
//       if (sheet.getCell(i, colIdx).value !== value) break;
//       endRow = i;
//     }

//     if (endRow > startRow) {
//       sheet.mergeCells(startRow, colIdx, endRow, colIdx);
//       sheet.getCell(startRow, colIdx).alignment = {
//         vertical: "middle",
//         horizontal: "center",
//       };
//     }

//     return endRow + 1;
//   }

//   // Perform merging
//   let row = 2;
//   while (row <= sheet.rowCount) {
//     const startRowForArrival = row;
//     const arrivalValue = sheet.getCell(row, arrivalCol).value;

//     // First, find how many rows have the same arrival_date
//     let endArrivalRow = row;
//     for (let r = row + 1; r <= sheet.rowCount; r++) {
//       if (sheet.getCell(r, arrivalCol).value !== arrivalValue) break;
//       endArrivalRow = r;
//     }

//     // Now, within this block, merge group_id cells
//     let subRow = startRowForArrival;
//     while (subRow <= endArrivalRow) {
//       subRow = mergeIfSame(subRow, groupCol);
//     }

//     // Merge arrival_date cells
//     if (endArrivalRow > startRowForArrival) {
//       sheet.mergeCells(
//         startRowForArrival,
//         arrivalCol,
//         endArrivalRow,
//         arrivalCol
//       );
//       sheet.getCell(startRowForArrival, arrivalCol).alignment = {
//         vertical: "middle",
//         horizontal: "center",
//       };
//     }

//     row = endArrivalRow + 1;
//   }

//   const fileName = `pickup-${Date.now()}.xlsx`;
//   await workbook.xlsx.writeFile(fileName);
//   console.log("✅ Excel file written:", fileName);

//   const uploadedPath = await uploadReport(fileName, "pickup", "xlsx");
//   console.log("📤 Uploaded to Supabase:", uploadedPath);

//   fs.unlinkSync(fileName); // Clean up local file
//   return uploadedPath;
// }

// module.exports = { generatePickupReport };

const ExcelJS = require("exceljs");
const { supabase } = require("../supabaseClient");
const { uploadReport } = require("../storageUtils");
const fs = require("fs");

async function generatePickupReport(programId) {
  console.log("📦 Generating Pickup report for program:", programId);

  const ordersRes = await supabase
    .from("orders")
    .select("id")
    .eq("program_id", programId);

  if (ordersRes.error) {
    console.error("❌ Error fetching orders:", ordersRes.error);
    throw new Error(ordersRes.error.message);
  }

  const orderIds = ordersRes.data.map((o) => o.id);
  console.log("📦 Order IDs:", orderIds);

  if (!orderIds.length) {
    throw new Error("No orders found for program");
  }

  const registrationsRes = await supabase
    .from("registrations")
    .select(
      `
      arrival,
      service_no_ongoing,
      group_id,
      participants(first_name, last_name, primary_phone_number, city),
      products(name)
    `
    )
    .in("order_id", orderIds);

  if (registrationsRes.error) {
    console.error("❌ Error fetching registrations:", registrationsRes.error);
    throw new Error(registrationsRes.error.message);
  }

  const data = registrationsRes.data || [];
  console.log("📦 Registrations fetched:", data.length);

  if (!data.length) {
    throw new Error("No registrations found");
  }

  const rows = data.map((r) => ({
    arrival_date: r.arrival?.split("T")[0],
    service_no: r.service_no_ongoing || "",
    group_id: r.group_id || "",
    Name: `${r.participants?.first_name || ""} ${
      r.participants?.last_name || ""
    }`.trim(),
    Phone: r.participants?.primary_phone_number || "",
    City: r.participants?.city || "",
    Product: r.products?.name || "",
  }));

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Pickup Report");

  sheet.columns = [
    { header: "Arrival Date", key: "arrival_date", width: 15 },
    { header: "Service No", key: "service_no", width: 15 },
    { header: "Group ID", key: "group_id", width: 15 },
    { header: "Name", key: "Name", width: 25 },
    { header: "Phone", key: "Phone", width: 15 },
    { header: "City", key: "City", width: 15 },
    { header: "Product", key: "Product", width: 20 },
  ];

  sheet.addRows(rows);

  const fileName = `pickup-${Date.now()}.xlsx`;
  await workbook.xlsx.writeFile(fileName);
  console.log("✅ Excel file written:", fileName);

  const uploadedPath = await uploadReport(fileName, "Pickup", "xlsx");
  console.log("📤 Uploaded to Supabase:", uploadedPath);

  fs.unlinkSync(fileName); // Clean up local file
  return uploadedPath;
}

module.exports = { generatePickupReport };
