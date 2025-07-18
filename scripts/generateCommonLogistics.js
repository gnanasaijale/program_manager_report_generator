const ExcelJS = require("exceljs");
const { supabase } = require("../supabaseClient");
const { uploadReport } = require("../storageUtils");
const fs = require("fs");

async function generateCommonLogisticsReport(programId) {
  console.log("📦 Generating common logistics report for program:", programId);

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
    group_id: r.group_id || "",
    Name: `${r.participants?.first_name || ""} ${
      r.participants?.last_name || ""
    }`.trim(),
    Phone: r.participants?.primary_phone_number || "",
    City: r.participants?.city || "",
    Product: r.products?.name || "",
  }));

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Common Logistics Report");

  sheet.columns = [
    { header: "Group ID", key: "group_id", width: 15 },
    { header: "Name", key: "Name", width: 25 },
    { header: "Phone", key: "Phone", width: 15 },
    { header: "City", key: "City", width: 15 },
    { header: "Product", key: "Product", width: 20 },
  ];

  sheet.addRows(rows);

  const fileName = `common-logistics-${Date.now()}.xlsx`;
  await workbook.xlsx.writeFile(fileName);
  console.log("✅ Excel file written:", fileName);

  const uploadedPath = await uploadReport(fileName, "common-logistics", "xlsx");
  console.log("📤 Uploaded to Supabase:", uploadedPath);

  fs.unlinkSync(fileName); // Clean up local file
  return uploadedPath;
}

module.exports = { generateCommonLogisticsReport };
