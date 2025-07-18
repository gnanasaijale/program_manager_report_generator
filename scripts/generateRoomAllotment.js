const ExcelJS = require("exceljs");
const { supabase } = require("../supabaseClient");
const { uploadReport } = require("../storageUtils");
const fs = require("fs");

async function generateRoomAllotmentReport(programId) {
  console.log("🛏️ Generating Room Allocation report for program:", programId);

  // Step 1: Fetch orders for the program
  const ordersRes = await supabase
    .from("orders")
    .select("id")
    .eq("program_id", programId);

  if (ordersRes.error) {
    console.error("❌ Error fetching orders:", ordersRes.error);
    throw new Error(ordersRes.error.message);
  }

  const orderIds = ordersRes.data.map((o) => o.id);
  console.log("🛏️ Order IDs fetched:", orderIds);

  if (!orderIds.length) {
    throw new Error("No orders found for program");
  }

  // Step 2: Fetch registrations linked to those orders
  const registrationsRes = await supabase
    .from("registrations")
    .select(
      `
      id,
      bed_number,
      room_id,
      rooms(room_number, checkin_date, checkout_date),
      participants(first_name, last_name, primary_phone_number, email, city)
    `
    )
    .in("order_id", orderIds);

  if (registrationsRes.error) {
    console.error("❌ Error fetching registrations:", registrationsRes.error);
    throw new Error(registrationsRes.error.message);
  }

  let data = registrationsRes.data || [];
  console.log("🛏️ Total Registrations fetched:", data.length);

  // ✅ Skip entries without room allocation
  data = data.filter((r) => r.rooms?.room_number);
  console.log("✅ Registrations with room allocation:", data.length);

  if (!data.length) {
    throw new Error("No room allocations found");
  }

  // Step 3: Transform to rows
  const rows = data.map((r, index) => ({
    sr_no: index + 1,
    room_no: r.rooms?.room_number || "",
    checkin_date: r.rooms?.checkin_date || "",
    checkout_date: r.rooms?.checkout_date || "",
    bed: r.bed_number || "",
    first_name: r.participants?.first_name || "",
    last_name: r.participants?.last_name || "",
    phone: r.participants?.primary_phone_number || "",
    email: r.participants?.email || "",
    city: r.participants?.city || "",
  }));

  console.log("📋 Processed rows ready for Excel:", rows.length);

  // Step 4: Create Excel workbook
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Room Allocation");

  sheet.columns = [
    { header: "Sr No", key: "sr_no", width: 6 },
    { header: "Room No.", key: "room_no", width: 15 },
    { header: "Checkin Date", key: "checkin_date", width: 20 },
    { header: "Checkout Date", key: "checkout_date", width: 20 },
    { header: "Allot", key: "bed", width: 10 },
    { header: "Name", key: "first_name", width: 15 },
    { header: "Surname", key: "last_name", width: 15 },
    { header: "Contact No.", key: "phone", width: 18 },
    { header: "Email ID", key: "email", width: 25 },
    { header: "City", key: "city", width: 15 },
  ];

  sheet.addRows(rows);
  console.log("📊 Excel worksheet populated.");

  // Step 5: Save and upload
  const fileName = `room-allocation-${Date.now()}.xlsx`;
  await workbook.xlsx.writeFile(fileName);
  console.log("✅ Excel file written:", fileName);

  const uploadedPath = await uploadReport(fileName, "room_allocation", "xlsx");
  console.log("📤 Uploaded to Supabase Storage:", uploadedPath);

  fs.unlinkSync(fileName); // delete local file
  console.log("🧹 Temporary file cleaned up.");

  return uploadedPath;
}

module.exports = { generateRoomAllotmentReport };
