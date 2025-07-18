const ExcelJS = require("exceljs");
const { supabase } = require("../supabaseClient");
const { uploadReport } = require("../storageUtils");
const fs = require("fs");
const path = require("path");
const os = require("os");

async function generatePickupReport(programId) {
  console.log("📦 Generating Pickup report for program:", programId);

  try {
    // Step 1: Fetch Orders
    const ordersRes = await supabase
      .from("orders")
      .select("id")
      .eq("program_id", programId);
    if (ordersRes.error) throw new Error(ordersRes.error.message);
    const orderIds = ordersRes.data.map((o) => o.id);
    console.log("✅ Orders fetched:", orderIds.length);
    if (!orderIds.length) throw new Error("No orders found for program");

    // Step 2: Fetch Registrations
    const registrationsRes = await supabase
      .from("registrations")
      .select(
        `
    id,
    arrival,
    service_no_ongoing,
    group_id,
    participants(first_name, last_name, primary_phone_number, city),
    products(name)
  `
      )
      .in("order_id", orderIds);

    if (registrationsRes.error) throw new Error(registrationsRes.error.message);
    const registrations = registrationsRes.data || [];
    console.log("✅ Registrations fetched:", registrations.length);
    const registrationIds = registrations.map((r) => r.id);
    if (!registrationIds.length)
      throw new Error("No registrations found for these orders.");

    // Step 3: Fetch Pickup Logistics Type
    const logisticsTypesRes = await supabase
      .from("logistics_types")
      .select("id, name")
      .eq("program_id", programId);

    if (logisticsTypesRes.error)
      throw new Error(logisticsTypesRes.error.message);
    const pickupType = logisticsTypesRes.data?.find(
      (lt) => lt.name === "Pickup"
    );
    if (!pickupType)
      throw new Error("Pickup logistics type not found for this program.");
    console.log("✅ Pickup logistics type found:", pickupType.name);

    // Step 4: Fetch pickup_drop_logistics
    const logisticsRes = await supabase
      .from("pickup_drop_logistics")
      .select("registration_id, vehicle_id, logistics_type")
      .in("registration_id", registrationIds);

    if (logisticsRes.error) throw new Error(logisticsRes.error.message);
    const pickupLogistics = (logisticsRes.data || []).filter(
      (l) => l.logistics_type === pickupType.id
    );
    console.log(
      "✅ Logistics fetched and filtered for pickup:",
      pickupLogistics.length
    );

    // Step 5: Fetch vehicle names
    const vehicleIds = [
      ...new Set(pickupLogistics.map((l) => l.vehicle_id)),
    ].filter(Boolean);
    let vehicleMap = {};
    if (vehicleIds.length > 0) {
      const vehiclesRes = await supabase
        .from("pickup_drop_vehicles")
        .select("id, vehicle_name")
        .in("id", vehicleIds);

      if (vehiclesRes.error) throw new Error(vehiclesRes.error.message);
      vehicleMap = Object.fromEntries(
        (vehiclesRes.data || []).map((v) => [v.id, v.vehicle_name])
      );
    }
    console.log("🛠️ Vehicle names fetched and mapped.");

    // Step 6: Map registration_id -> vehicle name
    const logisticsMap = Object.fromEntries(
      pickupLogistics.map((l) => [
        l.registration_id,
        vehicleMap[l.vehicle_id] || "",
      ])
    );
    console.log("🛠️ Registration to vehicle map created.");

    // Step 7: Prepare and sort rows
    const rows = registrations.map((r) => {
      const arrivalDate = r.arrival ? r.arrival.split("T")[0] : "";
      const arrivalTime = r.arrival
        ? r.arrival.split("T")[1]?.substring(0, 5)
        : ""; // Extract HH:MM
      return {
        arrival_date: arrivalDate,
        arrival_time: arrivalTime, // New column
        service_no: r.service_no_ongoing || "",
        group_id: r.group_id || "",
        Name: `${r.participants?.first_name || ""} ${
          r.participants?.last_name || ""
        }`.trim(),
        Phone: r.participants?.primary_phone_number || "",
        City: r.participants?.city || "",
        Product: r.products?.name || "",
        Vehicle: logisticsMap[r.id] || "",
      };
    });

    // Sort by arrival_date, then arrival_time, then service_no, then group_id
    rows.sort(
      (a, b) =>
        a.arrival_date.localeCompare(b.arrival_date) ||
        a.arrival_time.localeCompare(b.arrival_time) || // Added for sorting
        a.service_no.localeCompare(b.service_no) ||
        a.group_id.localeCompare(b.group_id)
    );
    console.log("🛠️ Rows prepared and sorted:", rows.length);

    // Step 8: Create Excel workbook
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Pickup Report");

    sheet.columns = [
      { header: "Arrival Date", key: "arrival_date", width: 15 },
      { header: "Arrival Time", key: "arrival_time", width: 15 }, // New column header
      { header: "Service No", key: "service_no", width: 15 },
      { header: "Group ID", key: "group_id", width: 15 },
      { header: "Name", key: "Name", width: 25 },
      { header: "Phone", key: "Phone", width: 15 },
      { header: "City", key: "City", width: 15 },
      { header: "Product", key: "Product", width: 20 },
      { header: "Vehicle", key: "Vehicle", width: 20 },
    ];

    sheet.getRow(1).font = { bold: true };
    console.log("🛠️ Excel workbook and sheet created, columns set.");

    // Step 9: Add rows and track merge positions
    let currentRow = 2;
    let lastArrivalDate = null;
    let lastArrivalTime = null; // New variable for merging
    let lastService = null;
    let lastGroup = null;
    let startArrivalDate = currentRow;
    let startArrivalTime = currentRow; // New variable for merging
    let startService = currentRow;
    let startGroup = currentRow;

    for (const row of rows) {
      try {
        const values = Object.values(row);
        sheet.addRow(values);

        const isNewArrivalDate = row.arrival_date !== lastArrivalDate;
        // Merge "Arrival Time" only if "Arrival Date" and "Arrival Time" are the same
        const isNewArrivalTime =
          isNewArrivalDate || row.arrival_time !== lastArrivalTime;
        const isNewService = isNewArrivalTime || row.service_no !== lastService;
        const isNewGroup = isNewService || row.group_id !== lastGroup;

        // Merge Group ID (Column D)
        if (lastGroup !== null && isNewGroup && currentRow - startGroup > 1) {
          sheet.mergeCells(`D${startGroup}:D${currentRow - 1}`);
        }
        if (isNewGroup) startGroup = currentRow;

        // Merge Service No (Column C)
        if (
          lastService !== null &&
          isNewService &&
          currentRow - startService > 1
        ) {
          sheet.mergeCells(`C${startService}:C${currentRow - 1}`);
        }
        if (isNewService) startService = currentRow;

        // Merge Arrival Time (Column B)
        if (
          lastArrivalTime !== null &&
          isNewArrivalTime &&
          currentRow - startArrivalTime > 1
        ) {
          sheet.mergeCells(`B${startArrivalTime}:B${currentRow - 1}`);
        }
        if (isNewArrivalTime) startArrivalTime = currentRow;

        // Merge Arrival Date (Column A)
        if (
          lastArrivalDate !== null &&
          isNewArrivalDate &&
          currentRow - startArrivalDate > 1
        ) {
          sheet.mergeCells(`A${startArrivalDate}:A${currentRow - 1}`);
        }
        if (isNewArrivalDate) startArrivalDate = currentRow;

        lastArrivalDate = row.arrival_date;
        lastArrivalTime = row.arrival_time;
        lastService = row.service_no;
        lastGroup = row.group_id;
        currentRow++;
      } catch (rowProcessingError) {
        console.error(
          "❌ Error processing row or merging cells for row:",
          row,
          rowProcessingError.message
        );
        throw rowProcessingError;
      }
    }
    console.log(
      "🛠️ All rows added and merging logic applied to individual rows."
    );

    // Final merge for the last group/service/arrival block
    try {
      if (currentRow - startGroup > 1)
        sheet.mergeCells(`D${startGroup}:D${currentRow - 1}`);
      if (currentRow - startService > 1)
        sheet.mergeCells(`C${startService}:C${currentRow - 1}`);
      if (currentRow - startArrivalTime > 1)
        sheet.mergeCells(`B${startArrivalTime}:B${currentRow - 1}`);
      if (currentRow - startArrivalDate > 1)
        sheet.mergeCells(`A${startArrivalDate}:A${currentRow - 1}`);
      console.log("🛠️ Final merges applied to the last block.");
    } catch (finalMergeError) {
      console.error(
        "❌ Error during final cell merging:",
        finalMergeError.message
      );
      throw finalMergeError;
    }

    // Generate a unique filename and write to a temporary directory
    const tempDir = os.tmpdir();
    const fileName = path.join(tempDir, `pickup-${Date.now()}.xlsx`);
    console.log(`📝 Attempting to write Excel file to: ${fileName}`);
    await workbook.xlsx.writeFile(fileName);
    console.log("✅ Excel file written:", fileName);

    // Upload the report
    const uploadedPath = await uploadReport(fileName, "Pickup", "xlsx");
    console.log("📤 Uploaded to Supabase:", uploadedPath);

    // Clean up the local file
    try {
      fs.unlinkSync(fileName);
      console.log("🗑️ Local file deleted:", fileName);
    } catch (unlinkError) {
      console.warn(
        "⚠️ Could not delete local file:",
        fileName,
        unlinkError.message
      );
    }

    return uploadedPath;
  } catch (error) {
    console.error(
      "❌ An error occurred during pickup report generation:",
      error.message
    );
    console.error(error.stack);
    throw error;
  }
}

module.exports = { generatePickupReport };
