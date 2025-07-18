const ExcelJS = require("exceljs");
const { supabase } = require("../supabaseClient");
const { uploadReport } = require("../storageUtils");
const fs = require("fs");
const path = require("path");
const os = require("os");

async function generateDropReport(programId) {
  console.log("📦 Generating Drop report for program:", programId);

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
    departure,
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

    // Step 3: Fetch Drop Logistics Type
    const logisticsTypesRes = await supabase
      .from("logistics_types")
      .select("id, name")
      .eq("program_id", programId);

    if (logisticsTypesRes.error)
      throw new Error(logisticsTypesRes.error.message);
    const dropType = logisticsTypesRes.data?.find((lt) => lt.name === "Drop"); // Changed variable name from pickupType to dropType for clarity
    if (!dropType)
      throw new Error("Drop logistics type not found for this program.");
    console.log("✅ Drop logistics type found:", dropType.name);

    // Step 4: Fetch pickup_drop_logistics (Consider renaming to drop_logistics if applicable in DB, but current use is fine)
    const logisticsRes = await supabase
      .from("pickup_drop_logistics")
      .select("registration_id, vehicle_id, logistics_type")
      .in("registration_id", registrationIds);

    if (logisticsRes.error) throw new Error(logisticsRes.error.message);
    const dropLogistics = (logisticsRes.data || []).filter(
      // Changed variable name to dropLogistics
      (l) => l.logistics_type === dropType.id // Use dropType here
    );
    console.log(
      "✅ Logistics fetched and filtered for Drop:",
      dropLogistics.length
    );

    // Step 5: Fetch vehicle names
    const vehicleIds = [
      ...new Set(dropLogistics.map((l) => l.vehicle_id)), // Use dropLogistics
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
      dropLogistics.map((l) => [
        // Use dropLogistics
        l.registration_id,
        vehicleMap[l.vehicle_id] || "",
      ])
    );
    console.log("🛠️ Registration to vehicle map created.");

    // Step 7: Prepare and sort rows
    const rows = registrations.map((r) => {
      const departureDate = r.departure // Corrected variable name
        ? r.departure.split("T")[0]
        : "";
      const departureTime = r.departure
        ? r.departure.split("T")[1]?.substring(0, 5)
        : ""; // Extract HH:MM
      return {
        departure_date: departureDate,
        departure_time: departureTime, // New column
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

    // Sort by departure_date, then departure_time, then service_no, then group_id
    rows.sort(
      (a, b) =>
        a.departure_date.localeCompare(b.departure_date) ||
        a.departure_time.localeCompare(b.departure_time) || // Added for sorting
        a.service_no.localeCompare(b.service_no) ||
        a.group_id.localeCompare(b.group_id)
    );
    console.log("🛠️ Rows prepared and sorted:", rows.length);

    // Step 8: Create Excel workbook
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Drop Report");

    sheet.columns = [
      { header: "Departure Date", key: "departure_date", width: 15 },
      { header: "Departure Time", key: "departure_time", width: 15 }, // New column header
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
    let lastDepartureDate = null;
    let lastDepartureTime = null; // New variable for merging
    let lastService = null;
    let lastGroup = null;
    let startDepartureDate = currentRow;
    let startDepartureTime = currentRow; // New variable for merging
    let startService = currentRow;
    let startGroup = currentRow;

    for (const row of rows) {
      try {
        const values = Object.values(row);
        sheet.addRow(values);

        const isNewDepartureDate = row.departure_date !== lastDepartureDate; // Corrected variable name
        // Merge "Departure Time" only if "Departure Date" and "Departure Time" are the same
        const isNewDepartureTime =
          isNewDepartureDate || row.departure_time !== lastDepartureTime;
        const isNewService =
          isNewDepartureTime || row.service_no !== lastService;
        const isNewGroup = isNewService || row.group_id !== lastGroup;

        // Merge Group ID (Column D) - shifted one column right due to new 'Departure Time' column
        if (lastGroup !== null && isNewGroup && currentRow - startGroup > 1) {
          sheet.mergeCells(`D${startGroup}:D${currentRow - 1}`);
        }
        if (isNewGroup) startGroup = currentRow;

        // Merge Service No (Column C) - shifted one column right
        if (
          lastService !== null &&
          isNewService &&
          currentRow - startService > 1
        ) {
          sheet.mergeCells(`C${startService}:C${currentRow - 1}`);
        }
        if (isNewService) startService = currentRow;

        // Merge Departure Time (Column B)
        if (
          lastDepartureTime !== null &&
          isNewDepartureTime &&
          currentRow - startDepartureTime > 1
        ) {
          sheet.mergeCells(`B${startDepartureTime}:B${currentRow - 1}`);
        }
        if (isNewDepartureTime) startDepartureTime = currentRow;

        // Merge Departure Date (Column A)
        if (
          lastDepartureDate !== null &&
          isNewDepartureDate &&
          currentRow - startDepartureDate > 1
        ) {
          sheet.mergeCells(`A${startDepartureDate}:A${currentRow - 1}`);
        }
        if (isNewDepartureDate) startDepartureDate = currentRow;

        lastDepartureDate = row.departure_date;
        lastDepartureTime = row.departure_time;
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

    // Final merge for the last group/service/departure block
    try {
      // Adjusted column letters for final merge as well
      if (currentRow - startGroup > 1)
        sheet.mergeCells(`D${startGroup}:D${currentRow - 1}`);
      if (currentRow - startService > 1)
        sheet.mergeCells(`C${startService}:C${currentRow - 1}`);
      if (currentRow - startDepartureTime > 1)
        sheet.mergeCells(`B${startDepartureTime}:B${currentRow - 1}`);
      if (currentRow - startDepartureDate > 1)
        sheet.mergeCells(`A${startDepartureDate}:A${currentRow - 1}`);
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
    const fileName = path.join(tempDir, `Drop-${Date.now()}.xlsx`);
    console.log(`📝 Attempting to write Excel file to: ${fileName}`);
    await workbook.xlsx.writeFile(fileName);
    console.log("✅ Excel file written:", fileName);

    // Upload the report
    const uploadedPath = await uploadReport(fileName, "Drop", "xlsx");
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
      "❌ An error occurred during Drop report generation:",
      error.message
    );
    console.error(error.stack);
    throw error;
  }
}

module.exports = { generateDropReport };
