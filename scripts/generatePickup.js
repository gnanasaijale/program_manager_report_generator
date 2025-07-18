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

    // Step 5: Fetch vehicle details (including new fields)
    const vehicleIds = [
      ...new Set(pickupLogistics.map((l) => l.vehicle_id)),
    ].filter(Boolean); // Filter out any null/undefined vehicle_ids

    let vehicleDetailsMap = {}; // Will store full vehicle objects, keyed by vehicle_id
    // let vehicleNameToIdMap = {}; // No longer directly needed for primary grouping keys

    if (vehicleIds.length > 0) {
      const vehiclesRes = await supabase
        .from("pickup_drop_vehicles")
        .select(
          "id, vehicle_name, vehicle_make_model, driver_name, driver_contact_number, seating_capacity"
        )
        .in("id", vehicleIds);

      if (vehiclesRes.error) throw new Error(vehiclesRes.error.message);
      vehicleDetailsMap = Object.fromEntries(
        (vehiclesRes.data || []).map((v) => [v.id, v])
      );
      // vehicleNameToIdMap = Object.fromEntries(
      //   (vehiclesRes.data || []).map((v) => [v.vehicle_name, v.id]) // This mapping can be problematic if names are not unique
      // );
    }
    console.log("🛠️ Vehicle details fetched and mapped by ID.");

    // Step 6: Map registration_id -> vehicle_id & group by vehicle_id for VehicleWise tab
    const registrationToVehicleIdMap = {}; // registration_id -> vehicle_id
    const vehicleWiseDataById = {}; // vehicle_id -> array of participants

    pickupLogistics.forEach((l) => {
      const vehicleId = l.vehicle_id;
      if (vehicleId) {
        // Check if vehicle_id exists (not null)
        registrationToVehicleIdMap[l.registration_id] = vehicleId;

        if (!vehicleWiseDataById[vehicleId]) {
          vehicleWiseDataById[vehicleId] = [];
        }
        const registration = registrations.find(
          (r) => r.id === l.registration_id
        );
        if (registration) {
          vehicleWiseDataById[vehicleId].push(registration);
        }
      }
    });
    console.log(
      "🛠️ Registration to vehicle ID map created and data grouped by vehicle ID."
    );

    // Step 7: Prepare rows for 'Detailed' tab
    const detailedRows = registrations
      .map((r) => {
        const vehicleId = registrationToVehicleIdMap[r.id];
        const vehicleName = vehicleId
          ? vehicleDetailsMap[vehicleId]?.vehicle_name || ""
          : "";

        const arrivalDate = r.arrival ? r.arrival.split("T")[0] : "";
        const arrivalTime = r.arrival
          ? r.arrival.split("T")[1]?.substring(0, 5)
          : ""; // Extract HH:MM
        return {
          id: r.id, // Keep ID for filtering
          arrival_date: arrivalDate,
          arrival_time: arrivalTime,
          service_no: r.service_no_ongoing || "",
          group_id: r.group_id || "",
          Name: `${r.participants?.first_name || ""} ${
            r.participants?.last_name || ""
          }`.trim(),
          Phone: r.participants?.primary_phone_number || "",
          City: r.participants?.city || "",
          Product: r.products?.name || "",
          Vehicle: vehicleName, // Use the correct vehicle name from map
        };
      })
      .filter((row) => row.Vehicle !== "No Pickup Needed"); // ✨ Filter out "No Pickup Needed"

    // Sort for 'Detailed' tab
    detailedRows.sort(
      (a, b) =>
        a.arrival_date.localeCompare(b.arrival_date) ||
        a.arrival_time.localeCompare(b.arrival_time) ||
        a.service_no.localeCompare(b.service_no) ||
        a.group_id.localeCompare(b.group_id)
    );
    console.log("🛠️ Rows prepared and sorted for 'Detailed' tab.");

    // Step 8: Prepare rows for 'VehicleWise' tab
    const vehicleWiseRows = [];
    const sortedVehicleIds = Object.keys(vehicleWiseDataById).sort(
      (idA, idB) => {
        // Sort vehicles by their names first, then by ID for stability
        const nameA = vehicleDetailsMap[idA]?.vehicle_name || "";
        const nameB = vehicleDetailsMap[idB]?.vehicle_name || "";
        if (nameA !== nameB) return nameA.localeCompare(nameB);
        return idA.localeCompare(idB); // Fallback to ID sort
      }
    );

    // Calculate earliest pickup time for each vehicle (by vehicle_id)
    const vehiclePickupTimesById = {}; // Keyed by vehicle_id
    for (const vehicleId of sortedVehicleIds) {
      const vehicleName = vehicleDetailsMap[vehicleId]?.vehicle_name || "";
      if (vehicleName === "No Pickup Needed") continue; // Skip calculation for "No Pickup Needed"

      const participantsForVehicle = vehicleWiseDataById[vehicleId];
      let earliestArrival = null;

      for (const participant of participantsForVehicle) {
        if (participant.arrival) {
          const currentArrival = new Date(participant.arrival);
          if (earliestArrival === null || currentArrival < earliestArrival) {
            earliestArrival = currentArrival;
          }
        }
      }
      if (earliestArrival) {
        const localDate = earliestArrival
          .toLocaleDateString("en-IN", {
            year: "numeric",
            month: "2-digit",
            day: "2-digit",
            timeZone: "Asia/Kolkata",
          })
          .split("/")
          .reverse()
          .join("-");
        const localTime = earliestArrival.toLocaleTimeString("en-IN", {
          hour: "2-digit",
          minute: "2-digit",
          hour12: false,
          timeZone: "Asia/Kolkata",
        });

        vehiclePickupTimesById[vehicleId] = {
          date: localDate,
          time: localTime,
        };
      } else {
        vehiclePickupTimesById[vehicleId] = { date: "", time: "" };
      }
    }

    for (const vehicleId of sortedVehicleIds) {
      const vehicleName = vehicleDetailsMap[vehicleId]?.vehicle_name || "";
      if (vehicleName === "No Pickup Needed") continue; // Skip "No Pickup Needed" vehicle group

      const participants = vehicleWiseDataById[vehicleId];

      // Sort participants within each vehicle by arrival date, then time
      // This 'arrival' is the raw UTC string from Supabase
      participants.sort((a, b) =>
        (a.arrival || "").localeCompare(b.arrival || "")
      );

      for (const r of participants) {
        // Filter out individuals with "No Pickup Needed" from being added to VehicleWise rows
        // This check is redundant here if the top-level vehicle group is already skipped
        // but can be a safeguard if logic changes.
        if (
          registrationToVehicleIdMap[r.id] &&
          vehicleDetailsMap[registrationToVehicleIdMap[r.id]]?.vehicle_name ===
            "No Pickup Needed"
        ) {
          continue;
        }

        const arrivalDate = r.arrival ? r.arrival.split("T")[0] : "";
        const arrivalTime = r.arrival
          ? r.arrival.split("T")[1]?.substring(0, 5)
          : "";
        vehicleWiseRows.push({
          vehicle_name: vehicleName, // Use the actual name for the row
          arrival_date: arrivalDate,
          arrival_time: arrivalTime,
          service_no: r.service_no_ongoing || "",
          group_id: r.group_id || "",
          Name: `${r.participants?.first_name || ""} ${
            r.participants?.last_name || ""
          }`.trim(),
          Phone: r.participants?.primary_phone_number || "",
          City: r.participants?.city || "",
          Product: r.products?.name || "",
        });
      }
    }
    console.log("🛠️ Rows prepared for 'VehicleWise' tab.");

    // Step 9: Prepare rows for 'Vehicles' tab
    // Iterate over vehicleDetailsMap directly to get all unique vehicles from the DB
    const vehiclesTabRows = Object.values(vehicleDetailsMap)
      .map((v) => {
        // Filter out "No Pickup Needed" vehicle
        if (v.vehicle_name === "No Pickup Needed") {
          return null;
        }

        const pickupInfo = vehiclePickupTimesById[v.id] || {
          // Get derived pickup time using vehicle_id
          date: "",
          time: "",
        };

        return {
          name: v.vehicle_name,
          vehicle_make_model: v.vehicle_make_model || "",
          driver_name: v.driver_name || "",
          driver_contact_number: v.driver_contact_number || "",
          pickup_date: pickupInfo.date,
          pickup_time: pickupInfo.time,
        };
      })
      .filter(Boolean) // Remove null entries
      .sort((a, b) => a.name.localeCompare(b.name)); // Sort by vehicle name alphabetically
    console.log("🛠️ Rows prepared for 'Vehicles' tab.");

    // Step 10: Create Excel workbook and sheets
    const workbook = new ExcelJS.Workbook();

    // --- Detailed Tab ---
    const detailedSheet = workbook.addWorksheet("Detailed");
    detailedSheet.columns = [
      { header: "Arrival Date", key: "arrival_date", width: 15 },
      { header: "Arrival Time", key: "arrival_time", width: 15 },
      { header: "Service No", key: "service_no", width: 15 },
      { header: "Group ID", key: "group_id", width: 15 },
      { header: "Name", key: "Name", width: 25 },
      { header: "Phone", key: "Phone", width: 15 },
      { header: "City", key: "City", width: 15 },
      { header: "Product", key: "Product", width: 20 },
      { header: "Vehicle", key: "Vehicle", width: 20 },
    ];
    detailedSheet.getRow(1).font = { bold: true };
    console.log("🛠️ 'Detailed' sheet created, columns set.");

    // Add rows and apply merging for 'Detailed' tab
    let currentRowDetailed = 2;
    let lastArrivalDateDetailed = null;
    let lastArrivalTimeDetailed = null;
    let lastServiceDetailed = null;
    let lastGroupDetailed = null;
    let startArrivalDateDetailed = currentRowDetailed;
    let startArrivalTimeDetailed = currentRowDetailed;
    let startServiceDetailed = currentRowDetailed;
    let startGroupDetailed = currentRowDetailed;

    for (const row of detailedRows) {
      const values = Object.values(row);
      detailedSheet.addRow(values);

      const isNewArrivalDate = row.arrival_date !== lastArrivalDateDetailed;
      const isNewArrivalTime =
        isNewArrivalDate || row.arrival_time !== lastArrivalTimeDetailed;
      const isNewService =
        isNewArrivalTime || row.service_no !== lastServiceDetailed;
      const isNewGroup = isNewService || row.group_id !== lastGroupDetailed;

      if (
        lastGroupDetailed !== null &&
        isNewGroup &&
        currentRowDetailed - startGroupDetailed > 1
      ) {
        detailedSheet.mergeCells(
          `D${startGroupDetailed}:D${currentRowDetailed - 1}`
        );
      }
      if (isNewGroup) startGroupDetailed = currentRowDetailed;

      if (
        lastServiceDetailed !== null &&
        isNewService &&
        currentRowDetailed - startServiceDetailed > 1
      ) {
        detailedSheet.mergeCells(
          `C${startServiceDetailed}:C${currentRowDetailed - 1}`
        );
      }
      if (isNewService) startServiceDetailed = currentRowDetailed;

      if (
        lastArrivalTimeDetailed !== null &&
        isNewArrivalTime &&
        currentRowDetailed - startArrivalTimeDetailed > 1
      ) {
        detailedSheet.mergeCells(
          `B${startArrivalTimeDetailed}:B${currentRowDetailed - 1}`
        );
      }
      if (isNewArrivalTime) startArrivalTimeDetailed = currentRowDetailed;

      if (
        lastArrivalDateDetailed !== null &&
        isNewArrivalDate &&
        currentRowDetailed - startArrivalDateDetailed > 1
      ) {
        detailedSheet.mergeCells(
          `A${startArrivalDateDetailed}:A${currentRowDetailed - 1}`
        );
      }
      if (isNewArrivalDate) startArrivalDateDetailed = currentRowDetailed;

      lastArrivalDateDetailed = row.arrival_date;
      lastArrivalTimeDetailed = row.arrival_time;
      lastServiceDetailed = row.service_no;
      lastGroupDetailed = row.group_id;
      currentRowDetailed++;
    }
    // Final merge for the last block in 'Detailed' tab
    if (currentRowDetailed - startGroupDetailed > 1)
      detailedSheet.mergeCells(
        `D${startGroupDetailed}:D${currentRowDetailed - 1}`
      );
    if (currentRowDetailed - startServiceDetailed > 1)
      detailedSheet.mergeCells(
        `C${startServiceDetailed}:C${currentRowDetailed - 1}`
      );
    if (currentRowDetailed - startArrivalTimeDetailed > 1)
      detailedSheet.mergeCells(
        `B${startArrivalTimeDetailed}:B${currentRowDetailed - 1}`
      );
    if (currentRowDetailed - startArrivalDateDetailed > 1)
      detailedSheet.mergeCells(
        `A${startArrivalDateDetailed}:A${currentRowDetailed - 1}`
      );
    console.log("🛠️ Rows added and merges applied for 'Detailed' tab.");

    // --- VehicleWise Tab ---
    const vehicleWiseSheet = workbook.addWorksheet("VehicleWise");
    vehicleWiseSheet.columns = [
      { header: "Vehicle", key: "vehicle_name", width: 20 },
      { header: "Arrival Date", key: "arrival_date", width: 15 },
      { header: "Arrival Time", key: "arrival_time", width: 15 },
      { header: "Service No", key: "service_no", width: 15 },
      { header: "Group ID", key: "group_id", width: 15 },
      { header: "Name", key: "Name", width: 25 },
      { header: "Phone", key: "Phone", width: 15 },
      { header: "City", key: "City", width: 15 },
      { header: "Product", key: "Product", width: 20 },
    ];
    vehicleWiseSheet.getRow(1).font = { bold: true };
    console.log("🛠️ 'VehicleWise' sheet created, columns set.");

    // Add rows and apply merging for 'VehicleWise' tab
    let currentRowVehicleWise = 2;
    let lastVehicleName = null; // Still merge by displayed name
    let lastArrivalDateVehicleWise = null;
    let lastArrivalTimeVehicleWise = null;
    let startVehicleName = currentRowVehicleWise;
    let startArrivalDateVehicleWise = currentRowVehicleWise;
    let startArrivalTimeVehicleWise = currentRowVehicleWise;

    for (const row of vehicleWiseRows) {
      const values = Object.values(row);
      vehicleWiseSheet.addRow(values);

      const isNewVehicle = row.vehicle_name !== lastVehicleName;
      const isNewArrivalDate =
        isNewVehicle || row.arrival_date !== lastArrivalDateVehicleWise;
      const isNewArrivalTime =
        isNewArrivalDate || row.arrival_time !== lastArrivalTimeVehicleWise;

      // Merge Arrival Time (Column C)
      if (
        lastArrivalTimeVehicleWise !== null &&
        isNewArrivalTime &&
        currentRowVehicleWise - startArrivalTimeVehicleWise > 1
      ) {
        vehicleWiseSheet.mergeCells(
          `C${startArrivalTimeVehicleWise}:C${currentRowVehicleWise - 1}`
        );
      }
      if (isNewArrivalTime) startArrivalTimeVehicleWise = currentRowVehicleWise;

      // Merge Arrival Date (Column B)
      if (
        lastArrivalDateVehicleWise !== null &&
        isNewArrivalDate &&
        currentRowVehicleWise - startArrivalDateVehicleWise > 1
      ) {
        vehicleWiseSheet.mergeCells(
          `B${startArrivalDateVehicleWise}:B${currentRowVehicleWise - 1}`
        );
      }
      if (isNewArrivalDate) startArrivalDateVehicleWise = currentRowVehicleWise;

      // Merge Vehicle Name (Column A)
      if (
        lastVehicleName !== null &&
        isNewVehicle &&
        currentRowVehicleWise - startVehicleName > 1
      ) {
        vehicleWiseSheet.mergeCells(
          `A${startVehicleName}:A${currentRowVehicleWise - 1}`
        );
      }
      if (isNewVehicle) startVehicleName = currentRowVehicleWise;

      lastVehicleName = row.vehicle_name;
      lastArrivalDateVehicleWise = row.arrival_date;
      lastArrivalTimeVehicleWise = row.arrival_time;
      currentRowVehicleWise++;
    }
    // Final merge for the last block in 'VehicleWise' tab
    if (currentRowVehicleWise - startArrivalTimeVehicleWise > 1)
      vehicleWiseSheet.mergeCells(
        `C${startArrivalTimeVehicleWise}:C${currentRowVehicleWise - 1}`
      );
    if (currentRowVehicleWise - startArrivalDateVehicleWise > 1)
      vehicleWiseSheet.mergeCells(
        `B${startArrivalDateVehicleWise}:B${currentRowVehicleWise - 1}`
      );
    if (currentRowVehicleWise - startVehicleName > 1)
      vehicleWiseSheet.mergeCells(
        `A${startVehicleName}:A${currentRowVehicleWise - 1}`
      );
    console.log("🛠️ Rows added and merges applied for 'VehicleWise' tab.");

    // --- Vehicles Tab ---
    const vehiclesSheet = workbook.addWorksheet("Vehicles");
    vehiclesSheet.columns = [
      { header: "Vehicle Name", key: "name", width: 25 },
      { header: "Make/Model", key: "vehicle_make_model", width: 20 },
      { header: "Driver Name", key: "driver_name", width: 20 },
      { header: "Driver Phone", key: "driver_contact_number", width: 20 },
      { header: "Earliest Pickup Date", key: "pickup_date", width: 20 },
      { header: "Earliest Pickup Time", key: "pickup_time", width: 20 },
    ];
    vehiclesSheet.getRow(1).font = { bold: true };
    vehiclesSheet.addRows(vehiclesTabRows);
    console.log("🛠️ 'Vehicles' sheet created and rows added.");

    // Step 11: Generate a unique filename and write to a temporary directory
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
