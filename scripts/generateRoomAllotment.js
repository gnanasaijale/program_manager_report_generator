// const ExcelJS = require("exceljs");
// const { supabase } = require("../supabaseClient");
// const { uploadReport } = require("../storageUtils");
// const fs = require("fs");

// async function generateRoomAllotmentReport(programId) {
//   console.log("🛏️ Generating Room Allocation report for program:", programId);

//   // Step 1: Fetch orders for the program
//   const ordersRes = await supabase
//     .from("orders")
//     .select("id")
//     .eq("program_id", programId);

//   if (ordersRes.error) {
//     console.error("❌ Error fetching orders:", ordersRes.error);
//     throw new Error(ordersRes.error.message);
//   }

//   const orderIds = ordersRes.data.map((o) => o.id);
//   console.log("🛏️ Order IDs fetched:", orderIds);

//   if (!orderIds.length) {
//     throw new Error("No orders found for program");
//   }

//   // Step 2: Fetch registrations linked to those orders
//   const registrationsRes = await supabase
//     .from("registrations")
//     .select(
//       `
//       id,
//       bed_number,
//       room_id,
//       rooms(room_number, checkin_date, checkout_date),
//       participants(first_name, last_name, primary_phone_number, email, city)
//     `
//     )
//     .in("order_id", orderIds);

//   if (registrationsRes.error) {
//     console.error("❌ Error fetching registrations:", registrationsRes.error);
//     throw new Error(registrationsRes.error.message);
//   }

//   let data = registrationsRes.data || [];
//   console.log("🛏️ Total Registrations fetched:", data.length);

//   // ✅ Skip entries without room allocation
//   data = data.filter((r) => r.rooms?.room_number);
//   console.log("✅ Registrations with room allocation:", data.length);

//   if (!data.length) {
//     throw new Error("No room allocations found");
//   }

//   // Step 3: Transform to rows
//   const rows = data.map((r, index) => ({
//     sr_no: index + 1,
//     room_no: r.rooms?.room_number || "",
//     checkin_date: r.rooms?.checkin_date || "",
//     checkout_date: r.rooms?.checkout_date || "",
//     bed: r.bed_number || "",
//     first_name: r.participants?.first_name || "",
//     last_name: r.participants?.last_name || "",
//     phone: r.participants?.primary_phone_number || "",
//     email: r.participants?.email || "",
//     city: r.participants?.city || "",
//   }));

//   console.log("📋 Processed rows ready for Excel:", rows.length);

//   // Step 4: Create Excel workbook
//   const workbook = new ExcelJS.Workbook();
//   const sheet = workbook.addWorksheet("Room Allocation");

//   sheet.columns = [
//     { header: "Sr No", key: "sr_no", width: 6 },
//     { header: "Room No.", key: "room_no", width: 15 },
//     { header: "Checkin Date", key: "checkin_date", width: 20 },
//     { header: "Checkout Date", key: "checkout_date", width: 20 },
//     { header: "Allot", key: "bed", width: 10 },
//     { header: "Name", key: "first_name", width: 15 },
//     { header: "Surname", key: "last_name", width: 15 },
//     { header: "Contact No.", key: "phone", width: 18 },
//     { header: "Email ID", key: "email", width: 25 },
//     { header: "City", key: "city", width: 15 },
//   ];

//   sheet.addRows(rows);
//   console.log("📊 Excel worksheet populated.");

//   // Step 5: Save and upload
//   const fileName = `room-allocation-${Date.now()}.xlsx`;
//   await workbook.xlsx.writeFile(fileName);
//   console.log("✅ Excel file written:", fileName);

//   const uploadedPath = await uploadReport(fileName, "room_allocation", "xlsx");
//   console.log("📤 Uploaded to Supabase Storage:", uploadedPath);

//   fs.unlinkSync(fileName); // delete local file
//   console.log("🧹 Temporary file cleaned up.");

//   return uploadedPath;
// }

// module.exports = { generateRoomAllotmentReport };

const ExcelJS = require("exceljs");
const { supabase } = require("../supabaseClient");
const { uploadReport } = require("../storageUtils");
const fs = require("fs");
const path = require("path");
const os = require("os");

async function generateRoomAllotmentReport(programId) {
  console.log("🛏️ Generating Room Allocation report for program:", programId);

  try {
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
    console.log("🛏️ Order IDs fetched:", orderIds.length);

    if (!orderIds.length) {
      throw new Error("No orders found for program");
    }

    // Step 2: Fetch registrations linked to those orders, including hotel name
    const registrationsRes = await supabase
      .from("registrations")
      .select(
        `
        id,
        bed_number,
        room_id,
        rooms(room_number, checkin_date, checkout_date, name_of_the_hotel),
        participants(first_name, last_name, primary_phone_number, email, city),
        orders(order_id)
        ` // <--- Removed the comment here
      )
      .in("orders.order_id", orderIds);

    if (registrationsRes.error) {
      console.error("❌ Error fetching registrations:", registrationsRes.error);
      throw new Error(registrationsRes.error.message);
    }

    let data = registrationsRes.data || [];
    console.log("🛏️ Total Registrations fetched:", data.length);

    // Filter out entries without room allocation
    data = data.filter((r) => r.rooms?.room_number);
    console.log("✅ Registrations with room allocation:", data.length);

    if (!data.length) {
      throw new Error("No room allocations found");
    }

    // Step 3: Transform and categorize data for different tabs
    const allRoomsRows = [];
    const groupedByHotel = {}; // { "Hotel Name": [ {row}, {row} ] }

    data.forEach((r, index) => {
      const hotelName = r.rooms?.name_of_the_hotel || "Unspecified Hotel"; // Extract hotel name
      const row = {
        sr_no: index + 1, // Sr No for overall report
        room_no: r.rooms?.room_number || "",
        checkin_date: r.rooms?.checkin_date || "",
        checkout_date: r.rooms?.checkout_date || "",
        bed: r.bed_number || "",
        first_name: r.participants?.first_name || "",
        last_name: r.participants?.last_name || "",
        phone: r.participants?.primary_phone_number || "",
        email: r.participants?.email || "",
        city: r.participants?.city || "",
        hotel_name: hotelName, // Store hotel name for internal use
      };
      allRoomsRows.push(row);

      if (!groupedByHotel[hotelName]) {
        groupedByHotel[hotelName] = [];
      }
      groupedByHotel[hotelName].push(row);
    });

    console.log("📋 Processed rows ready for Excel.");

    // Step 4: Define common Excel column headers
    const commonColumns = [
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

    // Step 5: Create Excel workbook
    const workbook = new ExcelJS.Workbook();

    // --- Tab 1: All rooms ---
    const allRoomsSheet = workbook.addWorksheet("All rooms");
    allRoomsSheet.columns = commonColumns;
    allRoomsSheet.getRow(1).font = { bold: true };

    // Sort 'All rooms' data: Room No, Checkin, Checkout, Bed
    allRoomsRows.sort((a, b) => {
      if (a.room_no !== b.room_no) return a.room_no.localeCompare(b.room_no);
      if (a.checkin_date !== b.checkin_date)
        return a.checkin_date.localeCompare(b.checkin_date);
      if (a.checkout_date !== b.checkout_date)
        return a.checkout_date.localeCompare(b.checkout_date);
      return (a.bed || "").localeCompare(b.bed || ""); // Sort by bed number if available
    });

    // Add rows and apply merging for 'All rooms' tab
    let currentRowAllRooms = 2;
    let lastRoomNoAllRooms = null;
    let lastCheckinAllRooms = null;
    let lastCheckoutAllRooms = null;
    let startRoomNoAllRooms = currentRowAllRooms;
    let startCheckinAllRooms = currentRowAllRooms;
    let startCheckoutAllRooms = currentRowAllRooms;

    for (const row of allRoomsRows) {
      const values = Object.values(row);
      allRoomsSheet.addRow(values); // Add all values including hotel_name, ExcelJS ignores extra keys

      const isNewRoomNo = row.room_no !== lastRoomNoAllRooms;
      const isNewCheckin =
        isNewRoomNo || row.checkin_date !== lastCheckinAllRooms;
      const isNewCheckout =
        isNewCheckin || row.checkout_date !== lastCheckoutAllRooms;

      // Merge Checkout Date (Column D)
      if (
        lastCheckoutAllRooms !== null &&
        isNewCheckout &&
        currentRowAllRooms - startCheckoutAllRooms > 1
      ) {
        allRoomsSheet.mergeCells(
          `D${startCheckoutAllRooms}:D${currentRowAllRooms - 1}`
        );
      }
      if (isNewCheckout) startCheckoutAllRooms = currentRowAllRooms;

      // Merge Checkin Date (Column C)
      if (
        lastCheckinAllRooms !== null &&
        isNewCheckin &&
        currentRowAllRooms - startCheckinAllRooms > 1
      ) {
        allRoomsSheet.mergeCells(
          `C${startCheckinAllRooms}:C${currentRowAllRooms - 1}`
        );
      }
      if (isNewCheckin) startCheckinAllRooms = currentRowAllRooms;

      // Merge Room No. (Column B)
      if (
        lastRoomNoAllRooms !== null &&
        isNewRoomNo &&
        currentRowAllRooms - startRoomNoAllRooms > 1
      ) {
        allRoomsSheet.mergeCells(
          `B${startRoomNoAllRooms}:B${currentRowAllRooms - 1}`
        );
      }
      if (isNewRoomNo) startRoomNoAllRooms = currentRowAllRooms;

      lastRoomNoAllRooms = row.room_no;
      lastCheckinAllRooms = row.checkin_date;
      lastCheckoutAllRooms = row.checkout_date;
      currentRowAllRooms++;
    }
    // Final merge for the last block in 'All rooms' tab
    if (currentRowAllRooms - startCheckoutAllRooms > 1)
      allRoomsSheet.mergeCells(
        `D${startCheckoutAllRooms}:D${currentRowAllRooms - 1}`
      );
    if (currentRowAllRooms - startCheckinAllRooms > 1)
      allRoomsSheet.mergeCells(
        `C${startCheckinAllRooms}:C${currentRowAllRooms - 1}`
      );
    if (currentRowAllRooms - startRoomNoAllRooms > 1)
      allRoomsSheet.mergeCells(
        `B${startRoomNoAllRooms}:B${currentRowAllRooms - 1}`
      );
    console.log("📊 'All rooms' sheet populated and merged.");

    // --- Tab 2+: Each Hotel ---
    const sortedHotelNames = Object.keys(groupedByHotel).sort(); // Sort hotel tabs alphabetically
    for (const hotelName of sortedHotelNames) {
      // Sanitize sheet name (Excel sheet names have character limits and cannot contain certain characters)
      const sheetName = hotelName
        .substring(0, 31)
        .replace(/[\[\]\*\?\/\\:]/g, ""); // Max 31 chars, remove invalid chars
      const hotelSheet = workbook.addWorksheet(sheetName);
      hotelSheet.columns = commonColumns;
      hotelSheet.getRow(1).font = { bold: true };

      const hotelRows = groupedByHotel[hotelName];

      // Sort hotel-specific data: Room No, Checkin, Checkout, Bed
      hotelRows.sort((a, b) => {
        if (a.room_no !== b.room_no) return a.room_no.localeCompare(b.room_no);
        if (a.checkin_date !== b.checkin_date)
          return a.checkin_date.localeCompare(b.checkin_date);
        if (a.checkout_date !== b.checkout_date)
          return a.checkout_date.localeCompare(b.checkout_date);
        return (a.bed || "").localeCompare(b.bed || "");
      });

      // Add rows and apply merging for hotel-specific tab
      let currentRowHotel = 2;
      let lastRoomNoHotel = null;
      let lastCheckinHotel = null;
      let lastCheckoutHotel = null;
      let startRoomNoHotel = currentRowHotel;
      let startCheckinHotel = currentRowHotel;
      let startCheckoutHotel = currentRowHotel;

      for (const row of hotelRows) {
        const values = Object.values(row);
        hotelSheet.addRow(values);

        const isNewRoomNo = row.room_no !== lastRoomNoHotel;
        const isNewCheckin =
          isNewRoomNo || row.checkin_date !== lastCheckinHotel;
        const isNewCheckout =
          isNewCheckin || row.checkout_date !== lastCheckoutHotel;

        // Merge Checkout Date (Column D)
        if (
          lastCheckoutHotel !== null &&
          isNewCheckout &&
          currentRowHotel - startCheckoutHotel > 1
        ) {
          hotelSheet.mergeCells(
            `D${startCheckoutHotel}:D${currentRowHotel - 1}`
          );
        }
        if (isNewCheckout) startCheckoutHotel = currentRowHotel;

        // Merge Checkin Date (Column C)
        if (
          lastCheckinHotel !== null &&
          isNewCheckin &&
          currentRowHotel - startCheckinHotel > 1
        ) {
          hotelSheet.mergeCells(
            `C${startCheckinHotel}:C${currentRowHotel - 1}`
          );
        }
        if (isNewCheckin) startCheckinHotel = currentRowHotel;

        // Merge Room No. (Column B)
        if (
          lastRoomNoHotel !== null &&
          isNewRoomNo &&
          currentRowHotel - startRoomNoHotel > 1
        ) {
          hotelSheet.mergeCells(`B${startRoomNoHotel}:B${currentRowHotel - 1}`);
        }
        if (isNewRoomNo) startRoomNoHotel = currentRowHotel;

        lastRoomNoHotel = row.room_no;
        lastCheckinHotel = row.checkin_date;
        lastCheckoutHotel = row.checkout_date;
        currentRowHotel++;
      }
      // Final merge for the last block in hotel-specific tab
      if (currentRowHotel - startCheckoutHotel > 1)
        hotelSheet.mergeCells(`D${startCheckoutHotel}:D${currentRowHotel - 1}`);
      if (currentRowHotel - startCheckinHotel > 1)
        hotelSheet.mergeCells(`C${startCheckinHotel}:C${currentRowHotel - 1}`);
      if (currentRowHotel - startRoomNoHotel > 1)
        hotelSheet.mergeCells(`B${startRoomNoHotel}:B${currentRowHotel - 1}`);
      console.log(`📊 '${hotelName}' sheet populated and merged.`);
    }

    // Step 6: Save and upload
    const fileName = `room-allocation-${Date.now()}.xlsx`;
    const tempDir = os.tmpdir();
    const filePath = path.join(tempDir, fileName);

    console.log(`📝 Attempting to write Excel file to: ${filePath}`);
    await workbook.xlsx.writeFile(filePath);
    console.log("✅ Excel file written:", filePath);

    const uploadedPath = await uploadReport(
      filePath,
      "room_allocation",
      "xlsx"
    );
    console.log("📤 Uploaded to Supabase Storage:", uploadedPath);

    fs.unlinkSync(filePath); // delete local file
    console.log("🧹 Temporary file cleaned up.");

    return uploadedPath;
  } catch (error) {
    console.error(
      "❌ An error occurred during room allocation report generation:",
      error.message
    );
    console.error(error.stack);
    throw error;
  }
}

module.exports = { generateRoomAllotmentReport };
