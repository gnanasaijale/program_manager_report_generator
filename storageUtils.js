const { supabase } = require("./supabaseClient");
const fs = require("fs");
const path = require("path");

async function uploadReport(filePath, reportType, format) {
  console.log("📤 Starting uploadReport with:", {
    filePath,
    reportType,
    format,
  });
  const fileBuffer = fs.readFileSync(filePath);
  const fileName = path.basename(filePath);
  const storagePath = `${reportType}/${Date.now()}-${fileName}`;

  console.log("📤 Uploading to Supabase Storage at:", storagePath);

  const { data, error: uploadError } = await supabase.storage
    .from("soham-reports")
    .upload(storagePath, fileBuffer, {
      contentType:
        format === "pdf"
          ? "application/pdf"
          : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      upsert: true,
    });

  if (uploadError) {
    console.error("❌ Upload to storage failed:", uploadError);
    throw uploadError;
  }

  console.log("✅ Upload successful:", data);

  const { error: insertError } = await supabase.from("report_files").insert([
    {
      file_name: fileName,
      file_path: storagePath,
      report_type: reportType,
      format,
    },
  ]);

  if (insertError) {
    console.error("❌ Insert into report_files failed:", insertError);
    throw insertError;
  }

  console.log("📝 Report logged in table: report_files");
  return storagePath;
}

module.exports = { uploadReport };
