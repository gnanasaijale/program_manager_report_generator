const express = require("express");
const cors = require("cors");
const { generatePickupReport } = require("../scripts/generatePickup");
const { generateDropReport } = require("../scripts/generateDrop");
const {
  generateCommonLogisticsReport,
} = require("../scripts/generateCommonLogistics");
const {
  generateRoomAllotmentReport,
} = require("../scripts/generateRoomAllotment");

const app = express();
app.use(
  cors({
    origin: "https://admin.sohamwithin.com", // Your real frontend domain
    methods: ["GET", "POST", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
  })
);

app.get("/generate-report", async (req, res) => {
  const { type, program_id } = req.query;
  if (!program_id) return res.status(400).send("Missing program_id");

  try {
    let result;
    if (type === "pickup") result = await generatePickupReport(program_id);
    else if (type === "drop") result = await generateDropReport(program_id);
    else if (type === "room-allocation")
      result = await generateRoomAllotmentReport(program_id);
    else if (type === "logistics")
      result = await generateCommonLogisticsReport(program_id);
    else return res.status(400).send("Invalid type");

    res.json({ success: true, file_path: result });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () =>
  console.log(`🚀 Report server running on http://localhost:${PORT}`)
);
