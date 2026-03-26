import express from "express";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { readFileSync, existsSync, readdirSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import fetch from "node-fetch";

const __dirname = dirname(fileURLToPath(import.meta.url));
const app = express();
app.use(express.json({ limit: "10mb" }));

// ── Auth middleware ──────────────────────────────────────────────────────────
const QLC_API_KEY = process.env.QLC_API_KEY;
if (!QLC_API_KEY) {
  console.error("FATAL: QLC_API_KEY environment variable is not set.");
  process.exit(1);
}

function requireApiKey(req, res, next) {
  const key = req.headers["x-api-key"];
  if (!key || key !== QLC_API_KEY) {
    return res.status(401).json({ error: "Unauthorized", code: "INVALID_API_KEY" });
  }
  next();
}

// ── Template loader ──────────────────────────────────────────────────────────
// Priorité : fichier local /templates/ → kDrive (si KDRIVE_TOKEN défini)
async function loadTemplate(templateName) {
  const safeName = templateName.replace(/[^a-zA-Z0-9_-]/g, "");
  const localPath = join(__dirname, "templates", `${safeName}.docx`);

  if (existsSync(localPath)) {
    return readFileSync(localPath);
  }

  // Fallback : fetch depuis kDrive
  const token = process.env.KDRIVE_TOKEN;
  const driveId = process.env.KDRIVE_DRIVE_ID;
  const folderPath = process.env.KDRIVE_TEMPLATES_PATH || "/QualityLinks_Compliance/00_SYSTEM/TEMPLATES";

  if (!token || !driveId) {
    throw new Error(`Template "${safeName}" not found locally and KDRIVE_TOKEN/KDRIVE_DRIVE_ID not set.`);
  }

  const url = `https://api.infomaniak.com/2/drive/${driveId}/files/search?query=${encodeURIComponent(safeName + ".docx")}&directory_id=0`;
  const searchRes = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const searchData = await searchRes.json();
  const file = searchData?.data?.files?.[0];
  if (!file) throw new Error(`Template "${safeName}.docx" not found on kDrive.`);

  const dlRes = await fetch(
    `https://api.infomaniak.com/2/drive/${driveId}/files/${file.id}/download`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!dlRes.ok) throw new Error(`kDrive download failed: ${dlRes.status}`);
  return Buffer.from(await dlRes.arrayBuffer());
}

// ── Variable substitution guard ──────────────────────────────────────────────
function checkResidualPlaceholders(doc) {
  const xml = doc.getZip().generate({ type: "string" });
  const match = xml.match(/\{[a-zA-Z_][a-zA-Z0-9_]*\}/);
  if (match) {
    throw { code: "ABORT_F2", message: `Residual placeholder found: ${match[0]}` };
  }
}

// ── POST /generate ────────────────────────────────────────────────────────────
app.post("/generate", requireApiKey, async (req, res) => {
  const { template, variables, options = {} } = req.body;

  // Validation
  if (!template) return res.status(400).json({ error: "Missing field: template", code: "BAD_REQUEST" });
  if (!variables || typeof variables !== "object") return res.status(400).json({ error: "Missing field: variables (object)", code: "BAD_REQUEST" });

  // Guardian F1 : AD + SANTE
  const pays = (variables.PAYS_CODE || "").toUpperCase();
  const niche = (variables.NICHE || "").toUpperCase();
  if (pays === "AD" && niche === "SANTE") {
    return res.status(422).json({ error: "Generation blocked", code: "ABORT_F1", rule: "Guardian F1: AD + SANTE combination is prohibited." });
  }

  try {
    // 1. Load template
    const content = await loadTemplate(template);

    // 2. Docxtemplater render
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      errorLogging: false,
    });
    doc.render(variables);

    // 3. Guardian F2 : residual placeholder check
    if (options.strict !== false) {
      checkResidualPlaceholders(doc);
    }

    // 4. Output
    const output = doc.getZip().generate({ type: "nodebuffer", compression: "DEFLATE" });
    const filename = `${template}_${pays}_${niche}_${Date.now()}.docx`;

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.setHeader("X-QLC-Template", template);
    res.setHeader("X-QLC-Pays", pays);
    res.setHeader("X-QLC-Niche", niche);
    res.send(output);

  } catch (err) {
    if (err.code === "ABORT_F2") {
      return res.status(422).json({ error: err.message, code: "ABORT_F2" });
    }
    if (err.properties?.errors) {
      return res.status(422).json({ error: "Docxtemplater render error", code: "TEMPLATE_ERROR", details: err.properties.errors.map(e => e.message) });
    }
    console.error("[/generate] Error:", err.message);
    res.status(500).json({ error: "Internal server error", code: "INTERNAL_ERROR", detail: err.message });
  }
});

// ── POST /generate-base64 ─────────────────────────────────────────────────────
// Variante : retourne le docx en base64 JSON (utile pour n8n sans binary node)
app.post("/generate-base64", requireApiKey, async (req, res) => {
  const { template, variables, options = {} } = req.body;

  if (!template || !variables) return res.status(400).json({ error: "Missing template or variables", code: "BAD_REQUEST" });

  const pays = (variables.PAYS_CODE || "").toUpperCase();
  const niche = (variables.NICHE || "").toUpperCase();
  if (pays === "AD" && niche === "SANTE") {
    return res.status(422).json({ error: "Generation blocked", code: "ABORT_F1" });
  }

  try {
    const content = await loadTemplate(template);
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, errorLogging: false });
    doc.render(variables);
    if (options.strict !== false) checkResidualPlaceholders(doc);

    const output = doc.getZip().generate({ type: "nodebuffer", compression: "DEFLATE" });
    const filename = `${template}_${pays}_${niche}_${Date.now()}.docx`;

    res.json({ success: true, filename, base64: output.toString("base64"), size: output.length });
  } catch (err) {
    if (err.code === "ABORT_F2") return res.status(422).json({ error: err.message, code: "ABORT_F2" });
    res.status(500).json({ error: err.message, code: "INTERNAL_ERROR" });
  }
});

// ── GET /health ───────────────────────────────────────────────────────────────
app.get("/health", (req, res) => {
  res.json({ status: "ok", service: "qlc-docxtemplater", version: "2.1.0", timestamp: new Date().toISOString() });
});

// ── GET /templates ────────────────────────────────────────────────────────────
app.get("/templates", requireApiKey, (req, res) => {
  try {
    const dir = join(__dirname, "templates");
    const files = existsSync(dir) ? readdirSync(dir).filter(f => f.endsWith(".docx")) : [];
    res.json({ templates: files.map(f => f.replace(".docx", "")), count: files.length });
  } catch {
    res.json({ templates: [], count: 0 });
  }
});

// ── Start ─────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`[QLC] Docxtemplater microservice v2.1.0 running on port ${PORT}`);
  console.log(`[QLC] Endpoints: POST /generate · POST /generate-base64 · GET /health`);
});
