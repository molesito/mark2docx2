import express from "express";
import multer from "multer";
import AdmZip from "adm-zip";

const upload = multer();
const app = express();

app.get("/health", (req, res) => res.json({ status: "ok" }));

app.post(
  "/merge-docx",
  upload.fields([{ name: "formulas" }, { name: "output" }]),
  (req, res) => {
    if (!req.files["formulas"] || !req.files["output"]) {
      return res.status(400).json({ error: "Missing files" });
    }

    try {
      const formulasZip = new AdmZip(req.files["formulas"][0].buffer);
      const outputZip = new AdmZip(req.files["output"][0].buffer);

      // Leer XML principal
      const formulasXml = formulasZip.readAsText("word/document.xml");
      const outputXml = outputZip.readAsText("word/document.xml");

      // 1) Construir mapa de bloques [{x}] ... [{x}]
      const blockMap = {};
      const regex = /\[\{(\d+)\}\]([\s\S]*?)\[\{\1\}\]/g;
      let match;
      while ((match = regex.exec(formulasXml)) !== null) {
        blockMap[match[1]] = match[2]; // Guardamos el fragmento XML crudo
      }

      // 2) Sustituir placeholders en output.docx
      let mergedXml = outputXml;
      for (const [id, fragment] of Object.entries(blockMap)) {
        const placeholder = `[{${id}}]`;
        const regexPlaceholder = new RegExp(escapeRegex(placeholder), "g");
        mergedXml = mergedXml.replace(regexPlaceholder, fragment);
      }

      // 3) Forzar que Word preserve los espacios en todos los <w:t>
      mergedXml = mergedXml.replace(
        /<w:t(?![^>]*xml:space)/g,
        '<w:t xml:space="preserve"'
      );

      // 4) Reempaquetar DOCX
      const mergedZip = new AdmZip();
      for (const entry of outputZip.getEntries()) {
        if (entry.entryName === "word/document.xml") {
          mergedZip.addFile(entry.entryName, Buffer.from(mergedXml, "utf8"));
        } else {
          mergedZip.addFile(entry.entryName, entry.getData());
        }
      }

      const filename = req.body.filename || "merged.docx";
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      );
      res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
      res.send(mergedZip.toBuffer());
    } catch (err) {
      console.error("Merge error:", err);
      return res.status(500).json({ error: "Failed to merge documents" });
    }
  }
);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));

// Helper para escapar caracteres especiales en el placeholder
function escapeRegex(string) {
  return string.replace(/[-\/\\^$*+?.()|[\]{}]/g, "\\$&");
}
