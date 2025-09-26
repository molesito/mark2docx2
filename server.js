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

    const formulasZip = new AdmZip(req.files["formulas"][0].buffer);
    const outputZip = new AdmZip(req.files["output"][0].buffer);

    // Leer XML
    const formulasXml = formulasZip.readAsText("word/document.xml");
    const outputXml = outputZip.readAsText("word/document.xml");

    // Construir mapa [{x}]...[{x}]
    const blockMap = {};
    const regex = /\[\{(\d+)\}\]([\s\S]*?)\[\{\1\}\]/g;
    let match;
    while ((match = regex.exec(formulasXml)) !== null) {
      blockMap[match[1]] = match[2]; // Guardamos XML crudo
    }

   // Reemplazar en output (con runs de espacio válidos)
let mergedXml = outputXml;
for (const [id, fragment] of Object.entries(blockMap)) {
  const placeholder = `[{${id}}]`;
  const regex = new RegExp(placeholder, "g");

  mergedXml = mergedXml.replace(regex, (match, offset) => {
    const before = mergedXml[offset - 1] || "";
    const after = mergedXml[offset + match.length] || "";

    // Run XML de un espacio
    const spaceRun = '<w:r><w:t xml:space="preserve"> </w:t></w:r>';

    const addBefore = before !== " " && before !== "\n" ? spaceRun : "";
    const addAfter = after !== " " && after !== "\n" ? spaceRun : "";

    return `${addBefore}${fragment}${addAfter}`;
  });
}


    // Crear nuevo docx
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
  }
);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));


