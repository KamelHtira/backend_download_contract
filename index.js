const express = require("express");
const router = require("./routes");
const cors = require("cors");

const app = express();
const PORT = process.env.PORT || 3005;

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(cors({ origin: "*" }));

app.use("/api", router);

app.use(express.json());

//-----------------------------------------------------------
var convertapi = require("convertapi")("5IG6CUKwhGAnuNHD");
const axios = require("axios");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");

async function generateAndSaveModifiedDocx(templateUrl, jsonData) {
  try {
    // Fetch the DOCX template from the provided URL
    const response = await axios.get(templateUrl, {
      responseType: "arraybuffer",
    });
    const templateBuffer = Buffer.from(response.data);

    // Load the template into docxtemplater
    const zip = new PizZip(templateBuffer);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    // Create a data object for replacements
    const data = {};
    jsonData.forEach((item) => {
      data[item.id] = item.value;
    });

    // Perform replacements in the template
    doc.setData(data);
    doc.render();

    // Generate the modified DOCX file
    const generatedBuffer = doc
      .getZip()
      .generate({ type: "nodebuffer", compression: "DEFLATE" });

    // Save the modified DOCX file
    fs.writeFileSync(path.resolve(__dirname, "output.docx"), generatedBuffer);

    convertapi
      .convert(
        "pdf",
        {
          File: "./output.docx",
        },
        "docx"
      )
      .then(function (result) {
        result.saveFiles("./resources");
      });

    console.log("Modified DOCX file generated successfully.");
  } catch (error) {
    console.error("An error occurred:", error.message);
  }
}

// Example usage
const templateUrl =
  "https://res.cloudinary.com/dh88lf62x/raw/upload/v1691405833/10.Vente_Achat_FR_hmvixs_3_dva8re_3_ozh1qe_rj2sr6_1_kcxwf8.docx";
const jsonData = [
  { id: "87", value: "John" },
  { id: "88", value: "Doe" },
  { id: "89", value: "0652455478" },
  { id: "90", value: "New Website" },
];

generateAndSaveModifiedDocx(templateUrl, jsonData);

//-----------------------------------------------------------
app.get("/", (req, res) => {
  res.send({ msg: "etafakna download contract web server working.." });
});

app.listen(PORT, function () {
  console.log("listening on port 3007!");
});
