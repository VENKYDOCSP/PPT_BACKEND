const express = require("express");
const pdfParse = require("pdf-parse");
const formidable = require("formidable");
const cors = require("cors");
const multer = require("multer");
// const pptxParser = require("pptx-parser");
const path = require("path");
const fs = require("fs/promises");
const axios = require("axios");
const generateGoogleSlidesFromTemplate = require("./generateGoogleSlidesFromTemplate");

const app = express();
app.use(cors({
    origin: "http://localhost:3000",
    methods: ["GET", "POST"],
    allowedHeaders: ["Content-Type"]
}));

const PORT = 5000;

const upload = multer({ dest: "uploads/" });

const processTextWithGemini = async (text) => {
    try {
        const response = await axios.post(
            "https://generativelanguage.googleapis.com/v1/models/gemini-1.5-pro:generateContent?key=AIzaSyDHXz_F-_yiLnPGGYGh-lXi4gHdFftbqMY",
            {
                contents: [{
                    parts: [{
                        text: `Analyze this text and return a structured JSON for PowerPoint slides. 
                               - Each slide must have a "slideNumber".
                               - Each slide must have a "title".
                               - Each slide may have a "subtitle" (if applicable).
                               - Each slide must have "content" as bullet points.
                               - Each Point should have atleast a 10 - 15 words.
                               - Ensure content fits within a slide (max 300 characters).
                               - The number of slides should be appropriate for the given text.

                               Input Text:\n\n${text}`
                    }]
                }]
            },
            { headers: { "Content-Type": "application/json" } }
        );

        console.log("Response==>", response.data);

        if (response.data.candidates && response.data.candidates.length > 0) {
            let rawText = response.data.candidates[0].content.parts[0].text;

            rawText = rawText.replace(/^```json\s*/, '').replace(/```$/, '');

            let structuredSlides = JSON.parse(rawText);

            structuredSlides = structuredSlides.map((slide, index) => ({
                slideNumber: index + 1,
                ...slide
            }));

            return structuredSlides;
        } else {
            throw new Error("Invalid response from Gemini API");
        }
    } catch (error) {
        console.error("Error processing text with Gemini:", error);
        return { error: "Failed to generate slides" };
    }
};

let structuredContent = [];

app.post("/extract-pdf", async (req, res) => {
    try {

        const form = new formidable.IncomingForm();
        form.parse(req, async (err, fields, files) => {
            if (err) {
                console.error("Error in line 78 ===>", err);
                return res.status(500).json({ error: "Failed to parse form data" });
            }

            if (!files.pdf || files.pdf.length === 0) {
                return res.status(400).json({ error: "No file uploaded" });
            }

            const fileBuffer = await fs.readFile(files.pdf[0].filepath);
            const data = await pdfParse(fileBuffer);
            console.log("Extracted text:", data.text);

            const structuredResult = await processTextWithGemini(data.text);

            structuredContent.push(structuredResult);

            res.status(200).json({
                extractedText: data.text,
                structuredContent: structuredResult
            });
        });
    } catch (error) {
        console.error("Error in Line 100 ===>", error);
        res.status(500).json({ error: "Failed to process PDF" });
    }
});


app.post("/upload-template-and-generate", upload.single("ppt"), async (req, res) => {
    try {
        const pptPath = req.file.path;


        const slideUrl = await generateGoogleSlidesFromTemplate(pptPath, structuredContent);


        // fs.unlinkSync(pptPath); 

        console.log("Google Slides URL", slideUrl);
        res.json({ presentationUrl: slideUrl });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: "Failed to generate presentation" });
    }
});


app.listen(PORT, () => {
    console.log(`Server is running on ${PORT}`);
});
