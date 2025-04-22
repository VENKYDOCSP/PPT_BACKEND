const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");

const credential = path.join(__dirname, "cred.json");

const auth = new google.auth.GoogleAuth({
    keyFile: credential,
    scopes: [
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/presentations",
    ],
});

const generateGoogleSlidesFromTemplate = async (pptFilePath, structuredContent) => {
    const authClient = await auth.getClient();
    const drive = google.drive({ version: "v3", auth: authClient });
    const slides = google.slides({ version: "v1", auth: authClient });

    const fileMetadata = {
        name: "Uploaded Template",
        mimeType: "application/vnd.google-apps.presentation",
    };

    const media = {
        mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        body: fs.createReadStream(pptFilePath),
    };

    const uploaded = await drive.files.create({
        resource: fileMetadata,
        media,
        fields: "id",
    });

    const uploadedFileId = uploaded.data.id;

    const copied = await drive.files.copy({
        fileId: uploadedFileId,
        requestBody: { name: "Generated PPT" },
    });

    const presentationId = copied.data.id;

    await drive.permissions.create({
        fileId: presentationId,
        requestBody: { role: "reader", type: "anyone" },
    });

    const presentation = await slides.presentations.get({ presentationId });
    const allSlides = presentation.data.slides;

    const slidesToUpdate = allSlides.slice(0, structuredContent.length);

    const updateSlideContent = async (slide, slideData, index) => {
        const slideId = slide.objectId;


        const deleteRequests = slide.pageElements.map(el => ({
            deleteObject: { objectId: el.objectId },
        }));

        await slides.presentations.batchUpdate({
            presentationId,
            requestBody: { requests: deleteRequests },
        });


        const titleBoxId = `title_box_${index}`;
        const contentBoxId = `content_box_${index}`;

        const requests = [
            {
                createShape: {
                    objectId: titleBoxId,
                    shapeType: "TEXT_BOX",
                    elementProperties: {
                        pageObjectId: slideId,
                        size: { height: { magnitude: 80, unit: "PT" }, width: { magnitude: 600, unit: "PT" } },
                        transform: {
                            scaleX: 1, scaleY: 1, translateX: 50, translateY: 50, unit: "PT",
                        },
                    },
                },
            },
            {
                insertText: { objectId: titleBoxId, text: slideData.title },
            },
            {
                updateTextStyle: {
                    objectId: titleBoxId,
                    style: { fontSize: { magnitude: 24, unit: "PT" }, bold: true },
                    fields: "fontSize,bold",
                },
            },
            {
                createShape: {
                    objectId: contentBoxId,
                    shapeType: "TEXT_BOX",
                    elementProperties: {
                        pageObjectId: slideId,
                        size: { height: { magnitude: 300, unit: "PT" }, width: { magnitude: 600, unit: "PT" } },
                        transform: {
                            scaleX: 1, scaleY: 1, translateX: 50, translateY: 150, unit: "PT",
                        },
                    },
                },
            },
            {
                insertText: {
                    objectId: contentBoxId,
                    text: slideData.content.map(line => `• ${line}`).join("\n"),
                },
            },
        ];

        await slides.presentations.batchUpdate({
            presentationId,
            requestBody: { requests },
        });
    };

    for (let i = 0; i < slidesToUpdate.length; i++) {
        await updateSlideContent(slidesToUpdate[i], structuredContent[i], i);
    }

    const extraSlides = allSlides.slice(structuredContent.length);
    if (extraSlides.length > 0) {
        const deleteRequests = extraSlides.map(s => ({
            deleteObject: { objectId: s.objectId },
        }));
        await slides.presentations.batchUpdate({
            presentationId,
            requestBody: { requests: deleteRequests },
        });
    }

    await drive.files.delete({ fileId: uploadedFileId });

    try {
        if (fs.existsSync(pptFilePath)) fs.unlinkSync(pptFilePath);
    } catch (err) {
        console.error("Failed to delete local file:", err.message);
    }

    return `https://docs.google.com/presentation/d/${presentationId}/edit`;
};

module.exports = generateGoogleSlidesFromTemplate;
