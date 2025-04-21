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
        requestBody: {
            name: "Generated PPT",
        },
    });

    const presentationId = copied.data.id;

    await drive.permissions.create({
        fileId: presentationId,
        requestBody: {
            role: "reader",
            type: "anyone",
        },
    });

    const presentation = await slides.presentations.get({ presentationId });
    const firstSlideId = presentation?.data?.slides?.[0]?.objectId;

    if (!firstSlideId) {
       console.log("Slide not found");
    }

    const requests = [];
    const slideIds = [];

    structuredContent?.[0]?.forEach((slide, index) => {
        const newSlideId = `generated_slide_${index}`;
        slideIds.push(newSlideId);


        requests.push({
            duplicateObject: {
                objectId: firstSlideId,
                objectIds: {
                    [firstSlideId]: newSlideId,
                },
            },
        });
    });

    await slides.presentations?.batchUpdate({
        presentationId,
        requestBody: { requests },
    });

    for (let i = 0; i < structuredContent?.[0]?.length; i++) {
        const slide = structuredContent?.[0][i];
        const slideId = slideIds[i];

        const updateRequests = [];


        updateRequests.push({
            replaceAllText: {
                containsText: {
                    text: "{{title}}",
                    matchCase: true,
                },
                replaceText: slide?.title,
                pageObjectIds: [slideId]
            },
        });


        updateRequests.push({
            replaceAllText: {
                containsText: {
                    text: "{{content}}",
                    matchCase: true,
                },
                replaceText: slide?.content?.map(line => `• ${line}`).join("\n"),
                pageObjectIds: [slideId]
            },
        });

        await slides?.presentations?.batchUpdate({
            presentationId,
            requestBody: { requests: updateRequests },
        });
    }

    
    await slides?.presentations?.batchUpdate({
        presentationId,
        requestBody: {
            requests: [{
                deleteObject: {
                    objectId: firstSlideId
                }
            }]
        }
    });

    await drive.files.delete({ fileId: uploadedFileId });

    try {
        if (fs.existsSync(pptFilePath)) {
            fs.unlinkSync(pptFilePath);
        }
    } catch (err) {
        console.error("Failed to delete local file", err.message);
    }

    return `https://docs.google.com/presentation/d/${presentationId}/edit`;
};

module.exports = generateGoogleSlidesFromTemplate;