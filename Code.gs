const CONFIG = {
  projectName: "Customer Feedback",
  folderName: "Pollution Images",
  sheetName: "Feedback Responses",
  maxImages: 5,
  maxImageSize: 2 * 1024 * 1024
};

// Run this function once to create Drive folder and Sheet.
function setupProject() {
  const props = PropertiesService.getScriptProperties();

  let mainFolderId = props.getProperty("MAIN_FOLDER_ID");
  let spreadsheetId = props.getProperty("SPREADSHEET_ID");

  if (!mainFolderId) {
    const folder = DriveApp.createFolder(CONFIG.projectName);
    mainFolderId = folder.getId();
    props.setProperty("MAIN_FOLDER_ID", mainFolderId);
  }

  const mainFolder = DriveApp.getFolderById(mainFolderId);

  let imageFolderId = props.getProperty("IMAGE_FOLDER_ID");

  if (!imageFolderId) {
    const imageFolder = mainFolder.createFolder(CONFIG.folderName);
    imageFolderId = imageFolder.getId();
    props.setProperty("IMAGE_FOLDER_ID", imageFolderId);
  }

  if (!spreadsheetId) {
    const ss = SpreadsheetApp.create("Customer Feedback Records");
    spreadsheetId = ss.getId();
    props.setProperty("SPREADSHEET_ID", spreadsheetId);

    const sheet = ss.getActiveSheet();
    sheet.setName(CONFIG.sheetName);

    sheet.appendRow([
      "Feedback ID",
      "Submission Date",
      "Customer Name",
      "Mobile Number",
      "Feedback",
      "Image Links"
    ]);

    sheet.setFrozenRows(1);
  }

  Logger.log("Project setup completed!");
  Logger.log("Drive Folder: " + mainFolder.getUrl());
  Logger.log("Spreadsheet: " +
    SpreadsheetApp.openById(spreadsheetId).getUrl());
  Logger.log("Image Folder: " +
    DriveApp.getFolderById(imageFolderId).getUrl());
}

// Display customer feedback form.
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Customer Feedback - Pollution Reporting")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// Receive and save customer feedback.
function submitFeedback(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    if (!data || typeof data !== "object") {
      throw new Error("Invalid submission.");
    }

    const name = String(data.name || "").trim();
    const mobile = String(data.mobile || "").trim();
    const feedback = String(data.feedback || "").trim();
    const images = data.images || [];

    if (!name || !mobile || !feedback) {
      throw new Error("Please fill in all required fields.");
    }

    if (!/^\d{10}$/.test(mobile)) {
      throw new Error("Enter a valid 10-digit mobile number.");
    }

    if (!Array.isArray(images) || images.length > CONFIG.maxImages) {
      throw new Error("You can upload a maximum of 5 images.");
    }

    const props = PropertiesService.getScriptProperties();
    const imageFolder = DriveApp.getFolderById(
      props.getProperty("IMAGE_FOLDER_ID")
    );

    const spreadsheet = SpreadsheetApp.openById(
      props.getProperty("SPREADSHEET_ID")
    );

    const sheet = spreadsheet.getSheetByName(CONFIG.sheetName);

    const feedbackId = "CF-" + Utilities.getUuid().slice(0, 8).toUpperCase();
    const imageLinks = [];

    images.forEach((image, index) => {
      const mimeType = String(image.mimeType || "").toLowerCase();

      const allowedTypes = [
        "image/jpeg",
        "image/png",
        "image/webp"
      ];

      if (!allowedTypes.includes(mimeType)) {
        throw new Error("Only JPG, PNG and WEBP images are allowed.");
      }

      if (!image.data || typeof image.data !== "string") {
        throw new Error("Invalid image data.");
      }

      const bytes = Utilities.base64Decode(image.data);

      if (bytes.length > CONFIG.maxImageSize) {
        throw new Error("Each image must be smaller than 2 MB.");
      }

      const extension = {
        "image/jpeg": "jpg",
        "image/png": "png",
        "image/webp": "webp"
      }[mimeType];

      const blob = Utilities.newBlob(
        bytes,
        mimeType,
        feedbackId + "_Image_" + (index + 1) + "." + extension
      );

      const file = imageFolder.createFile(blob);
      imageLinks.push(file.getUrl());
    });

    // Prevent spreadsheet formula injection.
    const safeCell = value =>
      /^[=+\-@]/.test(String(value)) ? "'" + value : value;

    sheet.appendRow([
      feedbackId,
      new Date(),
      safeCell(name),
      safeCell(mobile),
      safeCell(feedback),
      imageLinks.join("\n")
    ]);

    return {
      success: true,
      feedbackId: feedbackId,
      message: "Your feedback has been submitted successfully!"
    };

  } catch (error) {
    throw new Error(error.message);
  } finally {
    lock.releaseLock();
  }
}
