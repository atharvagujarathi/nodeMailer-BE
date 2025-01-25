const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const nodemailer = require("nodemailer");
const path = require("path");
const cors = require("cors");

const app = express();
const PORT = 3000;

// Middleware
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(
  cors({
    origin: "http://localhost:4200", // Replace with your Angular app's URL
  })
);

// Multer Configuration for File Uploads
const upload = multer({ dest: "uploads/" });

// Nodemailer Configuration
const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: "Your mail",
    pass: "Your app pass",
  },
});

// Function to Send Emails
const sendEmail = async (to, subject, body) => {
  if (!to) throw new Error("Recipient email is undefined");
  return transporter.sendMail({
    from: "Your Name",
    to,
    subject,
    html: body,
  });
};
let shouldStop = false;

app.post("/cancel", (req, res) => {
  shouldStop = true;
  res.status(200).send("Email sending canceled.");
});

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// API to Handle File Uploads and Email Sending
let isSendingEmails = false; // Flag to track email sending status

app.post("/upload", upload.single("file"), async (req, res) => {
  if (isSendingEmails) {
    return res
      .status(400)
      .json({ message: "Email sending is already in progress. Please wait." });
  }

  try {
    const subject = req.body.subject;
    const body = req.body.body;
    const filePath = req.file.path;

    isSendingEmails = true; // Set flag to true before starting the email sending process
    shouldStop = false; // Reset stop flag

    // Read Excel File
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet);

    // Extract Emails
    const emails = data
      .map((row) => row.Email)
      .filter((email) => email && email.trim());

    if (!emails.length) {
      isSendingEmails = false; // Reset the flag in case of no valid emails
      return res.status(400).json({
        message: "No valid email addresses found in the uploaded file.",
      });
    }

    // Return Email List to Frontend
    res.status(200).json({ emails });

    // Send Emails Asynchronously
    for (let i = 0; i < emails.length; i++) {
      if (shouldStop) {
        console.log("Email sending canceled.");
        break;
      }
      const email = emails[i];
      try {
        console.log(`Waiting to send email to ${email}`);
        await delay(30 * 1000); // Delay of 90 seconds
        if (shouldStop) {
          console.log("Email sending canceled after delay.");
          break; // Check again after delay
        }
        console.log(`Sending email to: ${email}`);
        await sendEmail(email, subject, body);
        console.log(`Email successfully sent to ${email}`);
      } catch (error) {
        console.error(`Failed to send email to ${email}:`, error.message);
      }
    }

    isSendingEmails = false; // Reset flag after all emails are sent or canceled
  } catch (error) {
    console.error("Error processing the file:", error);
    isSendingEmails = false; // Reset flag in case of error
    res.status(500).send("An error occurred while processing the file.");
  }
});

// Endpoint to get emails from the uploaded Excel file
app.get("/users", async (req, res) => {
  try {
    // Check if a file has been uploaded
    const filePath = req.query.filePath; // Assuming the file path is passed as a query parameter

    if (!filePath) {
      return res.status(400).json({ message: "No file uploaded yet." });
    }

    // Read the uploaded Excel file
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet);

    // Extract email addresses
    const emails = data
      .map((row) => row.Email)
      .filter((email) => email && email.trim());

    if (emails.length === 0) {
      return res.status(400).json({
        message: "No valid email addresses found in the uploaded file.",
      });
    }

    // Return the emails in the response
    res.status(200).json({ emails });
  } catch (error) {
    console.error("Error fetching emails:", error);
    res.status(500).send("An error occurred while reading the file.");
  }
});

// Start the Server
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
