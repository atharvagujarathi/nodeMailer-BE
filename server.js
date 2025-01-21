const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const nodemailer = require("nodemailer");
const path = require("path");

const app = express();
const PORT = 3000;

// Middleware
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Multer Configuration for File Uploads
const upload = multer({ dest: "uploads/" });

// Nodemailer Configuration
const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: "atharvagujarathi92@gmail.com",
    pass: "xflo skza hlwv hbhv",
  },
});

// Function to Send Emails
const sendEmail = async (to, subject, body) => {
  if (!to) throw new Error("Recipient email is undefined");
  return transporter.sendMail({
    from: "atharvagujarathi92@gmail.com",
    to,
    subject,
    html: body,
  });
};

// API to Handle File Uploads and Email Sending
app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    const subject = req.body.subject;
    const body = req.body.body;
    const filePath = req.file.path;

    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet);

    const emails = data
      .map((row) => row.Email)
      .filter((email) => email && email.trim());

    if (!emails.length) {
      return res
        .status(400)
        .send("No valid email addresses found in the uploaded file.");
    }

    console.log("Emails to be sent:", emails);

    for (let i = 0; i < emails.length; i++) {
      const email = emails[i];
      try {
        console.log(`Sending email to: ${email}`);
        await new Promise((resolve) => setTimeout(resolve, 90 * 1000)); // Wait 90 seconds
        await sendEmail(email, subject, body);
        console.log(`Email successfully sent to ${email}`);
      } catch (error) {
        console.error(`Failed to send email to ${email}:`, error.message);
        // Handle email-specific errors and continue with the next email
      }
    }

    res.status(200).send("Emails are being sent one by one.");
  } catch (error) {
    console.error("Error processing the file:", error);
    res.status(500).send("An error occurred while processing the file.");
  }
});

// Start the Server
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
