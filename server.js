const fs = require("fs");
const xlsx = require("xlsx");
const nodemailer = require("nodemailer");

// Load your Excel file
async function loadExcelData(fileName, sheetName) {
  try {
    console.log(`Attempting to read Excel file: ${fileName}`);
    const workbook = xlsx.readFile(fileName);

    console.log("Available sheets in the workbook:");
    workbook.SheetNames.forEach((sheet, index) => {
      console.log(`${index + 1}. ${sheet}`);
    });

    if (!sheetName) {
      sheetName = workbook.SheetNames[0];
      console.log(
        `No sheet name provided. Using the first sheet: "${sheetName}"`
      );
    }

    if (!workbook.Sheets[sheetName]) {
      throw new Error(`Sheet "${sheetName}" not found in the workbook`);
    }

    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet);
    console.log(
      `Successfully read ${data.length} rows from sheet "${sheetName}"`
    );
    return data;
  } catch (error) {
    console.error("Error reading Excel file:", error);
    throw error;
  }
}

// Email configuration
function createTransporter() {
  console.log("Creating email transporter...");
  return nodemailer.createTransport({
    host: "smtp.gmail.com",
    port: 465,
    service: "gmail",
    auth: {
      user: "garvitjhalani649@gmail.com",
      pass: "Your_Pass_Here",
    },
  });
}

async function sendEmail(transporter, row) {
  const { Name, Company, Email } = row;
  console.log(`Attempting to send email to: ${Name}, ${Company}, ${Email}`);

  const nameParts = Name.split(" ");
  const name = nameParts[0];
  const mailOptions = {
    from: "Garvit Jhalani <garvitjhalani649@gmail.com>",
    to: Email,
    subject: `Request for an Interview Opportunity - Full Stack Developer at ${Company}`,
    html: `
<p>Greetings ${name},</p>
<p>I'm Garvit Jhalani, a Software Developer with expertise in Full Stack Development. I’m reaching out to express my interest in the Full Stack Development or Generative AI Developer at your company. I would like to introduce myself. I have: 
<ul>
<li><b>Hands-on experience in Full Stack Development</b>, with a strong focus on the MERN stack as well as in Generative AI.</li>
<li>Expertise in <b>JavaScript, React, Node.js, MongoDB, ExpressJS, Tailwind CSS, Python, Langchain, Hugging Face, etc</b>.</li>
<li>Experience with deploying applications on AWS and using modern tools like <b>Auth0</b> for authentication.</li>
<li>Developed projects such as:
  <ul>
    <li><b>Markify:</b> A bookmark application to efficiently organize and manage saved content.</li>
    <li><b>Internshala Automation:</b> A tool designed to automate repetitive processes for improved productivity.</li>
    <li><b>Connectify:</b> A real-time chat application using the MERN stack and Socket.IO.</li>
  </ul>
</li>
<li>Familiarity with <b>REST APIs</b> and creating responsive designs for seamless user experiences.</li>
</ul>

<p>Currently, I am <b>actively exploring new opportunities</b> and can join within a short timeframe if given an offer. I believe my skills and experiences make me a strong fit for the position at ${Company}.</p>

<p>PS: I have attached my <b><a href="https://drive.google.com/file/d/1zTmbNpzAziXRsH6JTcljBdx2yyPkKUF5/view?usp=sharing">Resume</a></b> and <b><a href="https://www.linkedin.com/in/garvit-jhalani/">LinkedIn Profile</a></b>. If you find me suitable, I would greatly appreciate your help with an interview opportunity at ${Company}.</p>

<p>
Thanking You<br>
Regards,<br>
<b>Garvit Jhalani</b><br>
<b>Contact No: +91 6375606887</b><br>
</p>`,
  };

  try {
    const info = await transporter.sendMail(mailOptions);
    console.log(`Email sent to ${Email}. Message ID: ${info.messageId}`);
    return true;
  } catch (error) {
    console.error(`Error sending email to ${Email}:`, error);
    return false;
  }
}

async function sendEmailsSynchronously(data) {
  console.log(`Starting to send emails to ${data.length} recipients...`);
  const transporter = createTransporter();
  let successCount = 0;
  let failureCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    console.log(`Processing email ${i + 1} of ${data.length}`);
    const success = await sendEmail(transporter, row);
    if (success) {
      successCount++;
    } else {
      failureCount++;
    }
    console.log(
      `Current stats - Successes: ${successCount}, Failures: ${failureCount}`
    );
    // Random delay between 60 and 90 seconds
    const delay = 60000 + Math.random() * 30000;
    console.log(
      `Waiting for ${Math.round(delay / 1000)} seconds before next email...`
    );
    await new Promise((resolve) => setTimeout(resolve, delay));
  }

  console.log(
    `Finished sending emails. Final stats - Successes: ${successCount}, Failures: ${failureCount}`
  );
  await transporter.close();
}

async function main() {
  try {
    // const data = await loadExcelData();
    const fileName = process.argv[2] || "./CompanyWiseEmail.xlsx";
    const sheetName = process.argv[3];
    const data = await loadExcelData(fileName, sheetName);
    if (data.length === 0) {
      console.error(
        "No data found in Excel file. Please check the file content."
      );
      return;
    }
    await sendEmailsSynchronously(data);
  } catch (error) {
    console.error("An error occurred in main function:", error);
  }
}

main();

// Example usage:
// node send_emails.js
