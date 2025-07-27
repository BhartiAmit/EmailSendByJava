# 📧 HR Email Automation Tool (Java)

This project automates the process of sending personalized emails with attachments (like resumes) to HR contacts listed in an Excel sheet, and updates the same Excel with email sent status.

## 🚀 Features

- Reads HR details (Name, Email, Title, Company) from an Excel file
- Sends customized emails using JavaMail (Gmail SMTP)
- Attaches resume automatically
- Updates Excel sheet with "Mail sent" or "Mail failed" status
- Skips rows with missing email addresses or invalid data
- Uses HTML-formatted email body with LinkedIn/contact info

---

## 🛠 Technologies Used

- Java 8+
- JavaMail API
- Apache POI (for Excel)
- Apache PDFBox (if needed for PDF reading)
- Gmail SMTP (with App Password authentication)

---

## 📦 Folder Structure

├── src/
│ ├── org.amit_kumar/
│ │ ├── Main.java
│ │ ├── HandleMail.java
│ │ ├── MailAuthenticator.java
│ │ ├── MailConstants.java
│
├── input/
│ └── Hr Details Company Wise.xlsx # Input Excel
│
├── output/
│ └── HR_Contacts_Status_Updated.xlsx # Output Excel with status
│
├── resources/
│ └── AMIT_KUMAR_CV.pdf # Resume attachment


---

## ✅ Prerequisites

- Java JDK installed (Java 8+)
- Maven or proper build setup
- Internet connection
- Gmail account with:
  - 2-Step Verification **enabled**
  - A generated **App Password**

---

## 🔐 Setting Up App Password (Gmail)

1. Enable [2-Step Verification](https://myaccount.google.com/security)
2. Go to [App Passwords](https://myaccount.google.com/apppasswords)
3. Generate a password for “Mail” on “Windows Computer”
4. Use the 16-character password in `MailAuthenticator.java`:
   ```java
   return new PasswordAuthentication("your_email@gmail.com", "your_app_password");
