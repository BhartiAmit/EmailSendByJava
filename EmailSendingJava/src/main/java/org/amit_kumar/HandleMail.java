package org.amit_kumar;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import java.io.*;
import java.util.*;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

public class HandleMail {

    public void sendMail() throws IOException {

        String host = "smtp.gmail.com";
        Properties props = System.getProperties();
        props.put("mail.smtp.host", host);
        props.put("mail.smtp.port", "465");
        props.put("mail.smtp.ssl.enable", "true");
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.connectiontimeout", "1000000"); // 10s to connect
        props.put("mail.smtp.timeout", "1000000");           // 10s for read/write
        props.put("mail.smtp.writetimeout", "1000000");


        Session mailSession = Session.getInstance(props, new MailAuthenticator());

        final String HRContact_PATH = "C:\\Users\\52361838\\Documents\\My Documents\\Hr Details Company Wise.xlsx";

        FileInputStream fis = new FileInputStream(new File(HRContact_PATH));
        Workbook workbook = WorkbookFactory.create(fis);
        Sheet sheet = workbook.getSheetAt(0); // or workbook.getSheet("Contacts");

        for (int i = 1; i <=2/*sheet.getLastRowNum()*/; i++) {

            Row row = sheet.getRow(i);
            if (row == null) continue;

            String receiverName = row.getCell(1).getStringCellValue();
            String email = row.getCell(2).getStringCellValue();
            String title = row.getCell(3).getStringCellValue();
            String company = row.getCell(4).getStringCellValue();

            try {
                MimeMessage mimeMessage = new MimeMessage(mailSession);

                mimeMessage.setFrom(new InternetAddress(MailConstants.SENDER));
                mimeMessage.setRecipient(Message.RecipientType.TO, new InternetAddress(email));
                mimeMessage.setSubject(MailConstants.SUBJECT);

                // Mail body part
                MimeBodyPart messageBodyPart = new MimeBodyPart();
                messageBodyPart.setContent(buildDynamicMessage(receiverName, title, company), "text/html; charset=utf-8");

                // Attachment part
                MimeBodyPart attachmentPart = new MimeBodyPart();
                attachmentPart.attachFile(new File(MailConstants.RESUME_PATH));

                // Multipart for combining both
                Multipart multipart = new MimeMultipart();
                multipart.addBodyPart(messageBodyPart);
                multipart.addBodyPart(attachmentPart);

                mimeMessage.setContent(multipart);

                Transport.send(mimeMessage);
                Thread.sleep(2000);
                Cell statusCell = row.createCell(5); // Column E (index 5) for Status
                if (statusCell == null) {
                    statusCell = row.createCell(5);
                }
                statusCell.setCellValue("Mail sent");
                System.out.println(i+" Mail sent successfully to " + email);
            } catch (Exception e) {
                Cell statusCell = row.createCell(5); // Column E (index 5) for Status
                if (statusCell == null) {
                    statusCell = row.createCell(5);
                }
                statusCell.setCellValue("Mail Failed");
                System.out.println(i+" Failed to send mail to " + email);
                e.printStackTrace();
            }
        }
        fis.close(); // Close the FileInputStream (input file)

        FileOutputStream fos = new FileOutputStream("C:\\Users\\52361838\\Documents\\My Documents\\HR_Contacts_Status_Updated.xlsx");
        workbook.write(fos);
        fos.close();

        workbook.close();

    }

    private String buildDynamicMessage(String name, String title, String company) {
        return "<p>Dear <b>" + name + "</b> / " + title + " at <b>" + company + "</b>,</p>" +

                "<p>I hope you are having a great day.</p>" +

                "<p>My name is <b>Amit Kumar</b>, and I am reaching out to express my interest in the <b>Java Developer</b> role at your organization. " + "With <b>3 years</b> of experience as a <b>Java Developer</b>, I have worked extensively with technologies like " + "<b>Java</b>, <b>Spring Boot</b>, <b>Hibernate</b>, <b>RESTful APIs</b>, and <b>SQL</b>.</p>" +

                "<p>Currently, I am employed with Conduent Business Services India LLP, Noida.</p>" +

                "<p>I have 3 years of experience in Requirement Analysis, Development, Integration, and Support of enterprise applications using Java Technologies. </p>"+

                "<p>My core technical competencies include Java, Spring MVC, Spring Boot, REST API. I am also experienced in working with MySQL, Oracle, and MS Access databases, along with tools such as Eclipse/STS, Maven, Postman, Git, and more.</p>"+

                "<p>I believe my technical expertise could be a valuable asset to your team</p>"+

                "<p>Please find my resume attached for your reference. I would be happy to discuss how my background aligns with your team's goals.</p>" +

                "<p>Thank you for your time and consideration & I look forward to the opportunity to connect.</p>" +

                "<p>Best Regards,<br>" + "Amit Kumar<br>" + "📧 amitkumarbharti783@gmail.com<br>" + "📞 +91-7800389713<br>" + "🔗 <a href=\"https://www.linkedin.com/in/amit-kumar-84b37121b/\">LinkedIn Profile</a></p>";
    }

}

