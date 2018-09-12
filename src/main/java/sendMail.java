import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.*;
import java.util.Properties;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.FileReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;

import javax.json.Json;
import javax.json.JsonObject;
import javax.json.JsonReader;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class sendMail {
    public static void main(String[] args) throws InvalidFormatException, IOException,ParseException {

        String path = System.getProperty("user.dir");

        File jsonInputFile = new File(path + "/config.json");
        InputStream is;
        is = new FileInputStream(jsonInputFile);
        JsonReader reader = Json.createReader(is);
        JsonObject JsonObj = reader.readObject();
        reader.close();

        InputStream inp = new FileInputStream(path + "/Inputs/" + JsonObj.getString("NameListFile"));
        int ctr = 1;
        Workbook wb = WorkbookFactory.create(inp);
        Sheet sheet = wb.getSheetAt(JsonObj.getInt("SheetPositionwithinFile"));
        Row row = null;
        Cell cell = null;
        String toRecipient = null;
        boolean isNull = false;
        do{
            try{
                row = sheet.getRow(ctr);
                cell = row.getCell(JsonObj.getInt("IDPositionwithinSheet"));
                toRecipient = cell.toString();
                sendPDFReportByGMail(
                        JsonObj.getString("EmailFrom"),
                        JsonObj.getString("Password"),
                        toRecipient,
                        "Welcome!!!",
                        "Dear,\n" +
                                "\n" +
                                "Greetings!!! Hope you are doing well.\n\n\n" +
                                "Best, \n" +
                                "-Sandipan\n" +
                                "Phone: +xxx xxx xxx\n" +
                                "Web: http://xxxx.com/",
                        path,
                        JsonObj.getString("NameAttachFile"));
                ctr++;
            } catch(Exception e) {
                isNull = true;
            }

        }while(isNull!=true);
        inp.close();
    }
    private static void addAttachment(Multipart multipart, String filename) throws MessagingException {
        DataSource source = new FileDataSource(filename);
        BodyPart messageBodyPart = new MimeBodyPart();
        messageBodyPart.setDataHandler(new DataHandler(source));
        messageBodyPart.setFileName(source.getName());
        multipart.addBodyPart(messageBodyPart);

    }
    private static void sendPDFReportByGMail(String from, String pass, String to, String subject, String body,
                                             String path, String AttachmentFileName)
    {
        Properties props = System.getProperties();
        String host = "smtp.gmail.com";
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.smtp.host", host);
        props.put("mail.smtp.user", from);
        props.put("mail.smtp.password", pass);
        props.put("mail.smtp.port", "465");
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.socketFactory.port", "465");
        props.put("mail.smtp.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
        Session session = Session.getDefaultInstance(props);
        MimeMessage message = new MimeMessage(session);
        try
        {     //Set from address message.setFrom(new InternetAddress(from));
            message.addRecipient(Message.RecipientType.TO, new InternetAddress(to));
            message.setSubject(subject);
            message.setText(body);
            BodyPart objMessageBodyPart = new MimeBodyPart();
            objMessageBodyPart.setText(body);
            Multipart multipart = new MimeMultipart();
            multipart.addBodyPart(objMessageBodyPart);
            objMessageBodyPart = new MimeBodyPart();
            //Set path to the pdf report file
            String file = path + "/Inputs/" + AttachmentFileName;
            //Create data source to attach the file in mail
//            String file2 = path + "/Inputs/welcome-image1.jpg";

            addAttachment(multipart, file);
            message.setContent(multipart);
            Transport transport = session.getTransport("smtp");
            transport.connect(host, from, pass);
            transport.sendMessage(message, message.getAllRecipients());
            transport.close();
//          addAttachment(multipart, file2);
        }
        catch (AddressException ae)
        {
            ae.printStackTrace();
        }
        catch (MessagingException me)
        {
            me.printStackTrace();
        }
    }
}