package com.openDoc.testOpenWord;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import jakarta.mail.*;
import jakarta.mail.Authenticator;
import jakarta.mail.PasswordAuthentication;
import com.spire.presentation.*;

import java.awt.Color;
import java.util.Properties;
//import java.util.Properties;

import java.io.File;
//import jakarta.mail.*;
import jakarta.mail.internet.*;

//import java.io.File;
//import java.util.Properties;
//import java.io.File;
import org.apache.poi.sl.usermodel.Insets2D;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
//import org.apache.poi.sl.usermodel.VerticalAlignment;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xslf.usermodel.*;

public class ReadExcelFileToList {

	public static List<person> readExalData(String fileName) {
		List<person> personsList = new ArrayList<person>();

		try {
			FileInputStream fis = new FileInputStream(fileName);
			Workbook workbook = null;

			if (fileName.toLowerCase().endsWith("xlsx")) {
				workbook = new XSSFWorkbook(fis);
			} else if (fileName.toLowerCase().endsWith("xls")) {
				workbook = new HSSFWorkbook(fis);
			}

			int numberOfSheets = workbook.getNumberOfSheets();

			for (int i = 0; i < numberOfSheets; i++) {
				Sheet sheet = workbook.getSheetAt(i);
				Iterator<Row> rowIterator = sheet.iterator();

				while (rowIterator.hasNext()) {
					String name = "";
					String gmail = "";

					Row row = rowIterator.next();
					Iterator<Cell> cellIterator = row.cellIterator();

					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();

						switch (cell.getCellType()) {
						case STRING:
							if (gmail.equalsIgnoreCase("")) {
								gmail = cell.getStringCellValue().trim();
							} else if (name.equalsIgnoreCase("")) {
								name = cell.getStringCellValue().trim();
							} else {
								// System.out.println("Random data::" + cell.getStringCellValue());
							}
							break;

						}
					}

					person p = new person(name, gmail);
					personsList.add(p);
				}
			}
			fis.close();

		} catch (IOException e) {
			e.printStackTrace();
		}

		return personsList;
	}

	public static void main(String args[]) {
		List<person> names = readExalData("C:\\Users\\ba664\\OneDrive\\Documents\\table.xlsx");
		for (person p : names) {
			if (!(p.getName().equals("Job ID"))) {
				System.out.println(p.getName() + p.getGmail());
				readPowerPoint(p.getName(), p.getGmail());
			}

		}

	}

	public static void converPPtx() {
		try {
			Presentation ppt = new Presentation();

			// Load a PowerPoint presentation
			ppt.loadFromFile("C:\\Users\\ba664\\OneDrive\\Documents\\output.pptx");

			// Save the whole PowerPoint to PDF
			ppt.saveToFile("C:\\Users\\ba664\\OneDrive\\Documents\\شهادة_حضور.pdf", FileFormat.PDF);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void sendMessage(String email) {
		converPPtx();
		final String username = "aasgh.devolper+certificate@gmail.com"; // بريدك
		final String password = "iuqi jucq iwek rupe"; // كلمة مرور التطبيق (App Password)

		String toEmail = email; // المستقبل
		String subject = "شهادة حضور دورة";
		String messageText = "";
		String filePath = "C:\\Users\\ba664\\OneDrive\\Documents\\شهادة_حضور.pdf";
		//String filePath = "C:\\Users\\ba664\\OneDrive\\Documents\\output.pptx";

		// إعدادات SMTP
		Properties props = new Properties();
		props.put("mail.smtp.auth", "true");
		props.put("mail.smtp.starttls.enable", "true");
		props.put("mail.smtp.host", "smtp.gmail.com");
		props.put("mail.smtp.port", "587");

		// إنشاء الجلسة
		Session session = Session.getInstance(props, new Authenticator() {
			protected PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(username, password);
			}
		});

		try {
			// إنشاء الرسالة
			Message message = new MimeMessage(session);
			message.setFrom(new InternetAddress(username, "شهادة حضور دورة"));
			message.setRecipients(Message.RecipientType.BCC, InternetAddress.parse(toEmail));
			message.setSubject(subject);

			// النص
			MimeBodyPart textPart = new MimeBodyPart();
			textPart.setText(messageText);

			// المرفق
			MimeBodyPart attachmentPart = new MimeBodyPart();
			attachmentPart.attachFile(new File(filePath));

			// تركيب النص والمرفق مع بعض
			Multipart multipart = new MimeMultipart();
			multipart.addBodyPart(textPart);
			multipart.addBodyPart(attachmentPart);

			// ربطها بالرسالة
			message.setContent(multipart);

			// send message
			Transport.send(message);
			System.out.println(" تم إرسال الإيميل مع الملف بنجاح");

		} catch (Exception e) {
			e.printStackTrace();
			System.out.println(" فشل في إرسال الإيميل");
		}
	}

	public static void searchAndReplacePowerPoint(XMLSlideShow ppt, String name, String gmail) {
		XSLFTextParagraph p;
		XSLFTextRun r1;
		for (XSLFSlide slide : ppt.getSlides()) {
			for (XSLFShape sh : slide.getShapes()) {
				String nameOfShape = sh.getShapeName();

				if (sh instanceof XSLFTextShape) {
					XSLFTextShape shape = (XSLFTextShape) sh;

					if (shape.getText().equals("<<Name>>")) {
						shape.setText("");

						// shape.setInsets(new java.awt.Insets(50,50,400,100));
						p = shape.addNewTextParagraph();
						p.setTextAlign(TextAlign.RIGHT);
						r1 = p.addNewTextRun();
						r1.setText(name);
						r1.setFontSize(29.0);
						r1.setBold(true);
						r1.setFontColor(Color.red);
						shape.setAnchor(new java.awt.Rectangle(50, 260, 500, 100));
						shape.setInsets(new Insets2D(10.0, 10.0, 10.0, 10.0));

					}
				}
			}

		}
	}

	public static void readPowerPoint(String name, String gmail) {
		try {
			XMLSlideShow ppt = new XMLSlideShow(
					new FileInputStream("C:\\Users\\ba664\\Downloads\\Telegram Desktop\\شهادة ورشة.pptx"));
			searchAndReplacePowerPoint(ppt, name, gmail);

			System.out.println(" العملية تمت بنجاح");
			FileOutputStream out = new FileOutputStream("C:\\Users\\ba664\\OneDrive\\Documents\\output.pptx");
			ppt.write(out);
			out.close();
			sendMessage(gmail);

		} catch (FileNotFoundException e) {
			e.printStackTrace();
			System.out.println(" ملف PowerPoint غير موجود");
		} catch (IOException e) {
			e.printStackTrace();
			System.out.println(" خطأ أثناء قراءة الملف");
		}
	}
}
