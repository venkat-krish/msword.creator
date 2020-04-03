package worddoc.components;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.sl.usermodel.PictureData.PictureType;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class MSWordDocumentCreator {

	private File getOutputFolder() {
		URL ouputDir = this.getClass().getClassLoader().getResource("output/");

		if (ouputDir == null) {
			throw new IllegalArgumentException("Folder not found");
		} else {
			return new File(ouputDir.getPath());
		}
	}
	
	/**
	 * getImageFileAsStream - method opens the image file and creates an inputstream object
	 * @throws FileNotFoundException 
	 * 
	 * */
	private FileInputStream getImageFileAsStream() throws FileNotFoundException {
		
		URL imageFile = this.getClass().getClassLoader().getResource("input/images/Protection1.jpg");
		
		if(imageFile == null) {
			throw new FileNotFoundException("Image file not found");
		}
		else {
			// New FileInputStream object
			FileInputStream fis = new FileInputStream(new File(imageFile.getFile()));
			return fis;
		}
	}
	
	/**
	 * Create word document with paragraph and image into the document
	 * 
	 */
	public void createWordDocument() {
		
		// Initializing the document object
		XWPFDocument document = new XWPFDocument();
		
		// Adding a new paragraph
		XWPFParagraph paragraph = document.createParagraph();
		
		// Create a page
		XWPFRun page = paragraph.createRun();
		page.setText("Prevention methods for COVID-19");
		page.setBold(true);
		paragraph.setAlignment(ParagraphAlignment.CENTER);
		
		// Open file
		try {
			String imgFile = "protection.jpg";
			FileInputStream imgStream = getImageFileAsStream();

			page.addPicture(imgStream, XWPFDocument.PICTURE_TYPE_JPEG, imgFile, Units.toEMU(480), Units.toEMU(375));
			imgStream.close();
			
			// Creating an output stream
			FileOutputStream docOutStream = new FileOutputStream("prevention.docx");
			document.write(docOutStream);
			docOutStream.close();
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

	public static void main(String[] args) throws Exception {
		// Initializing the document creator object 
		MSWordDocumentCreator msWordDoc = new MSWordDocumentCreator();

		msWordDoc.createWordDocument();
		
		System.out.println("Document created successfully...");
	}
}
