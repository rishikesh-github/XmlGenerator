package com.cts;

import org.springframework.boot.SpringApplication;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Scanner;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.w3c.dom.Document;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class XmLgeneratorApplication {

	private static String excelFile = "C:\\Users\\rishi\\Desktop\\xml generator\\Book1.xlsx";
	private static int alreadyExsistFile = 0;
	private static String fileSaveIn="";
	private static String xmlPath="";
	private static String inputExcelPath="";
	
	private static String sampleXmlTemplate = "<FormData>\n" + "<LetterType>Cancellation Letter</LetterType>\n"
			+ "<LetterCode>148</LetterCode>\n" + "<BrandName>ACRF</BrandName>\n"
			+ "<ProductName>Cancer Care Plus</ProductName>\n" + "<DocId>96703101</DocId>\n"
			+ "<CurrentDate>26/03/2020</CurrentDate>\n" + "<CustCareNo>1300 555 625</CustCareNo>\n"
			+ "<Add1>Unit 1 / 35 - 47 Tullidge St</Add1>\n" + "<Add2/>\n" + "<Add3/>\n" + "<Suburb>MELTON</Suburb>\n"
			+ "<State>VIC</State>\n" + "<Postcode>3337</Postcode>\n" + "<PolicyNo>4215912</PolicyNo>\n"
			+ "<SourceType/>\n" + "<SourceCode>26</SourceCode>\n" + "<CancellationDate>25/03/2020</CancellationDate>\n"
			+ "<PolicyOwners>\n" + "<PolicyOwner>\n" + "<Title>Mr</Title>\n" + "<FirstName>Craig</FirstName>\n"
			+ "<LastName>Davis</LastName>\n" + "</PolicyOwner>\n" + "<PolicyOwner>\n" + "<Title>Mr</Title>\n"
			+ "<FirstName>Craig</FirstName>\n" + "<LastName>Davis</LastName>\n" + "</PolicyOwner>\n"
			+ "</PolicyOwners>\n" + "<Insureds>\n" + "<Insured>\n" + "<Title>Mr</Title>\n"
			+ "<FirstName>Craig</FirstName>\n" + "<LastName>Davis</LastName>\n" + "</Insured>\n" + "</Insureds>\n"
			+ "</FormData>\n" + "";
	
	private static JLabel status;
	
	public static void main(String[] args) {

		SpringApplication.run(XmLgeneratorApplication.class, args);
		System.setProperty("java.awt.headless", "false"); 
		Gui();
		// ReadExcelFile();

	}
	public static void Gui() {

		JPanel panel = new JPanel();
		final JFrame frame = new JFrame();
		frame.setSize(550, 350);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.add(panel);
		panel.setLayout(null);

		JLabel header = new JLabel("XML Generator");
		header.setBounds(250, 30, 100, 25);
		panel.add(header);
		

		JLabel selectDirectory = new JLabel("File Save in");
		selectDirectory.setBounds(50, 80, 100, 25);
		panel.add(selectDirectory);

		final JTextField path = new JTextField();
		path.setBounds(150, 80, 300, 25);
		panel.add(path);

		JLabel outputFileName = new JLabel("Xml Path");
		outputFileName.setBounds(50, 125, 100, 25);
		panel.add(outputFileName);

		final JTextField ofn = new JTextField(20);
		ofn.setBounds(150, 125, 300, 25);
		panel.add(ofn);

		JButton btn = new JButton("...");
		btn.setBounds(450, 80, 30, 25);
		btn.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				int userChoice = fileChooser.showOpenDialog(frame);
				if (userChoice == JFileChooser.APPROVE_OPTION) {
					File file = fileChooser.getSelectedFile();
					path.setText(file.getPath());
					fileSaveIn = file.getPath();
					System.out.println(file.getPath());
				} else {
					path.setText("No file selected");
				}
			}
		});

		panel.add(btn);
		
		JButton xbtn2 = new JButton("...");
		xbtn2.setBounds(450, 125, 30, 25);
		xbtn2.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
				int userChoice = fileChooser.showOpenDialog(frame);
				if (userChoice == JFileChooser.APPROVE_OPTION) {
					File file = fileChooser.getSelectedFile();
					ofn.setText(file.getPath());
					xmlPath = file.getPath();
					System.out.println(file.getPath());
				} else {
					path.setText("No file selected");
				}
			}
		});

		panel.add(xbtn2);

		JLabel inputXML = new JLabel("Input Excel");
		inputXML.setBounds(50, 170, 100, 25);
		panel.add(inputXML);

		final JTextField pathinputxml = new JTextField();
		pathinputxml.setBounds(150, 170, 300, 25);
		panel.add(pathinputxml);

		JButton btn2 = new JButton("...");
		btn2.setBounds(450, 170, 30, 25);
		btn2.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
				int userChoice = fileChooser.showOpenDialog(frame);
				if (userChoice == JFileChooser.APPROVE_OPTION) {
					File file = fileChooser.getSelectedFile();
					pathinputxml.setText(file.getPath());
					inputExcelPath = file.getPath();
					System.out.println(file.getPath());
				} else {
					pathinputxml.setText("No File selected");
				}
			}
		});
		panel.add(btn2);

		JButton submit = new JButton("Submit");
		submit.setBounds(200, 210, 80, 25);
		
		status = new JLabel("");
		status.setBounds(50, 250, 150, 25);
		panel.add(status);
		
		
		final Map<String, String> map = new HashMap<String, String>();
		submit.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				System.out.println("btn clicked");
				ReadExcelFile();
				status.setText(alreadyExsistFile+" files generated..");
				try {
					
				}catch(Exception exc) {
					
					System.out.println(exc.getMessage());
				}
			}
		});

		panel.add(submit);

		frame.setVisible(true);
	}

	private static void ReadExcelFile() {
		System.out.println(fileSaveIn);
		System.out.println(xmlPath);
		System.out.println(inputExcelPath);

		try {

			FileInputStream excel = new FileInputStream(new File(inputExcelPath));
			
			File file = new File(xmlPath);
			Scanner scan = new Scanner(file);
			String sampleXmlTemplate = "";
			while(scan.hasNext()) {
				sampleXmlTemplate+=scan.next();
			}
			System.out.println(sampleXmlTemplate);
			
			Workbook workbook = new XSSFWorkbook(inputExcelPath);
			Sheet datatypeSheet = workbook.getSheetAt(0);
			Iterator<Row> iterator = datatypeSheet.iterator();
			iterator.next();
			while (iterator.hasNext()) {
				Row currentRow = iterator.next();
				
				Iterator<Cell> cellIterator = currentRow.iterator();
				String brand = "";
				String productName = "";
				String letterCode = "";
				String lettertype = "";
				int i=0;
				while (cellIterator.hasNext()) {
					Cell currentCell = cellIterator.next();
					if(i==0) {
						lettertype = currentCell.getStringCellValue();
					}else if(i==1) {
						brand = currentCell.getStringCellValue();
					}else if(i==2) {
						productName = currentCell.getStringCellValue();
					}
					else {
						letterCode  = ((int)currentCell.getNumericCellValue())+"";
					}
					i++;
				}
				
				constructxml(lettertype,brand,productName,letterCode);
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static void constructxml(String lettertype, String brand, String productName, String letterCode) {
		String res = sampleXmlTemplate.replaceAll("(?s)<LetterType[^>]*>.*?</LetterType>",
                "<LetterType>"+lettertype+"</LetterType>").replaceAll("(?s)<ProductName[^>]*>.*?</ProductName>",
                		"<ProductName>"+productName+"</ProductName>").replaceAll("(?s)<BrandName[^>]*>.*?</BrandName>",
                				"<BrandName>"+brand+"</BrandName>").replaceAll("(?s)<LetterCode[^>]*>.*?</LetterCode>",
                						"<LetterCode>"+letterCode+"</LetterCode>");
		    
			try {
				String filePath = fileSaveIn+"\\"
						+letterCode+"_"+lettertype+"_"+brand+"_"+productName+".xml";
				if(!new File(filePath).exists()) {					
					FileOutputStream fs  = new FileOutputStream(filePath, true);
					byte[] bytes = res.getBytes();
					fs.write(bytes);
					fs.close();
					alreadyExsistFile++;
				}
				else {
					System.out.println("FIle already exisit");
				
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		
	}
	
}
