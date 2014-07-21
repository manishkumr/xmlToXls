package com.krishagni.xml;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.log4j.Logger;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 * 
 * @author Manish
 *
 */

public class XmlToXlsConverter {
	private static final Logger logger = Logger.getLogger(XmlToXlsConverter.class);
	private static File inputXmlFilesDir ;
	private static File outputXlsFileDir ;
	private static int rowNumberOther;
	private static int imageSectionRowNumber;
	private static String fileName;
	private static HSSFWorkbook workbook;

	public static void main(String [] args){
		//String [] args = {"C:\\Users\\krishagni\\Downloads\\Xml to Excel App\\Xml to CSV Auto\\Sample Xml Files","C:\\Users\\krishagni\\Downloads\\Xml to Excel App\\Xml to CSV Auto\\output"};

		
		XmlToXlsConverter converter = new XmlToXlsConverter();

		try {

			//check parameters
			converter.checkParam(args);
			//create workbook
			converter.createWorkbook();
			//process
			converter.convert();


		} catch (Exception e) {
			e.printStackTrace();
			logger.error("Error occurred while processing "+e.getMessage());
			logger.fatal("cannot proceed... aborting");
			return;
		}

	}


	private void createWorkbook() throws Exception {
		logger.info("creating workbook");
		workbook  = new HSSFWorkbook();
		File outFile = new File(outputXlsFileDir,"outFile.xls");
		FileOutputStream fileOut = new FileOutputStream(outFile);
		Sheet subjectTable = workbook.createSheet("Subject_Table");
		Sheet caseTable = workbook.createSheet("Case_Table");
		Sheet instanceTable = workbook.createSheet("Instance_Table");
		Sheet instanceStructureTable = workbook.createSheet("Instance_Structure_Table");
		Sheet instanceDisorderTable = workbook.createSheet("Instance_Disorder_Table");
		Sheet imagingDetailsTable = workbook.createSheet("Imaging_Details_Table");
		//create  header
		logger.info("creating header");
		String [] subjectTableHeaderArray = {"File Name","SubjectID","ExternalIdentifierSpace","ExternalIdentifierValue","Species","Gender","ScientificSpecies"};
		workbook = createHeader(workbook,subjectTable,subjectTableHeaderArray);
		String [] caseTableHeaderArray = {"File Name","CaseID","SubjectID","ReferenceID","Age in years","Case_ClinicalHistory","Case_Material_Description","Significance","CaseDetailsComplete","Difficulty Level (primary/intermediate/advanced)"};
		workbook = createHeader(workbook,caseTable,caseTableHeaderArray);
		String [] instanceHeaderArray = {"File Name ","InstanceID ","CaseID ","OriginalFileName ","RevisedFileName ","SliceURL ","Slice ID ","FileType ","MimeType ","ContentType ","Size ","Pixels ","Abnormal Tissue ","ContributorID ",
				"CopyrightOwnerID ","Credits ","FormalConsentProvided ","DepositAgreementVersion ","DateDeposited ","Approved ","DateApproved ","ForReview ","Reviewer ","ReviewNote ","Procedure ","TransferStatus ","TransferComment ","Keywords"};
		workbook = createHeader(workbook,instanceTable,instanceHeaderArray);
		String [] instanceStructureTableHaeaderArray = {"File Name ","InstanceID ","InstanceStructureCode ","InstanceStructureText"};
		workbook = createHeader(workbook,instanceStructureTable,instanceStructureTableHaeaderArray);
		String [] instanceDisorderTableHeaderArray = {"File Name ","InstanceID ","InstanceDisorderCode ","InstanceDisorderText ","F5"};
		workbook = createHeader(workbook,instanceDisorderTable,instanceDisorderTableHeaderArray);
		String [] imagingDetailsTableHeaderArray = {"File Name ","ImagingDetailID ","InstanceID ","Modality ","Orientation ","ImagePlane ","Manufacturer ","ImageScale ","SliceCount ","Description","Differential Diagnosis","Discussion","Comments","Prognosis","References"};
		workbook = createHeader(workbook,imagingDetailsTable,imagingDetailsTableHeaderArray);
		workbook.write(fileOut);
		fileOut.close();

	}


	private HSSFWorkbook createHeader(HSSFWorkbook workbook2,Sheet subjectTable, 
			String[] subjectTableHeaderArray) {

		HSSFCellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(HSSFColor.LIME.index);
		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		Row row = subjectTable.createRow(0);
		for (int i = 0; i < subjectTableHeaderArray.length; i++) {
			Cell cell = row.createCell(i);
			cell.setCellValue(subjectTableHeaderArray[i]);
			cell.setCellStyle(style);
		}
		return workbook;
	}


	private void convert() throws Exception {
		File [] inputXMLfiles = inputXmlFilesDir.listFiles();
		logger.info("Number of xml files to process is "+inputXMLfiles.length);
		rowNumberOther = 1;
		imageSectionRowNumber = 1;
		for (File xmlFile : inputXMLfiles) {
			fileName = xmlFile.getName();
			parseXML(xmlFile);
			rowNumberOther++;
		}
	}


	private void parseXML(File xmlFile) throws Exception {
		logger.info("parsing XML file "+xmlFile.getName());
		//write file name
		

		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
		Document doc = dBuilder.parse(xmlFile);
		doc.getDocumentElement().normalize();
		if (doc.hasChildNodes()) {
			XMLdata xd = new XMLdata();
			ArrayList<String> imageSection = new ArrayList<String>();
			
			xd.setFileName(fileName);
			xd.setImageSection(imageSection);
			printNote(doc.getChildNodes(),xd);
			xlsWriter(xd);

		}

	}

	private void xlsWriter(XMLdata xd) throws Exception {
		//write subjectTable
		writeCellValue(0,rowNumberOther,0,xd.getFileName());
		writeCellValue(0,rowNumberOther,5,xd.getSex());
		//case table
		writeCellValue(1, rowNumberOther, 0, xd.getFileName());
		writeCellValue(1, rowNumberOther, 3, xd.getDocumentType());
		writeCellValue(1, rowNumberOther, 4, xd.getAge());
		writeCellValue(1, rowNumberOther, 5, xd.getSection());
		writeCellValue(1, rowNumberOther, 6, xd.getTitle());
		writeCellValue(1, rowNumberOther, 9, xd.getDifficultyLevel());
		//instance table
		for (int i = 0; i < xd.imageSection.size(); i++) {
			writeCellValue(2, imageSectionRowNumber, 0, xd.getFileName());
			writeCellValue(2, imageSectionRowNumber, 3, xd.imageSection.get(i));
			writeCellValue(2, imageSectionRowNumber, 13, xd.getAuthor());
			writeCellValue(2, imageSectionRowNumber, 18, xd.getPublicationDate());
			writeCellValue(2, imageSectionRowNumber, 27, xd.getKeyword());
			imageSectionRowNumber++;
		}
		//instance structure table
		writeCellValue(3, rowNumberOther, 0, xd.getFileName());
		writeCellValue(3, rowNumberOther, 3, xd.getCategory());
		//instance disorder Table
		writeCellValue(4, rowNumberOther, 0, xd.getFileName());
		writeCellValue(4, rowNumberOther, 3, xd.getDiagnosis());
		//imaging details
		writeCellValue(5, rowNumberOther, 0, xd.getFileName());
		writeCellValue(5, rowNumberOther, 3, xd.getModality());
		writeCellValue(5, rowNumberOther, 8, xd.getFindings());
		writeCellValue(5, rowNumberOther, 10, xd.getDifferentialDiagnosis());
		writeCellValue(5, rowNumberOther, 11, xd.getDiscussion());
		writeCellValue(5, rowNumberOther, 12, xd.getComments());
		writeCellValue(5, rowNumberOther, 12, xd.getPrognosis());
		writeCellValue(5, rowNumberOther, 12, xd.getReferences());
	}


	private void writeCellValue(int sheetNumber, int rowNumber, int cellNumber, String cellValue) throws Exception {
		Sheet sheet = workbook.getSheetAt(sheetNumber);
		if(sheet.getRow(rowNumber)!=null){
			Cell cell = sheet.getRow(rowNumber).createCell(cellNumber);
			cell.setCellValue(cellValue);
		}
		else{
			Row row = sheet.createRow(rowNumber);
			Cell cell = row.createCell(cellNumber);
			cell.setCellValue(cellValue);
		}
		
		
		File outFile = new File(outputXlsFileDir,"outFile.xls");
		FileOutputStream fileOut = new FileOutputStream(outFile);
		workbook.write(fileOut);
		fileOut.close();
	}


	private XMLdata printNote(NodeList nodeList,XMLdata xd) {
		
		for (int count = 0; count < nodeList.getLength(); count++) {

			Node tempNode = nodeList.item(count);

			if (tempNode.getNodeType() == Node.ELEMENT_NODE) {

				//sex
				if(tempNode.getNodeName().equals("pt-sex")){
					xd.setSex(tempNode.getTextContent().trim());
				}
				//document-type
				if(tempNode.getNodeName().equals("document-type")){
					xd.setDocumentType(tempNode.getTextContent().trim());
				}
				//age
				if(tempNode.getNodeName().equals("pt-age")){
					xd.setAge(tempNode.getTextContent().trim());
				}
				//section
				if(tempNode.getNodeName().equals("p")&& tempNode.getParentNode().getAttributes().getNamedItem("heading").getNodeValue().equals("History")){
					xd.setSection(tempNode.getTextContent().trim());

				}
				//title
				if(tempNode.getNodeName().equals("title")){
					xd.setTitle(tempNode.getTextContent().trim());
				}
				//image-section multiple
				if(tempNode.getNodeName().equals("alternative-image")&& tempNode.getAttributes().getNamedItem("role").getNodeValue().equals("original-dimensions")){
					xd.getImageSection().add(tempNode.getAttributes().getNamedItem("src").getNodeValue());
				}
				//author multiple
				if(tempNode.getNodeName().equals("author")){
					xd.setAuthor(tempNode.getTextContent().trim());
				}
				//publication-date
				if(tempNode.getNodeName().equals("publication-date")){
					xd.setPublicationDate(tempNode.getTextContent().trim());
				}
				//category
				if(tempNode.getNodeName().equals("category")){
					xd.setCategory(tempNode.getTextContent().trim());
				}
				//Diagnosis
				if(tempNode.getNodeName().equals("p")&& tempNode.getParentNode().getAttributes().getNamedItem("heading").getNodeValue().equals("Diagnosis")){
					xd.setDiagnosis(tempNode.getTextContent().trim());

				}
				//modality
				if(tempNode.getNodeName().equals("modality")){
					xd.setModality(tempNode.getTextContent().trim());
				}
				//finding
				if(tempNode.getNodeName().equals("p")&& tempNode.getParentNode().getAttributes().getNamedItem("heading").getNodeValue().equals("Findings")){
					xd.setFindings(tempNode.getTextContent().trim());
				}
				//level
				if(tempNode.getNodeName().equals("level")){
					xd.setDifficultyLevel((tempNode.getTextContent().trim()));
				}
				//keyword
				if(tempNode.getNodeName().equals("keywords")){
					xd.setKeyword(((tempNode.getTextContent().trim())));
				}
				//differential diagnosis
				if(tempNode.getNodeName().equals("p")&& tempNode.getParentNode().getAttributes().getNamedItem("heading").getNodeValue().equals("DDx")){
					xd.setDifferentialDiagnosis((tempNode.getTextContent().trim()));

				}
				//Discussion
				if(tempNode.getNodeName().equals("p")&& tempNode.getParentNode().getAttributes().getNamedItem("heading").getNodeValue().equals("Discussion")){
					xd.setDiscussion((tempNode.getTextContent().trim()));

				}
				//Comments
				if(tempNode.getNodeName().equals("p")&& tempNode.getParentNode().getAttributes().getNamedItem("heading").getNodeValue().equals("Comments")){
					xd.setComments((tempNode.getTextContent().trim()));

				}
				//prognosis
				if(tempNode.getNodeName().equals("p")&& tempNode.getParentNode().getAttributes().getNamedItem("heading").getNodeValue().equals("Prognosis")){
					xd.setPrognosis((tempNode.getTextContent().trim()));

				}
				//reference
				if(tempNode.getNodeName().equals("reference")){
					xd.setReferences((tempNode.getTextContent().trim()));
				}
				if (tempNode.hasChildNodes()) {
					// loop again if has child nodes
					printNote(tempNode.getChildNodes(),xd);
				}
			}

		}
		
		return xd;

	}
	private void checkParam(String [] args) throws Exception {
		logger.info("Number of params: "+args.length);
		if(args.length<2){

			logger.error("Either input directory path or output directory path is missing");
			logger.error("please check parameters: please provide Input xml Directory Path and Output xls Directory path\n \t\t\t re run with: java -jar XmlToXlsConverter pathToInputXMLdirectory pathToOutputXlsDirectory");
			throw new Exception("Please check parameters");
		}
		else{
			logger.info("args 1: "+ args[0]);
			logger.info("args 2: "+ args[1]);
			inputXmlFilesDir = new File(args[0]);
			outputXlsFileDir = new File(args[1]);
			if(!inputXmlFilesDir.isDirectory()){
				logger.error("Input xml directory is not a Directory");
				throw new Exception("Input xml directory is not a Directory");
			}else if(!outputXlsFileDir.isDirectory()){
				logger.error("Output xls directory is not a Directory");
				throw new Exception("Ouput xls directory is not a Directory");
			}else if (!inputXmlFilesDir.canRead()){
				logger.error("Cannot read input xml directory");
				throw new Exception("Cannot read input xml directory");
			}//can write

		}
		logger.info("Input XML directory path is: "+inputXmlFilesDir.getAbsolutePath());
		logger.info("Output File directory is: "+outputXlsFileDir);

	}
}
