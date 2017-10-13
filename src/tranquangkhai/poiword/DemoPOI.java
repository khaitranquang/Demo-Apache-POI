package tranquangkhai.poiword;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHeight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVerticalJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

public class DemoPOI {
	public static void main(String[] args) throws Exception {
		/* Note: You need a Referenced Libraries xmlbeans-x.x.x.jar in your local machine */
		/* I put it in Resource Folder*/
		
		/*Blank Document */
		XWPFDocument document = new XWPFDocument(); 
	          
	    /* Create Paragraph */
	    XWPFParagraph paragraph = document.createParagraph();
	    XWPFRun run = paragraph.createRun();
	    run.setText("At tutorialspoint.com, we strive hard to " +
	    			"provide quality tutorials for self-learning " +
	    			"purpose in the domains of Academics, Information " +
	    			"Technology, Management and Computer Programming Languages.");
				
	    /*
	     * Create table with formating text in XWPFTable
	     */
	    XWPFTable table = document.createTable();
	    XWPFTableRow tableRow1 = table.getRow(0);

	    
	    /* Removing paragraph before setText for cell*/
	    tableRow1.getCell(0).removeParagraph(0);
	    
	    /* Create cellRun to format text*/
	    XWPFRun cellRun1 = tableRow1.getCell(0).addParagraph().createRun();
	    cellRun1.setText("TRƯỜNG ĐẠI HỌC BÁCH KHOA HÀ NỘI");
	    cellRun1.addBreak();
	    cellRun1.setText("Thư viện Tạ Quang Bửu");
	    cellRun1.addBreak();
	    cellRun1.setText("Trần Quang Khải - 20152005");
	    cellRun1.addBreak();
	    cellRun1.setFontSize(12);
	    cellRun1.setBold(true);
	    
	    /* Insert image */
	    Path imagePath = Paths.get(ClassLoader.getSystemResource("bachkhoa.png").toURI());
	    cellRun1.addPicture(Files.newInputStream(imagePath),
	    					XWPFDocument.PICTURE_TYPE_PNG, imagePath.getFileName().toString(), 
	    					Units.toEMU(40), Units.toEMU(60));
	   
	    /* Set Alignment "CENTER" of data in cell */
	    XWPFTableCell cell1 = tableRow1.getCell(0);
	    XWPFParagraph paragraph1 = cell1.getParagraphs().get(0);
	    paragraph1.setAlignment(ParagraphAlignment.CENTER);
	  
	    /* Cell 2 */
	    XWPFRun cellRun2 = tableRow1.addNewTableCell().addParagraph().createRun();
	    tableRow1.getCell(1).removeParagraph(0);
	    cellRun2.setText("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM");
	    cellRun2.setFontSize(12);
	    cellRun2.setBold(true);
	    
	    /* Set width for table*/
	    table.getCTTbl().addNewTblPr().addNewTblW().setW(BigInteger.valueOf(10000));
	      
	    /*
	     * Write the Document in file system  Y
	     * You can reset path of file.
	     * Here, I use "C:/Users/User/Music/Desktop/DemoPOIWord.docx
	     */
	    FileOutputStream out = new FileOutputStream(new File("C:/Users/User/Music/Desktop/DemoPOIWord1.docx"));  
	      
	      
	    document.write(out);
	    out.close();
	    System.out.println("DemoPOIWord.docx written successfully");
	}
	
	
	/*
	 * Demo: Creating table with style
	 */
	public static void createStyledTable() throws Exception {
	    // Create a new document from scratch

	    try (XWPFDocument doc = new XWPFDocument()) {
	    	// -- OR --
	        // open an existing empty document with styles already defined
	    	//XWPFDocument doc = new XWPFDocument(new FileInputStream("base_document.docx"));

	    	// Create a new table with 6 rows and 3 columns
	    	int nRows = 6;
	        int nCols = 3;
	        XWPFTable table = doc.createTable(nRows, nCols);

	        // Set the table style. If the style is not defined, the table style
	        // will become "Normal".
	        CTTblPr tblPr = table.getCTTbl().getTblPr();
	        CTString styleStr = tblPr.addNewTblStyle();
	        styleStr.setVal("StyledTable");

	        // Get a list of the rows in the table
	        List<XWPFTableRow> rows = table.getRows();
	        int rowCt = 0;
	        int colCt = 0;
	        for (XWPFTableRow row : rows) {
	        	// get table row properties (trPr)
	            CTTrPr trPr = row.getCtRow().addNewTrPr();
	            // set row height; units = twentieth of a point, 360 = 0.25"
	            CTHeight ht = trPr.addNewTrHeight();
	            ht.setVal(BigInteger.valueOf(360));

	            // get the cells in this row
	            List<XWPFTableCell> cells = row.getTableCells();
	            // add content to each cell
	            for (XWPFTableCell cell : cells) {
	            	// get a table cell properties element (tcPr)
	                CTTcPr tcpr = cell.getCTTc().addNewTcPr();
	                // set vertical alignment to "center"
	                CTVerticalJc va = tcpr.addNewVAlign();
	                va.setVal(STVerticalJc.CENTER);

	                // create cell color element
	                CTShd ctshd = tcpr.addNewShd();
	                ctshd.setColor("auto");
	                ctshd.setVal(STShd.CLEAR);
	                if (rowCt == 0) {
	                	// header row
	                    ctshd.setFill("A7BFDE");
	                } 
	                else if (rowCt % 2 == 0) {
	                	// even row
	                    ctshd.setFill("D3DFEE");
	                } 
	                else {
	                	// odd row
	                    ctshd.setFill("EDF2F8");
	                }

	                // get 1st paragraph in cell's paragraph list
	                XWPFParagraph para = cell.getParagraphs().get(0);
	                // create a run to contain the content
	                XWPFRun rh = para.createRun();
	                // style cell as desired
	                if (colCt == nCols - 1) {
	                	// last column is 10pt Courier
	                    rh.setFontSize(10);
	                    rh.setFontFamily("Courier");
	                }
	                if (rowCt == 0) {
	                    // header row
	                    rh.setText("header row, col " + colCt);
	                    rh.setBold(true);
	                    para.setAlignment(ParagraphAlignment.CENTER);
	                } 
	                else {
	                    // other rows
	                    rh.setText("row " + rowCt + ", col " + colCt);
	                    para.setAlignment(ParagraphAlignment.LEFT);
	                }
	                colCt++;
	            } // for cell
	            colCt = 0;
	            rowCt++;
	        } // for row

	        // write the file
	        try (OutputStream out = new FileOutputStream("styledTable.docx")) {
	        	doc.write(out);
	        }
	    }
	}
}
