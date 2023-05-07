	
	import java.io.*;
	import org.apache.poi.ss.usermodel.*;
	import org.apache.poi.xssf.usermodel.*;
	import org.jsoup.Jsoup;
	import org.jsoup.nodes.Document;
	import org.jsoup.nodes.Element;
	import org.jsoup.select.Elements;

	public class ClassA {
	  public static void main(String[] args) throws Exception {
	    // Load the HTML file
	 
	    File htmlFile = new File("C:\\Users\\HP\\git\\repository\\HTMLtoExcel\\HTML file\\reports.html");
	    Document doc = Jsoup.parse(htmlFile,"UTF-8","");
		
		  
		// Create a new Excel workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        // Create a new sheet within the workbook
        XSSFSheet sheet = workbook.createSheet("Sheet1");

        // Get the table element from the HTML data
        Element table = doc.select("table").first();

        // Get the table rows from the table element
        Elements rows = table.select("tr");

        // Loop through the rows and populate the Excel sheet
        int rowNum = 1;
        for (Element row : rows) {
            // Create a new row within the Excel sheet
            XSSFRow excelRow = sheet.createRow(rowNum++);

            // Get the table cells from the row
            Elements cells = row.select("td");

            // Loop through the cells and populate the Excel row
            int cellNum = 1;
            for (Element cell : cells) {
                // Create a new cell within the Excel row and set its value
                XSSFCell excelCell = excelRow.createCell(cellNum++);
                excelCell.setCellValue(cell.text());
            }
        }

        // Write the workbook to a file
        FileOutputStream fileOut = new FileOutputStream("HTML file\\output.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Close the workbook
        workbook.close();

        System.out.println("Excel sheet created successfully.");
	  }
	}
	

