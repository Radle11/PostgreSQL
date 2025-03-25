import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Utility class for reading Excel files (.xls and .xlsx formats)
 * Requires Apache POI library
 */
public class ExcelReaderUtil {
    
    /**
     * Read an entire Excel workbook and return data as a list of sheets
     * 
     * @param filePath Path to the Excel file
     * @return List of sheets, where each sheet is a list of rows
     * @throws IOException If there's an issue reading the file
     */
    public static List<List<List<Object>>> readEntireWorkbook(String filePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = getWorkbook(fis, filePath)) {
            
            List<List<List<Object>>> allSheetsData = new ArrayList<>();
            
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                allSheetsData.add(readSheet(sheet));
            }
            
            return allSheetsData;
        }
    }
    
    /**
     * Read a specific sheet from an Excel file
     * 
     * @param filePath Path to the Excel file
     * @param sheetName Name of the sheet to read
     * @return List of rows from the specified sheet
     * @throws IOException If there's an issue reading the file
     */
    public static List<List<Object>> readSheet(String filePath, String sheetName) throws IOException {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = getWorkbook(fis, filePath)) {
            
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                throw new IllegalArgumentException("Sheet not found: " + sheetName);
            }
            
            return readSheet(sheet);
        }
    }
    
    /**
     * Read a specific sheet from an Excel file by index
     * 
     * @param filePath Path to the Excel file
     * @param sheetIndex Index of the sheet to read (0-based)
     * @return List of rows from the specified sheet
     * @throws IOException If there's an issue reading the file
     */
    public static List<List<Object>> readSheetByIndex(String filePath, int sheetIndex) throws IOException {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = getWorkbook(fis, filePath)) {
            
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            return readSheet(sheet);
        }
    }
    
    /**
     * Read a sheet as a list of maps, using the first row as headers
     * 
     * @param filePath Path to the Excel file
     * @param sheetName Name of the sheet to read
     * @return List of maps, where each map represents a row with header keys
     * @throws IOException If there's an issue reading the file
     */
    public static List<Map<String, Object>> readSheetAsMap(String filePath, String sheetName) throws IOException {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = getWorkbook(fis, filePath)) {
            
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                throw new IllegalArgumentException("Sheet not found: " + sheetName);
            }
            
            return readSheetAsMap(sheet);
        }
    }
    
    /**
     * Internal method to read a sheet and convert to a list of rows
     * 
     * @param sheet Sheet to read
     * @return List of rows, where each row is a list of cell values
     */
    private static List<List<Object>> readSheet(Sheet sheet) {
        List<List<Object>> sheetData = new ArrayList<>();
        
        for (Row row : sheet) {
            List<Object> rowData = new ArrayList<>();
            
            for (Cell cell : row) {
                rowData.add(getCellValue(cell));
            }
            
            sheetData.add(rowData);
        }
        
        return sheetData;
    }
    
    /**
     * Internal method to read a sheet as a list of maps
     * 
     * @param sheet Sheet to read
     * @return List of maps representing rows with header keys
     */
    private static List<Map<String, Object>> readSheetAsMap(Sheet sheet) {
        List<Map<String, Object>> sheetData = new ArrayList<>();
        
        // Get headers from the first row
        Row headerRow = sheet.getRow(0);
        List<String> headers = new ArrayList<>();
        for (Cell cell : headerRow) {
            headers.add(cell.getStringCellValue());
        }
        
        // Read data rows
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;
            
            Map<String, Object> rowData = new HashMap<>();
            for (int j = 0; j < headers.size(); j++) {
                Cell cell = row.getCell(j);
                rowData.put(headers.get(j), cell != null ? getCellValue(cell) : null);
            }
            
            sheetData.add(rowData);
        }
        
        return sheetData;
    }
    
    /**
     * Determine the workbook type based on file extension
     * 
     * @param fis FileInputStream of the Excel file
     * @param filePath Path to the Excel file
     * @return Workbook object (XSSFWorkbook or HSSFWorkbook)
     * @throws IOException If there's an issue reading the file
     */
    private static Workbook getWorkbook(FileInputStream fis, String filePath) throws IOException {
        if (filePath.toLowerCase().endsWith(".xlsx")) {
            return new XSSFWorkbook(fis);
        } else if (filePath.toLowerCase().endsWith(".xls")) {
            return new HSSFWorkbook(fis);
        } else {
            throw new IllegalArgumentException("Unsupported file format. Only .xls and .xlsx are supported.");
        }
    }
    
    /**
     * Extract the value from a cell based on its type
     * 
     * @param cell Cell to extract value from
     * @return Object representation of the cell value
     */
    private static Object getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                // Check if it's a date
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                }
                return cell.getNumericCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                // For formula cells, return the calculated value
                return getFormulaCellValue(cell);
            case BLANK:
                return null;
            default:
                return cell.toString();
        }
    }
    
    /**
     * Handle formula cell values
     * 
     * @param cell Formula cell
     * @return Calculated value of the formula
     */
    private static Object getFormulaCellValue(Cell cell) {
        FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
        CellValue cellValue = evaluator.evaluate(cell);
        
        switch (cellValue.getCellType()) {
            case NUMERIC:
                return cellValue.getNumberValue();
            case STRING:
                return cellValue.getStringValue();
            case BOOLEAN:
                return cellValue.getBooleanValue();
            case BLANK:
                return null;
            default:
                return cellValue.toString();
        }
    }
    
    /**
     * Example usage method (for demonstration)
     * 
     * @param args Command line arguments
     */
    /**
     * Get the number of rows in a specific sheet
     * 
     * @param filePath Path to the Excel file
     * @param sheetName Name of the sheet
     * @return Number of rows in the sheet (including the header row)
     * @throws IOException If there's an issue reading the file
     */
    public static int getRowCount(String filePath, String sheetName) throws IOException {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = getWorkbook(fis, filePath)) {
            
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                throw new IllegalArgumentException("Sheet not found: " + sheetName);
            }
            
            // Returns the last row index + 1 (which gives total number of rows)
            return sheet.getLastRowNum() + 1;
        }
    }
    
    /**
     * Get the number of rows in a sheet by index
     * 
     * @param filePath Path to the Excel file
     * @param sheetIndex Index of the sheet (0-based)
     * @return Number of rows in the sheet (including the header row)
     * @throws IOException If there's an issue reading the file
     */
    public static int getRowCountByIndex(String filePath, int sheetIndex) throws IOException {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = getWorkbook(fis, filePath)) {
            
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            // Returns the last row index + 1 (which gives total number of rows)
            return sheet.getLastRowNum() + 1;
        }
    }
    
    /**
     * Get the number of non-empty rows in a sheet
     * 
     * @param filePath Path to the Excel file
     * @param sheetName Name of the sheet
     * @return Number of non-empty rows in the sheet
     * @throws IOException If there's an issue reading the file
     */
    public static int getNonEmptyRowCount(String filePath, String sheetName) throws IOException {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = getWorkbook(fis, filePath)) {
            
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                throw new IllegalArgumentException("Sheet not found: " + sheetName);
            }
            
            int nonEmptyRowCount = 0;
            for (Row row : sheet) {
                if (!isRowEmpty(row)) {
                    nonEmptyRowCount++;
                }
            }
            
            return nonEmptyRowCount;
        }
    }
    
    /**
     * Check if a row is empty
     * 
     * @param row Row to check
     * @return true if the row is empty, false otherwise
     */
    private static boolean isRowEmpty(Row row) {
        if (row == null) {
            return true;
        }
        
        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
            Cell cell = row.getCell(c);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                return false;
            }
        }
        
        return true;
    }

    public static void main(String[] args) {
        try {
            String filePath = "path/to/your/file.xlsx";
            
            // Example of reading entire workbook
            List<List<List<Object>>> workbookData = readEntireWorkbook(filePath);
            
            // Example of reading a specific sheet by name
            List<List<Object>> sheetData = readSheet(filePath, "Sheet1");
            
            // Example of reading sheet as map
            List<Map<String, Object>> mappedSheetData = readSheetAsMap(filePath, "Sheet1");
            
            // Get row counts
            int totalRowCount = getRowCount(filePath, "Sheet1");
            int nonEmptyRowCount = getNonEmptyRowCount(filePath, "Sheet1");
            
            // Print some data (remove in production)
            System.out.println("Workbook Sheets: " + workbookData.size());
            System.out.println("First Sheet Total Rows: " + totalRowCount);
            System.out.println("First Sheet Non-Empty Rows: " + nonEmptyRowCount);
            System.out.println("First Sheet Data Rows: " + sheetData.size());
            System.out.println("Mapped Sheet Rows: " + mappedSheetData.size());
            
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
