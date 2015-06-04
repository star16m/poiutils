package star16m.utils.poi;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import star16m.utils.file.FileUtil;
import star16m.utils.poi.value.SimpleExcelBooleanValue;
import star16m.utils.poi.value.SimpleExcelDateValue;
import star16m.utils.poi.value.SimpleExcelIntValue;
import star16m.utils.poi.value.SimpleExcelStringValue;
import star16m.utils.poi.value.SimpleExcelValue;
import star16m.utils.string.StringUtil;

public class POIUtil {

    private static SimpleExcelValue readCell(Cell cell) {
    	
        switch( cell.getCellType()) {
        case Cell.CELL_TYPE_FORMULA :
        	Workbook wb = cell.getSheet().getWorkbook();
            CreationHelper crateHelper = wb.getCreationHelper();
            FormulaEvaluator evaluator = crateHelper.createFormulaEvaluator();
        	return new SimpleExcelStringValue(evaluator.evaluateInCell(cell).toString(), true);
        case Cell.CELL_TYPE_STRING :
            return new SimpleExcelStringValue(cell.getRichStringCellValue().getString(), false);
        case Cell.CELL_TYPE_NUMERIC :
            if(DateUtil.isCellDateFormatted(cell)) {
                return new SimpleExcelDateValue(cell.getDateCellValue());
            } else {
                return new SimpleExcelIntValue((int)cell.getNumericCellValue());
            }
        case Cell.CELL_TYPE_BOOLEAN :
            return new SimpleExcelBooleanValue(cell.getBooleanCellValue());
        default:
            return null;
        }
    }
    
    private static void setValue(Cell cell, SimpleExcelValue value) {
        value.setCellValue(cell);
    }
    
    private static String getColumnIndex(Cell cell) {
        return CellReference.convertNumToColString(cell.getColumnIndex());
    }
    
    private static int getColumnIndex(String columnLetter) {
        return CellReference.convertColStringToIndex(columnLetter);
    }
    
    private static int getRowIndex(Cell cell) {
        return cell.getRowIndex() + 1;
    }
    
    private static int getRowIndex(Integer rowIndex) {
        return rowIndex - 1;
    }
    
    private static Cell getCell(Sheet sheet, String columnLetter, Integer rowIndex) {
        Row row = sheet.getRow(getRowIndex(rowIndex));
        int columnIndex = getColumnIndex(columnLetter);
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }
        return cell;
    }
    
    private static void evaluate(POIReader poiReader) {
        if (poiReader.getVersion().equals(SpreadsheetVersion.EXCEL97)) {
            HSSFFormulaEvaluator.evaluateAllFormulaCells((HSSFWorkbook) poiReader.getWorkBook());
        } else if (poiReader.getVersion().equals(SpreadsheetVersion.EXCEL2007)) {
            XSSFFormulaEvaluator.evaluateAllFormulaCells((XSSFWorkbook) poiReader.getWorkBook());
        }
    }

    public static SimpleExcelTable readExcel(POIReader poiReader) throws IOException {
    	Workbook workbook = poiReader.getWorkBook();
    	// get sheet (order by follow as below. 1. sheet name. 2. sheet index(default:0))
    	String sheetName = poiReader.getSheetName();
    	if (StringUtil.isEmpty(poiReader.getSheetName())) {
    		sheetName = workbook.getSheetName(poiReader.getSheetIndex());
    	}
        SimpleExcelTable excelTable = new SimpleExcelTable();
        for( Row row : workbook.getSheet(sheetName) ) {
            for( Cell cell : row ) {
                excelTable.add(getColumnIndex(cell), getRowIndex(cell), readCell(cell));
            }
        }
        return excelTable;
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    public static void writeExcel(Workbook workbook, SimpleExcelTable replaceExcelTable, String outputFileName) throws IOException {
        writeExcel(workbook, replaceExcelTable, new File(outputFileName));
    }
    public static void writeExcel(Workbook workbook, SimpleExcelTable replaceExcelTable, File outputFile) throws IOException {
        writeExcel(workbook, 0, replaceExcelTable, new BufferedOutputStream(new FileOutputStream(outputFile)));
    }
    public static void writeExcel(Workbook workbook, SimpleExcelTable replaceExcelTable, OutputStream output) throws IOException {
        writeExcel(workbook, 0, replaceExcelTable, output);
    }
    public static void writeExcel(Workbook workbook, int workSheetIndex, SimpleExcelTable replaceExcelTable, OutputStream output) throws IOException {
        writeExcel(workbook, workbook.getSheetName(workSheetIndex), replaceExcelTable, output);
    }
    public static void writeExcel(Workbook workbook, String workSheetName, SimpleExcelTable replaceExcelTable, OutputStream output) throws IOException {
        Sheet sheet = workbook.getSheet(workSheetName);
        Cell cell = null;
        for (String columnLetter : replaceExcelTable.getColumns()) {
            for (Integer rowIndex : replaceExcelTable.getRows(columnLetter)) {
                cell = getCell(sheet, columnLetter, rowIndex);
                setValue(cell, replaceExcelTable.getValue(columnLetter, rowIndex));
            }
        }
        evaluate(workbook);
        workbook.write(output);
    }
    
    public static class POIReader {
    	private final Workbook workbook;
    	private SpreadsheetVersion version;
    	private int sheetIndex = 0;
    	private String sheetName;
    	private boolean detectTable;
    	private String start;
    	public POIReader(String fileName) throws IOException {
    		File file = new File(fileName);
            String fileExtension = FileUtil.getFileExtension(file);
            if (fileExtension.equalsIgnoreCase("xls")) {
                workbook = new HSSFWorkbook(new BufferedInputStream(new FileInputStream(file)));
                this.version = SpreadsheetVersion.EXCEL97;
            } else if (fileExtension.equalsIgnoreCase("xlsx")) {
                workbook = new XSSFWorkbook(new BufferedInputStream(new FileInputStream(file)));
                this.version = SpreadsheetVersion.EXCEL2007;
            } else {
            	throw new IOException("Unknown file format.");
            }
    	}
    	
    	public POIReader sheet(int sheetIndex) {
    		this.sheetIndex = sheetIndex;
    		return this;
    	}
    	public POIReader sheet(String sheetName) {
    		this.sheetName = sheetName;
    		return this;
    	}
    	public POIReader detect(boolean detect) {
    		this.detectTable = detect;
    		return this;
    	}
    	public POIReader start(String cellString) {
    		CellReference reference = new CellReference(cellString);
    		reference.get
    	}
    	
    	public Workbook getWorkBook() {
    		return workbook;
    	}
    	public SpreadsheetVersion getVersion() {
    		return this.version;
    	}
		public int getSheetIndex() {
			return sheetIndex;
		}

		public String getSheetName() {
			return sheetName;
		}

		public boolean isDetectTable() {
			return detectTable;
		}
    	
		public String getStart() {
			return start;
		}
    }
    public static void main(String[] args) throws Exception {
    	SimpleExcelTable excelTable = POIUtil.readExcel(new POIReader("C:\\data\\data.xlsx"));
    	System.out.println(excelTable);
    }
}
