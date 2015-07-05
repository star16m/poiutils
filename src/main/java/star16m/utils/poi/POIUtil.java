package star16m.utils.poi;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

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
    	if (cell == null) {
    		return null;
    	}
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
    	if (value == null) {
    		return;
    	}
        value.setCellValue(cell);
    }
    
//    private static int getColumnIndex(String columnLetter) {
//        return CellReference.convertColStringToIndex(columnLetter);
//    }
    
//    private static Cell getCell(Sheet sheet, String columnLetter, Integer rowIndex) {
//        Row row = sheet.getRow(rowIndex - 1);
//        int columnIndex = getColumnIndex(columnLetter);
//        Cell cell = row.getCell(columnIndex);
//        if (cell == null) {
//            cell = row.createCell(columnIndex);
//        }
//        return cell;
//    }
    
    private static void evaluate(POIHelper poiHelper, Workbook workbook) {
        if (poiHelper.getVersion().equals(SpreadsheetVersion.EXCEL97)) {
            HSSFFormulaEvaluator.evaluateAllFormulaCells((HSSFWorkbook)workbook);
        } else if (poiHelper.getVersion().equals(SpreadsheetVersion.EXCEL2007)) {
            XSSFFormulaEvaluator.evaluateAllFormulaCells((XSSFWorkbook)workbook);
        }
    }

    public static SimpleExcelTable read(POIHelper poiHelper) throws IOException {
    	Workbook workbook = poiHelper.getWorkBook(true);
    	// get sheet (order by follow as below. 1. sheet name. 2. sheet index(default:0))
    	String sheetName = poiHelper.getSheetName();
    	if (StringUtil.isEmpty(poiHelper.getSheetName())) {
    		sheetName = workbook.getSheetName(poiHelper.getSheetIndex());
    	}
        SimpleExcelTable excelTable = new SimpleExcelTable();
        Row row = null;
        Cell cell = null;
        Sheet sheet = workbook.getSheet(sheetName);
        for (int rowNum = Math.max(poiHelper.getStartRow(), sheet.getFirstRowNum()); rowNum <= sheet.getLastRowNum(); rowNum++) {
        	row = sheet.getRow(rowNum);
        	for (int colNum = Math.max(poiHelper.getStartCol(), row.getFirstCellNum()); colNum <= row.getLastCellNum(); colNum++) {
        		cell = row.getCell(colNum);
        		excelTable.add(CellReference.convertNumToColString(colNum), new Integer(rowNum), readCell(cell));
        	}
        }
        return excelTable;
    }
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    public static void write(POIHelper poiHelper, SimpleExcelTable excelTable) throws IOException {
    	File outputFile = new File(poiHelper.getFileName());
    	if (!poiHelper.overwrite() && outputFile.exists()) {
    		throw new IOException("Already exists file[" + poiHelper.getFileName() + "]");
    	}
        Workbook workbook = poiHelper.getWorkBook(false);
        
        String sheetName = "sheet" + (poiHelper.getSheetIndex() + 1);
        if (!StringUtil.isEmpty(poiHelper.getSheetName())) {
        	sheetName = poiHelper.getSheetName();
        }
        Sheet sheet = workbook.createSheet(sheetName);
        Row row = null;
        Cell cell = null;
        int realColIndex = 0;
        int realRowIndex = 0;
        for (Integer rowIndex : excelTable.getRows()) {
        	row = sheet.createRow(realRowIndex++);
        	realColIndex = 0;
        	for (String columnLetter : excelTable.getColumns()) {
            	cell = row.createCell(realColIndex++);
                setValue(cell, excelTable.getValue(columnLetter, rowIndex));
            }
        }
        evaluate(poiHelper, workbook);
        workbook.write(new BufferedOutputStream(new FileOutputStream(outputFile)));
    }
    
    public static class POIHelper {
    	private String fileName;
    	private SpreadsheetVersion version;
    	private int sheetIndex = 0;
    	private String sheetName;
    	private boolean detectTable;
    	private int startRow;
    	private int startCol;
    	private boolean overwrite;
    	public POIHelper(String fileName) throws IOException {
    		this.fileName = fileName;
    		File file = new File(fileName);
            String fileExtension = FileUtil.getFileExtension(file);
            if (fileExtension.equalsIgnoreCase("xls")) {
                this.version = SpreadsheetVersion.EXCEL97;
            } else if (fileExtension.equalsIgnoreCase("xlsx")) {
                this.version = SpreadsheetVersion.EXCEL2007;
            } else {
            	throw new IOException("Unknown file format.");
            }
    	}
    	
    	public String getFileName() {
    		return this.fileName;
    	}
    	public POIHelper sheet(int sheetIndex) {
    		this.sheetIndex = sheetIndex;
    		return this;
    	}
    	public POIHelper sheet(String sheetName) {
    		this.sheetName = sheetName;
    		return this;
    	}
    	public POIHelper detect(boolean detect) {
    		this.detectTable = detect;
    		return this;
    	}
    	public POIHelper start(String cellString) {
    		CellReference reference = new CellReference(cellString);
    		this.startRow = reference.getRow();
    		this.startCol = reference.getCol();
    		return this;
    	}
    	public POIHelper overwrite(boolean overwrite) {
    		this.overwrite = overwrite;
    		return this;
    	}
    	public boolean overwrite() {
    		return this.overwrite;
    	}
    	
    	public Workbook getWorkBook(boolean readFromFile) throws IOException {
    		Workbook workbook = null;
    		if (this.version.equals(SpreadsheetVersion.EXCEL97)) {
    			workbook = readFromFile ? new HSSFWorkbook(new BufferedInputStream(new FileInputStream(this.fileName))) : new HSSFWorkbook();
    		} else if (this.version.equals(SpreadsheetVersion.EXCEL2007)) {
    			workbook = readFromFile ? new XSSFWorkbook(new BufferedInputStream(new FileInputStream(this.fileName))) : new XSSFWorkbook();
    		}
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
    	
		public int getStartRow() {
			return this.startRow;
		}
		public int getStartCol() {
			return this.startCol;
		}
    }
//    public static void main(String[] args) throws Exception {
//    	SimpleExcelTable excelTable = POIUtil.read(new POIHelper("C:\\data\\data.xlsx").sheet(0).start("A3"));
//    	System.out.println(excelTable);
//    	
//    	POIUtil.write(new POIHelper("C:\\data\\haha.xlsx").overwrite(true), excelTable);
//    }
}
