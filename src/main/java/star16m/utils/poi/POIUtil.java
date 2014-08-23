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
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import star16m.utils.poi.value.SimpleExcelBooleanValue;
import star16m.utils.poi.value.SimpleExcelDateValue;
import star16m.utils.poi.value.SimpleExcelIntValue;
import star16m.utils.poi.value.SimpleExcelStringValue;
import star16m.utils.poi.value.SimpleExcelValue;
import star16m.utils.file.FileUtil;

public class POIUtil {

    private static SimpleExcelValue readCell(Cell cell) {
        switch( cell.getCellType()) {
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
        case Cell.CELL_TYPE_FORMULA :
            return new SimpleExcelStringValue(cell.getCellFormula(), true);
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
    
    private static void evaluate(Workbook workbook) {
        if (workbook instanceof HSSFWorkbook) {
            HSSFFormulaEvaluator.evaluateAllFormulaCells((HSSFWorkbook) workbook);
        } else if (workbook instanceof XSSFWorkbook) {
            XSSFFormulaEvaluator.evaluateAllFormulaCells((XSSFWorkbook) workbook);
        }
    }

    public static Workbook getWorkBook(String fileName) throws IOException {
        return getWorkBook(new File(fileName));
    }
    public static Workbook getWorkBook(File file) throws IOException {
        String fileExtension = FileUtil.getFileExtension(file);
        Workbook workbook = null;
        if (fileExtension.equalsIgnoreCase("xls")) {
            workbook = new HSSFWorkbook(new BufferedInputStream(new FileInputStream(file)));
        } else if (fileExtension.equalsIgnoreCase("xlsx")) {
            workbook = new XSSFWorkbook(new BufferedInputStream(new FileInputStream(file)));
        }
        return workbook;
    }
    public static SimpleExcelTable readExcel(String fileName) throws IOException {
        return readExcel(new File(fileName));
    }
    public static SimpleExcelTable readExcel(File file) throws IOException {
        Workbook workbook = getWorkBook(file);
        if (workbook == null) {
            throw new IllegalArgumentException();
        }
        return readExcel(getWorkBook(file));
    }
    public static SimpleExcelTable readExcel(Workbook workbook) throws IOException {
        return readExcel(workbook, 0);
    }
    public static SimpleExcelTable readExcel(Workbook workbook, int sheetIndex) throws IOException {
        return readExcel(workbook, workbook.getSheetName(sheetIndex));
    }
    public static SimpleExcelTable readExcel(Workbook workbook, String workSheetName) throws IOException {
        SimpleExcelTable excelTable = new SimpleExcelTable();
        for( Row row : workbook.getSheet(workSheetName) ) {
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
}
