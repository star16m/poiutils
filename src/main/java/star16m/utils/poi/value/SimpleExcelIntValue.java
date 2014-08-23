package star16m.utils.poi.value;

import org.apache.poi.ss.usermodel.Cell;

public class SimpleExcelIntValue implements SimpleExcelValue {

    private int value;
    public SimpleExcelIntValue(int value) {
        this.value = value;
    }
    public void setCellValue(Cell cell) {
        cell.setCellValue(this.value);
    }
    
    public String toString() {
        return String.valueOf(value);
    }
}
