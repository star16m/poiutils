package star16m.utils.poi.value;

import org.apache.poi.ss.usermodel.Cell;

public class SimpleExcelBooleanValue implements SimpleExcelValue {

    private boolean value;
    
    public SimpleExcelBooleanValue(boolean value) {
        this.value = value;
    }
    public void setCellValue(Cell cell) {
        cell.setCellValue(this.value);
    }
    
    public String toString() {
        return String.valueOf(value);
    }
}
