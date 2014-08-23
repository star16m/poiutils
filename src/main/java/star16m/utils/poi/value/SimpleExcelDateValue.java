package star16m.utils.poi.value;

import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;

public class SimpleExcelDateValue implements SimpleExcelValue {

    private Date value;
    public SimpleExcelDateValue(Date value) {
        this.value = value;
    }
    public void setCellValue(Cell cell) {
        cell.setCellValue(this.value);
    }

    public String toString() {
        return String.valueOf(value);
    }
}
