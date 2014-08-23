package star16m.utils.poi.value;

import org.apache.poi.ss.usermodel.Cell;

public class SimpleExcelStringValue implements SimpleExcelValue {

    private String value;
    private boolean isFormula;
    public SimpleExcelStringValue(String value, boolean isFormula) {
        this.value = value;
        this.isFormula = isFormula;
    }
    public void setCellValue(Cell cell) {
        if (isFormula) {
            cell.setCellFormula(this.value);
        } else {
            cell.setCellValue(this.value);
        }
    }
    
    public String toString() {
        return value;
    }
}
