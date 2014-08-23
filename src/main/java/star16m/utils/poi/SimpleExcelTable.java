package star16m.utils.poi;

import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.TreeSet;

import star16m.utils.poi.value.SimpleExcelBooleanValue;
import star16m.utils.poi.value.SimpleExcelIntValue;
import star16m.utils.poi.value.SimpleExcelStringValue;
import star16m.utils.poi.value.SimpleExcelValue;

public class SimpleExcelTable {

    private Map<String, Map<Integer, SimpleExcelValue>> workbookTableByColumn;
    private Set<Integer> rowSet;
    
    public SimpleExcelTable ()  {
        workbookTableByColumn = new TreeMap<String, Map<Integer, SimpleExcelValue>>();
        rowSet = new TreeSet<Integer>();
    }
    public void add(String columnLetter, Integer rowIndex, SimpleExcelValue value) {
        Map<Integer, SimpleExcelValue> valueColumnMap = null;
        if (this.workbookTableByColumn.containsKey(columnLetter)) {
            valueColumnMap = this.workbookTableByColumn.get(columnLetter);
        } else {
            valueColumnMap = new TreeMap<Integer, SimpleExcelValue>();
            this.workbookTableByColumn.put(columnLetter, valueColumnMap);
        }
        valueColumnMap.put(rowIndex, value);
        rowSet.add(rowIndex);
    }
    public void add(String columnLetter, Integer rowIndex, String value) {
        add(columnLetter, rowIndex, value, false);
    }
    public void add(String columnLetter, Integer rowIndex, String value, boolean isFormula) {
        add(columnLetter, rowIndex, new SimpleExcelStringValue(value, isFormula));
    }
    public void add(String columnLetter, Integer rowIndex, int value) {
        add(columnLetter, rowIndex, new SimpleExcelIntValue(value));
    }
    public void add(String columnLetter, Integer rowIndex, boolean value) {
        add(columnLetter, rowIndex, new SimpleExcelBooleanValue(value));
    }
    
    public Set<String> getColumns() {
        return this.workbookTableByColumn.keySet();
    }
    public Set<Integer> getRows() {
        return this.rowSet;
    }
    public Set<Integer> getRows(String columnName) {
        return this.workbookTableByColumn.get(columnName).keySet();
    }
    
    public SimpleExcelValue getValue(String columnLetter, Integer rowIndex) {
        return this.workbookTableByColumn.get(columnLetter).get(rowIndex);
    }
    
    public boolean contains(String columnLetter) {
        return this.workbookTableByColumn.containsKey(columnLetter);
    }
    
    public boolean contains(Integer rowIndex) {
        return this.rowSet.contains(rowIndex);
    }
}
