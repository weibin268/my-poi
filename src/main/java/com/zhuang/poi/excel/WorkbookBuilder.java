package com.zhuang.poi.excel;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 用于Excel的导入
 *
 * @author zhuang
 */
public class WorkbookBuilder {

    public class BuildContext {

        private WorkbookBuilder workbookBuilder;
        private Workbook workbook;
        private Row row;
        private Cell cell;
        private boolean isHeadRow;

        public WorkbookBuilder getWorkbookBuilder() {
            return workbookBuilder;
        }

        public void setWorkbookBuilder(WorkbookBuilder workbookBuilder) {
            this.workbookBuilder = workbookBuilder;
        }

        public Workbook getWorkbook() {
            return workbook;
        }

        public void setWorkbook(Workbook workbook) {
            this.workbook = workbook;
        }

        public Row getRow() {
            return row;
        }

        public void setRow(Row row) {
            this.row = row;
        }

        public Cell getCell() {
            return cell;
        }

        public void setCell(Cell cell) {
            this.cell = cell;
        }

        public boolean getIsHeadRow() {
            return isHeadRow;
        }

        public void setIsHeadRow(boolean isHeadRow) {
            this.isHeadRow = isHeadRow;
        }

    }

    public interface BuildHandler {
        void handle(BuildContext context);
    }

    private Workbook workbook;
    private List<String> columnNames;
    private BuildHandler rowCreatedHandler;
    private BuildHandler cellCreatedHandler;
    private Integer defaultColumnWidth;
    private Map<String, Integer> mapColumnWidth;
    private String currentColumnName;
    private Map<String, Integer> mapColumnOrdinal;
    private List<String> removeColumns;
    private Map<String, String> mapColumnCaption;

    public WorkbookBuilder setRowCreatedHandler(BuildHandler rowCreatedHandler) {
        this.rowCreatedHandler = rowCreatedHandler;
        return this;
    }

    public WorkbookBuilder setCellCreatedHandler(BuildHandler cellCreatedHandler) {
        this.cellCreatedHandler = cellCreatedHandler;
        return this;
    }

    public WorkbookBuilder setDefaultColumnWidth(int defaultColumnWidth) {
        this.defaultColumnWidth = defaultColumnWidth;
        return this;
    }

    public WorkbookBuilder setColumnOrdinal(String columnName, Integer ordinal) {
        currentColumnName = columnName == null ? currentColumnName : columnName;
        mapColumnOrdinal.put(currentColumnName, ordinal);
        return this;
    }

    public WorkbookBuilder setColumnOrdinal(Integer ordinal) {
        return setColumnOrdinal(null, ordinal);
    }

    public WorkbookBuilder removeColumn(String columnName) {

        removeColumns.add(columnName);

        return this;
    }

    public WorkbookBuilder setColumnCaption(String columnName, String caption) {
        currentColumnName = columnName == null ? currentColumnName : columnName;
        mapColumnCaption.put(currentColumnName, caption);
        return this;
    }

    public WorkbookBuilder setColumnCaption(String caption) {
        return setColumnCaption(null, caption);
    }

    public WorkbookBuilder(ExcelType excelType) {
        if (excelType == ExcelType.XLS) {
            workbook = new HSSFWorkbook();
        } else {
            workbook = new XSSFWorkbook();
        }
        columnNames = new ArrayList<String>();
        mapColumnWidth = new HashMap<String, Integer>();
        mapColumnOrdinal = new HashMap<String, Integer>();
        removeColumns = new ArrayList<String>();
        mapColumnCaption = new HashMap<String, String>();
    }

    public Workbook build() {
        return workbook;
    }

    public Workbook build(ArrayList<Map<String, Object>> data) {
        return build(data, null);
    }

    public Workbook build(ArrayList<Map<String, Object>> data, String sheetName) {
        buildSheet(data, sheetName);
        return build();
    }

    public WorkbookBuilder buildSheet(ArrayList<Map<String, Object>> data, String sheetName) {
        loadColumnNamesFromData(data);
        Sheet sheet = sheetName == null ? workbook.createSheet() : workbook.createSheet(sheetName);
        int currentRowIndex = 0;
        BuildContext context = new BuildContext();
        context.setWorkbook(workbook);
        Row headRow = sheet.createRow(currentRowIndex++);
        context.setRow(headRow);
        context.setCell(null);
        context.setIsHeadRow(true);
        if (rowCreatedHandler != null)
            rowCreatedHandler.handle(context);
        for (int i = 0; i < columnNames.size(); i++) {
            Cell tempCell = headRow.createCell(i);
            tempCell.setCellValue(getColumnCaption(columnNames.get(i)));
            context.setCell(tempCell);
            if (cellCreatedHandler != null)
                cellCreatedHandler.handle(context);
            setColumnWidthByCell(tempCell, columnNames.get(i));
        }
        for (int i = 0; i < data.size(); i++) {
            Map<String, Object> tempRecord = data.get(i);
            Row tempRow = sheet.createRow(currentRowIndex++);
            context.setRow(tempRow);
            context.setCell(null);
            context.setIsHeadRow(false);
            if (rowCreatedHandler != null)
                rowCreatedHandler.handle(context);
            for (int j = 0; j < columnNames.size(); j++) {
                Object objValue = tempRecord.get(columnNames.get(j));
                Cell tempCell = tempRow.createCell(j);
                context.setCell(tempCell);
                if (cellCreatedHandler != null)
                    cellCreatedHandler.handle(context);
                tempCell.setCellValue(objValue.toString());
            }
        }
        return this;
    }

    private void loadColumnNamesFromData(ArrayList<Map<String, Object>> data) {
        if (data.size() == 0) {
            throw new RuntimeException("Data为空！");
        }
        columnNames.clear();
        Map<String, Object> firstRecord = data.get(0);
        for (String key : firstRecord.keySet()) {
            columnNames.add(key);
        }
        applyColumnOrdinal();
        applyRemoveColumn();
    }

    private void setColumnWidthByCell(Cell cell, String columnName) {
        if (mapColumnWidth.containsKey(columnName)) {
            CellUtils.setColumnWidth(cell, mapColumnWidth.get(columnName));
        } else {
            if (defaultColumnWidth != null)
                CellUtils.setColumnWidth(cell, defaultColumnWidth);
        }
    }

    private void applyColumnOrdinal() {
        for (String key : mapColumnOrdinal.keySet()) {
            int index = columnNames.indexOf(key);
            if (index == -1)
                continue;
            String temp = columnNames.get(index);
            columnNames.remove(index);
            columnNames.add(mapColumnOrdinal.get(key), temp);
        }
    }

    private void applyRemoveColumn() {
        for (String col : removeColumns) {
            columnNames.remove(col);
        }
    }

    private String getColumnCaption(String columnName) {
        if (mapColumnCaption.containsKey(columnName)) {
            return mapColumnCaption.get(columnName);
        } else {
            return columnName;
        }
    }

}
