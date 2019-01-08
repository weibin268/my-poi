package com.zhuang.poi.excel;

import java.beans.IntrospectionException;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.zhuang.poi.util.AnnotationUtils;
import com.zhuang.poi.util.ConvertUtils;

/**
 * 用于Excel的导出
 *
 * @author zhuang
 */
public class WorkbookAdapter {

    public class AdaptContext {

        private WorkbookAdapter workbookAdapter;
        private Workbook workbook;
        private Row row;
        private Cell cell;
        private Object dataRow;
        private Object dataCell;

        public WorkbookAdapter getWorkbookAdapter() {
            return workbookAdapter;
        }

        public void setWorkbookAdapter(WorkbookAdapter workbookAdapter) {
            this.workbookAdapter = workbookAdapter;
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

        public Object getDataRow() {
            return dataRow;
        }

        public void setDataRow(Object dataRow) {
            this.dataRow = dataRow;
        }

        public Object getDataCell() {
            return dataCell;
        }

        public void setDataCell(Object dataCell) {
            this.dataCell = dataCell;
        }

    }

    public interface AdaptHandler {
        boolean handle(AdaptContext context);
    }

    private Workbook workbook;
    private int skipRowCount;
    private AdaptHandler rowCellAdaptHandler;
    private AdaptHandler rowAdaptHandler;

    public void setSkipRowCount(int skipRowCount) {
        this.skipRowCount = skipRowCount;
    }

    public WorkbookAdapter setRowCellAdaptHandler(AdaptHandler rowCellAdaptHandler) {
        this.rowCellAdaptHandler = rowCellAdaptHandler;
        return this;
    }

    public WorkbookAdapter setRowAdaptHandler(AdaptHandler rowAdaptHandler) {
        this.rowAdaptHandler = rowAdaptHandler;
        return this;
    }

    public WorkbookAdapter(ExcelType excelType, InputStream inputStream) {
        try {
            if (excelType == ExcelType.XLS) {
                workbook = new HSSFWorkbook(inputStream);
            } else {
                workbook = new XSSFWorkbook(inputStream);
            }
            skipRowCount = 1;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public <T> List<T> toList(Class<T> classT) {
        try {
            List<T> result = new ArrayList<T>();
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            Iterator<Row> rows = sheet.rowIterator();
            int colCount = headerRow.getLastCellNum();
            int rowCount = sheet.getLastRowNum();
            for (int i = 0; i < skipRowCount; i++) {
                rows.next();
            }
            rowLoop:
            while (rows.hasNext()) {
                Row row = rows.next();
                if (!row.cellIterator().hasNext())
                    continue;
                T entity = classT.getConstructor().newInstance();
                cellLoop:
                for (int i = 0; i < colCount; i++) {
                    Cell cell = row.getCell(i);
                    if (cell == null)
                        continue;
                    Method setMethod = AnnotationUtils.getSetMethodByColumnName(classT, headerRow.getCell(i).toString());
                    if (setMethod == null) {
                        throw new RuntimeException("excel 列（" + headerRow.getCell(i).toString() + "）找不到对应的实体属性！");
                    }
                    Class<?> classP = setMethod.getParameters()[0].getType();
                    Object objValue = null;
                    try {
                        objValue = ConvertUtils.changeType(cell.toString(), classP);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                    if (rowCellAdaptHandler != null) {
                        AdaptContext context = new AdaptContext();
                        context.setWorkbookAdapter(this);
                        context.setWorkbook(workbook);
                        context.setRow(row);
                        context.setCell(cell);
                        context.setDataCell(objValue);
                        boolean handleResult = rowCellAdaptHandler.handle(context);
                        if (!handleResult) {
                            continue cellLoop;
                        }
                    }
                    setMethod.invoke(entity, objValue);
                }
                if (rowAdaptHandler != null) {
                    AdaptContext context = new AdaptContext();
                    context.setWorkbookAdapter(this);
                    context.setWorkbook(workbook);
                    context.setRow(row);
                    context.setDataRow(entity);
                    boolean handleResult = rowAdaptHandler.handle(context);
                    if (!handleResult) {
                        continue rowLoop;
                    }
                }
                result.add(entity);
            }
            return result;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

}
