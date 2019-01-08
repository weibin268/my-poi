package com.zhuang.poi.excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.Map;

import static org.junit.Assert.*;

/**
 * Created by zhuang on 12/30/2017.
 */
public class WorkbookBuilderTest {

    @Test
    public void build() throws Exception {

        ArrayList<Map<String, Object>> data = new ArrayList<Map<String, Object>>();

        //Map<String, Object> rec = new HashMap<String, Object>();
        Map<String, Object> rec = new LinkedHashMap<String, Object>();

        rec.put("col1", "a");
        rec.put("col2", "b");
        rec.put("col3", "c");
        rec.put("col4", "d");


        data.add(rec);
        data.add(rec);

        WorkbookBuilder workbookBuilder = new WorkbookBuilder(ExcelType.XLSX);

        Workbook workbook = workbookBuilder
                .setCellCreatedHandler(new WorkbookBuilder.BuildHandler() {
                    public void handle(WorkbookBuilder.BuildContext context) {

                        if (context.getIsHeadRow()) {
                            CellStyle style = context.getWorkbook().createCellStyle();
                            Font font = context.getWorkbook().createFont();
                            style.setFont(font);
                            context.getCell().setCellStyle(style);

                            font.setBold(true);

                            font.setFontHeightInPoints((short) 11);
                        }
                    }
                })
                //.setColumnOrdinal("col4", 0)
                .removeColumn("col3")
                .setColumnCaption("col2","列2")
                .setDefaultColumnWidth(15)
                .buildSheet(data, "sheet1")
                .buildSheet(data, "sheet2")
                .build();

        FileOutputStream fileOutputStream = new FileOutputStream(new File("d:/temp/poitest.xlsx"));
        workbook.write(fileOutputStream);

    }

}