package com.zhuang.poi.excel;

import org.apache.poi.ss.usermodel.Cell;

public class CellUtils {

	public static void setColumnWidth(Cell cell, int width)
    {
        cell.getSheet().setColumnWidth(cell.getColumnIndex(), (int)((width + 0.72) * 256));
    }

}
