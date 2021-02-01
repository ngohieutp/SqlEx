package com.vpbank.sqlex.excel;

import org.apache.poi.ss.usermodel.Cell;

public class StringCellData implements CellData<String> {

    public final static StringCellData INSTANCE = new StringCellData();

    @Override
    public void setCellValue(Cell cell, String value) {
        cell.setCellValue(value);
    }

}
