package com.vpbank.sqlex.excel;

import org.apache.poi.ss.usermodel.Cell;

public class BooleanCellData implements CellData<Boolean> {

    public final static BooleanCellData INSTANCE = new BooleanCellData();

    @Override
    public void setCellValue(Cell cell, Boolean value) {
        cell.setCellValue(value);
    }

}
