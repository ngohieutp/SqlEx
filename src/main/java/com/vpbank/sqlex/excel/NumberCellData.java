package com.vpbank.sqlex.excel;

import org.apache.poi.ss.usermodel.Cell;

public class NumberCellData implements CellData<Number> {

    public final static NumberCellData INSTANCE = new NumberCellData();

    @Override
    public void setCellValue(Cell cell, Number value) {
        if (value != null) {
            cell.setCellValue(value.doubleValue());
        }
    }

}
