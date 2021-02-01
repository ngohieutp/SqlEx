package com.vpbank.sqlex.excel;

import org.apache.poi.ss.usermodel.Cell;

import java.util.Date;

public class DateCellData implements CellData<Date> {

    public final static DateCellData INSTANCE = new DateCellData();

    @Override
    public void setCellValue(Cell cell, Date value) {
        cell.setCellValue(value);
    }

}
