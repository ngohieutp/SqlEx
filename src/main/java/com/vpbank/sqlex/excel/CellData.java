package com.vpbank.sqlex.excel;

import org.apache.poi.ss.usermodel.Cell;

public interface CellData<T> {

    void setCellValue(Cell cell, T value);

}
