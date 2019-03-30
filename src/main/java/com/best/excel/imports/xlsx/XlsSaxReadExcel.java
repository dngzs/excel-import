package com.best.excel.imports.xlsx;


import com.best.excel.imports.config.ImportParams;
import com.best.excel.imports.handler.ReadRowHandler;
import com.best.excel.imports.parse.ISaxRowRead;
import com.best.excel.imports.xls.ExcelXlsReader;

import java.io.InputStream;
import java.util.List;

/**
 * 解析 2003版本的excl
 *
 * @author zhangbo
 * @date 2019年3月30日13:46:35
 */
public class XlsSaxReadExcel {

    public <T> List<T> readExcel(InputStream inputstream, Class<?> pojoClass, ImportParams params,
                                 ISaxRowRead rowRead, ReadRowHandler handler, boolean outputFormulaValues) throws Exception {
        try {
            new ExcelXlsReader(rowRead, inputstream, outputFormulaValues);
            return rowRead.getList();
        } catch (Exception e) {
            throw e;
        }
    }
}
