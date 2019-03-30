package com.best.excel.imports.expose;

import com.best.excel.exception.ExcelImportException;
import com.best.excel.imports.config.ImportParams;
import com.best.excel.imports.handler.defaultExcelReadRowHandler;
import com.best.excel.imports.parse.SaxRowRead;

import java.io.InputStream;
import java.util.List;

/**
 * 默认文件导入工具类
 *
 * @author zhangbo
 * @date 2019年3月30日13:58:02
 */
public abstract class DefaultImportDataUtils extends AbstractExcelImportUtil {

    private DefaultImportDataUtils() {
    }

    /**
     * 03版本扩展名
     */
    static final String EXCEL03_EXTENSION = ".xls";

    /**
     * 07版本扩展
     *
     * @param inputstream
     * @param pojoClass
     * @param eos
     * @throws Exception
     */
    static final String EXCEL07_EXTENSION = ".xlsx";


    /**
     * 默认sax解析,默认两万行
     *
     * @param inputstream
     * @param pojoClass
     * @param eos
     * @throws Exception
     */
    public static void defaultImportExcelBySax(String fileName, InputStream inputstream, Class<?> pojoClass, List eos) throws Exception {
        ImportParams params = new ImportParams();
        defaultExcelReadRowHandler handler = new defaultExcelReadRowHandler(params, eos);
        params.setReadRows(20000);
        SaxRowRead saxRowRead = new SaxRowRead(pojoClass, params, handler);
        if (fileName.endsWith(EXCEL03_EXTENSION)) {
            AbstractExcelImportUtil.importExcelXls(inputstream, pojoClass, params, handler, saxRowRead);
        } else if (fileName.endsWith(EXCEL07_EXTENSION)) {
            AbstractExcelImportUtil.importExcelBySax(inputstream, pojoClass, params, handler, saxRowRead);
        } else {
            throw new ExcelImportException("文件格式错误，excel扩展名只能是xls或xlsx。");
        }
    }

}
