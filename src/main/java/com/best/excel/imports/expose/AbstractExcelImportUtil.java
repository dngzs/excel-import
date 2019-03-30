package com.best.excel.imports.expose;


import com.best.excel.imports.config.ImportParams;
import com.best.excel.imports.handler.ReadRowHandler;
import com.best.excel.imports.parse.ISaxRowRead;
import com.best.excel.imports.sax.SaxReadExcel;
import com.best.excel.imports.xlsx.XlsSaxReadExcel;

import java.io.InputStream;

/**
 * 工具类
 */
public abstract class AbstractExcelImportUtil {

    public AbstractExcelImportUtil() {

    }


    /**
     * Excel 通过SAX解析方法,适合大数据导入,不支持图片
     * 导入 数据源本地文件,不返回校验结果 导入 字 段类型 Integer,Long,Double,Date,String,Boolean
     *
     * @param inputstream
     * @param pojoClass
     * @param params
     * @param hanlder
     */
    @SuppressWarnings("rawtypes")
    public static void importExcelXls(InputStream inputstream, Class<?> pojoClass,
                                      ImportParams params, ReadRowHandler hanlder, ISaxRowRead rowRead) throws Exception {
        new XlsSaxReadExcel().readExcel(inputstream, pojoClass, params, rowRead, hanlder, false);
    }


    /**
     * Excel 通过SAX解析方法,适合大数据导入
     * 导入 数据源本地文件,不返回校验结果 导入 字 段类型 Integer,Long,Double,Date,String,Boolean
     *
     * @param inputstream
     * @param pojoClass
     * @param params
     * @param hanlder
     */
    public static void importExcelBySax(InputStream inputstream, Class<?> pojoClass,
                                        ImportParams params, ReadRowHandler hanlder, ISaxRowRead rowRead) throws Exception {
        new SaxReadExcel().readExcel(inputstream, pojoClass, params, rowRead, hanlder);
    }


}
