package com.best.excel.imports.eneity;

import java.util.Map;

public class CollectionParams {

    /**
     * 集合对应的名称
     */
    private String name;
    /**
     * Excel 列名称
     */
    private String excelName;
    /**
     * 实体对象
     */
    private Class<?> type;
    /**
     * 这个list下面的参数集合实体对象
     */
    private Map<String, ImportEntity> excelParams;

    public Map<String, ImportEntity> getExcelParams() {
        return excelParams;
    }

    public String getName() {
        return name;
    }

    public Class<?> getType() {
        return type;
    }

    public void setExcelParams(Map<String, ImportEntity> excelParams) {
        this.excelParams = excelParams;
    }

    public void setName(String name) {
        this.name = name;
    }

    public void setType(Class<?> type) {
        this.type = type;
    }

    public String getExcelName() {
        return excelName;
    }

    public void setExcelName(String excelName) {
        this.excelName = excelName;
    }
}

