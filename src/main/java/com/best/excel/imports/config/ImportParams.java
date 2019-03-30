package com.best.excel.imports.config;

public class ImportParams {

    private Integer titleRows = 0;

    /**
     * 表头行数,默认1
     */
    private int headRows = 1;

    /**
     * 最大读取行数
     */
    private int readRows = 20000;

    /**
     * 校验组
     */
    private Class[] verfiyGroup;

    /**
     * 上传表格需要读取的sheet 数量,默认为1
     */
    private int sheetNum = 1;

    public int getSheetNum() {
        return sheetNum;
    }

    public void setSheetNum(int sheetNum) {
        this.sheetNum = sheetNum;
    }

    public Integer getTitleRows() {
        return titleRows;
    }

    public void setTitleRows(Integer titleRows) {
        titleRows = titleRows;
    }

    public int getHeadRows() {
        return headRows;
    }

    public void setHeadRows(int headRows) {
        this.headRows = headRows;
    }

    public int getReadRows() {
        return readRows;
    }

    public void setReadRows(int readRows) {
        this.readRows = readRows;
    }

    public Class[] getVerfiyGroup() {
        return verfiyGroup;
    }

    public void setVerfiyGroup(Class[] verfiyGroup) {
        this.verfiyGroup = verfiyGroup;
    }

}
