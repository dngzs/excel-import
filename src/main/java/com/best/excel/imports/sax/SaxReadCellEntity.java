package com.best.excel.imports.sax;


import com.best.excel.enmus.CellValueType;

/**
 * Cell 对象，用来保存cell的类型和值
 *
 * @author zhangbo
 * @date 2019年3月30日12:01:12
 */
public class SaxReadCellEntity {
    /**
     * 值类型
     */
    private CellValueType cellType;
    /**
     * 值
     */
    private Object value;

    public SaxReadCellEntity(CellValueType cellType, Object value) {
        this.cellType = cellType;
        this.value = value;
    }

    public CellValueType getCellType() {
        return cellType;
    }

    public void setCellType(CellValueType cellType) {
        this.cellType = cellType;
    }

    public Object getValue() {
        return value;
    }

    public void setValue(Object value) {
        this.value = value;
    }

    @Override
    public String toString() {
        return "[type=" + cellType.toString() + ",value=" + value + "]";
    }

}

