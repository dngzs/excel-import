package com.best.excel.exception;

import com.best.excel.enmus.ExcelImportEnum;

/**
 * 导入异常
 *
 * @author zhangbo
 * @date 2019年3月30日12:40:23
 */
public class ExcelImportException extends RuntimeException {

    private static final long serialVersionUID = 7927443960049370453L;

    private ExcelImportEnum type;

    public ExcelImportException() {
        super();
    }

    public ExcelImportException(ExcelImportEnum type) {
        super(type.getMsg());
        this.type = type;
    }

    public ExcelImportException(ExcelImportEnum type, Throwable cause) {
        super(type.getMsg(), cause);
    }

    public ExcelImportException(String message) {
        super(message);
    }

    public ExcelImportException(String message, ExcelImportEnum type) {
        super(message);
        this.type = type;
    }

    public ExcelImportException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelImportEnum getType() {
        return type;
    }

    public void setType(ExcelImportEnum type) {
        this.type = type;
    }

}

