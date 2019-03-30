package com.best.excel.imports.handler;

/**
 * 接口自定义处理类(用来处理解析的每行数据)
 */
public interface ReadRowHandler<T> {

    /**
     * 处理解析对象
     *
     * @param t
     */
    void handler(T t);
}
