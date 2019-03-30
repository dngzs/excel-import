package com.best.excel.imports.parse;

import com.best.excel.imports.sax.SaxReadCellEntity;

import java.util.List;

/**
 * 数据回调解析
 */
public interface ISaxRowRead {
    /**
     * 获取返回数据
     *
     * @param <T>
     * @return
     */
    <T> List<T> getList();

    /**
     * 解析数据
     *
     * @param index
     * @param datas
     */
    void parse(int index, List<SaxReadCellEntity> datas);

}