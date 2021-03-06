package com.best.excel.imports.parse;


import com.best.excel.annotation.ExcelEo;
import com.best.excel.exception.ExcelImportException;
import com.best.excel.imports.base.ImportBaseService;
import com.best.excel.imports.config.ImportParams;
import com.best.excel.imports.eneity.CollectionParams;
import com.best.excel.imports.eneity.ImportEntity;
import com.best.excel.imports.handler.ReadRowHandler;
import com.best.excel.imports.sax.SaxReadCellEntity;
import com.best.excel.imports.utils.CellValueService;
import com.best.excel.imports.utils.PoiPublicUtil;
import com.best.excel.imports.utils.PoiReflectorUtil;
import com.google.common.collect.Lists;
import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.util.*;

/**
 * 当行读取数据
 *
 * @author zhangbo
 * 2015年1月1日 下午7:59:39
 */
public class SaxRowRead extends ImportBaseService implements ISaxRowRead {

    private static final Logger LOGGER = LoggerFactory
            .getLogger(SaxRowRead.class);
    /**
     * 需要返回的数据
     **/
    private List list;
    /**
     * 导出的对象
     **/
    private Class<?> pojoClass;
    /**
     * 导入参数
     **/
    private ImportParams params;
    /**
     * 列表头对应关系
     **/
    private Map<Integer, String> titlemap = new HashMap<Integer, String>();
    /**
     * 当前的对象
     **/
    private Object object = null;

    private Map<String, ImportEntity> excelParams = new HashMap<String, ImportEntity>();

    private List<CollectionParams> excelCollection = new ArrayList<CollectionParams>();

    private String targetId;

    private CellValueService cellValueServer;

    private ReadRowHandler hanlder;

    public SaxRowRead(Class<?> pojoClass, ImportParams params, ReadRowHandler hanlder) {
        list = Lists.newArrayList();
        this.params = params;
        this.pojoClass = pojoClass;
        cellValueServer = new CellValueService();
        this.hanlder = hanlder;
        initParams(pojoClass);
    }

    private void initParams(Class<?> pojoClass) {
        try {

            Field[] fileds = PoiPublicUtil.getClassFields(pojoClass);
            ExcelEo etarget = pojoClass.getAnnotation(ExcelEo.class);
            if (etarget != null) {
                targetId = etarget.value();
            }
            getAllExcelField(targetId, fileds, excelParams, pojoClass);
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            e.printStackTrace();
            throw new ExcelImportException(e.getMessage());
        }

    }

    public <T> List<T> getList() {
        return list;
    }

    public void parse(int index, List<SaxReadCellEntity> datas) {
        try {
            if (datas == null || datas.size() == 0) {
                return;
            }
            //标题行跳过
            if (index < params.getTitleRows()) {
                return;
            }
            //表头行
            if (index < params.getTitleRows() + params.getHeadRows()) {
                addHeadData(datas);
            } else {
                addListData(datas);
            }
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new ExcelImportException(e.getMessage());
        }
    }

    /**
     * 集合元素处理
     *
     * @param datas
     */
    private void addListData(List<SaxReadCellEntity> datas) throws Exception {

        object = PoiPublicUtil.createObject(pojoClass, targetId);
        SaxReadCellEntity entity;
        for (int i = 0, le = datas.size(); i < le; i++) {
            entity = datas.get(i);
            String titleString = titlemap.get(i);
            if (excelParams.containsKey(titleString)) {
                saveFieldValue(object, entity, excelParams, titleString);
            }
        }
        if (object != null && hanlder != null) {
            hanlder.handler(object);
        }
        for (CollectionParams param : excelCollection) {
            addListContinue(object, param, datas, titlemap, targetId, params);
        }
        if (hanlder == null) {
            list.add(object);
        }

    }

    /**
     * 向List里面继续添加元素
     *
     * @param object
     * @param param
     * @param datas
     * @param titlemap
     * @param targetId
     * @param params
     * @throws Exception
     */
    private void addListContinue(Object object, CollectionParams param,
                                 List<SaxReadCellEntity> datas, Map<Integer, String> titlemap,
                                 String targetId, ImportParams params) throws Exception {
        Collection collection = (Collection) PoiReflectorUtil.fromCache(pojoClass).getValue(object,
                param.getName());
        Object entity = PoiPublicUtil.createObject(param.getType(), targetId);
        // 标记是否需要加上这个对象
        boolean isUsed = false;
        for (int i = 0; i < datas.size(); i++) {
            String titleString = titlemap.get(i);
            if (param.getExcelParams().containsKey(titleString)) {
                saveFieldValue(entity, datas.get(i), param.getExcelParams(), titleString);
                isUsed = true;
            }
        }
        if (isUsed) {
            collection.add(entity);
        }
    }

    /**
     * 设置值
     *
     * @param object
     * @param entity
     * @param excelParams
     * @param titleString
     * @throws Exception
     */
    private void saveFieldValue(Object object, SaxReadCellEntity entity,
                                Map<String, ImportEntity> excelParams,
                                String titleString) throws Exception {
        Object value = cellValueServer.getValue(object, entity,
                excelParams, titleString);
        setValues(excelParams.get(titleString), object, value);
    }

    /**
     * put 表头数据
     *
     * @param datas
     */
    private void addHeadData(List<SaxReadCellEntity> datas) {
        for (int i = 0; i < datas.size(); i++) {
            if (StringUtils.isNotEmpty(String.valueOf(datas.get(i).getValue()))) {
                titlemap.put(i, String.valueOf(datas.get(i).getValue()));
            }
        }
    }
}
