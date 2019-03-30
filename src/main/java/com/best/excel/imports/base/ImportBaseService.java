package com.best.excel.imports.base;

import com.best.excel.annotation.ExcelTitle;
import com.best.excel.imports.eneity.CollectionParams;
import com.best.excel.imports.eneity.ImportEntity;
import com.best.excel.imports.utils.PoiPublicUtil;
import com.best.excel.imports.utils.PoiReflectorUtil;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;

/**
 * 导入基础和,普通方法和Sax共用
 *
 * @author JueYue
 * 2015年1月9日 下午10:25:53
 */
public class ImportBaseService {


    /**
     * 把这个注解解析放到类型对象中
     *
     * @param targetId
     * @param field
     * @param excelEntity
     * @param pojoClass
     * @param temp
     * @throws Exception
     */
    public void addEntityToMap(String targetId, Field field, ImportEntity excelEntity,
                               Class<?> pojoClass,
                               Map<String, ImportEntity> temp) throws Exception {
        ExcelTitle excel = field.getAnnotation(ExcelTitle.class);
        excelEntity = new ImportEntity();
        excelEntity.setType(excel.type());
        excelEntity.setFormat(excel.format());
        excelEntity.setEnumImportMethod(excel.enumImportMethod());
        excelEntity.setName(PoiPublicUtil.getValueByTargetId(excel.name(), targetId, null));
        excelEntity.setMethod(PoiReflectorUtil.fromCache(pojoClass).getSetMethod(field.getName()));
        temp.put(excelEntity.getName(), excelEntity);
    }

    /**
     * 获取需要导出的全部字段
     *
     * @param targetId
     * @param fields
     * @param excelParams
     * @param pojoClass
     * @throws Exception
     */
    public void getAllExcelField(String targetId, Field[] fields,
                                 Map<String, ImportEntity> excelParams, Class<?> pojoClass) throws Exception {
        ImportEntity excelEntity = null;
        for (int i = 0; i < fields.length; i++) {
            Field field = fields[i];
            if (PoiPublicUtil.isNotUserExcelUserThis(null, field, targetId)) {
                continue;
            }
            if (PoiPublicUtil.isJavaClass(field) || field.getType().isEnum()) {
                addEntityToMap(targetId, field, excelEntity, pojoClass, excelParams);
            }
        }
    }


    public Object getFieldBySomeMethod(List<Method> list, Object t) throws Exception {
        Method m;
        for (int i = 0; i < list.size() - 1; i++) {
            m = list.get(i);
            t = m.invoke(t, new Object[]{});
        }
        return t;
    }


    /**
     * 多个get 最后再set
     *
     * @param setMethods
     * @param object
     */
    public void setFieldBySomeMethod(List<Method> setMethods, Object object,
                                     Object value) throws Exception {
        Object t = getFieldBySomeMethod(setMethods, object);
        setMethods.get(setMethods.size() - 1).invoke(t, value);
    }

    /**
     * @param entity
     * @param object
     * @param value
     * @throws Exception
     */
    public void setValues(ImportEntity entity, Object object, Object value) throws Exception {
        if (entity.getMethods() != null) {
            setFieldBySomeMethod(entity.getMethods(), object, value);
        } else {
            entity.getMethod().invoke(object, value);
        }
    }

}
