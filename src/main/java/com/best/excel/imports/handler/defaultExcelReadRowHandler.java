package com.best.excel.imports.handler;

import com.best.excel.annotation.ExcelTitle;
import com.best.excel.exception.ExcelImportException;
import com.best.excel.imports.config.ImportParams;
import com.best.excel.imports.utils.PoiValidationUtil;
import org.apache.commons.lang3.StringUtils;

import java.lang.reflect.Field;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * excel数据处理（如果存在oom，建议sax使用它）
 *
 * @author BG317957
 * @date 2019年3月29日13:48:28
 */
public class defaultExcelReadRowHandler<T> implements ReadRowHandler<T> {

    private final ImportParams importParams;
    private List<T> ts;
    private AtomicInteger count = new AtomicInteger(0);

    public defaultExcelReadRowHandler(ImportParams importParams, List<T> ts) {
        this.importParams = importParams;
        this.ts = ts;
    }

    public void handler(T t) {
        if (importParams.getReadRows() > 0 && count.addAndGet(1) > importParams.getReadRows()) {
            throw new ExcelImportException("超过最大导入条数" + importParams.getReadRows());
        }
        if (checkObjAllFieldsIsNull(t)) {
            throw new ExcelImportException("excel存在假空行，请删除最下面的假空行数据或者重新下载模板，填写正确的值,再导入");
        }
        String validation = PoiValidationUtil.validation(t, importParams.getVerfiyGroup());
        if (StringUtils.isNotBlank(validation)) {
            throw new ExcelImportException(validation);
        }

        ts.add(t);
    }

    /**
     * 判断对象中属性值是否全为空
     *
     * @param object
     * @return
     */
    public boolean checkObjAllFieldsIsNull(T object) {
        if (null == object) {
            return true;
        }
        try {
            for (Field f : object.getClass().getDeclaredFields()) {
                if (f.getAnnotation(ExcelTitle.class) == null) {
                    continue;
                }
                f.setAccessible(true);
                if (f.get(object) != null && StringUtils.isNotBlank(f.get(object).toString())) {
                    return false;
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return true;
    }
}
