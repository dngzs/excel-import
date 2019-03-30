package com.best.excel.imports.utils;

import java.lang.reflect.Method;
import java.lang.reflect.Type;
import java.math.BigDecimal;
import java.sql.Time;
import java.sql.Timestamp;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

import com.best.excel.enmus.ExcelImportEnum;
import com.best.excel.exception.ExcelImportException;
import com.best.excel.imports.eneity.ImportEntity;
import com.best.excel.imports.sax.SaxReadCellEntity;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


/**
 * Cell 取值服务
 * 判断类型处理数据 1.判断Excel中的类型 2.根据replace替换值 3.handler处理数据 4.判断返回类型转化数据返回
 *
 * @author JueYue
 *  2014年6月26日 下午10:42:28
 */
public class CellValueService {

    private static final Logger LOGGER = LoggerFactory.getLogger(CellValueService.class);

    private List<String> handlerList = null;

    /**
     * 获取单元格内的值
     *
     * @param cell
     * @param entity
     * @return
     */
    private Object getCellValue(String classFullName, Cell cell, ImportEntity entity) {
        if (cell == null) {
            return "";
        }
        Object result = null;
        if ("class java.util.Date".equals(classFullName) || "class java.sql.Date".equals(classFullName)
                || ("class java.sql.Time").equals(classFullName)
                || ("class java.sql.Timestamp").equals(classFullName)) {
            /*
            if (Cell.CELL_TYPE_NUMERIC == cell.getCellType()) {
                // 日期格式
                result = cell.getDateCellValue();
            } else {
                cell.setCellType(Cell.CELL_TYPE_STRING);
                result = getDateData(entity, cell.getStringCellValue());
            }*/
            //FIX: 单元格yyyyMMdd数字时候使用 cell.getDateCellValue() 解析出的日期错误
            if (Cell.CELL_TYPE_NUMERIC == cell.getCellType() && DateUtil.isCellDateFormatted(cell)) {
                result = DateUtil.getJavaDate(cell.getNumericCellValue());
            } else {
                cell.setCellType(CellType.STRING);
                result = getDateData(entity, cell.getStringCellValue());
            }
            if (("class java.sql.Date").equals(classFullName)) {
                result = new java.sql.Date(((Date) result).getTime());
            }
            if (("class java.sql.Time").equals(classFullName)) {
                result = new Time(((Date) result).getTime());
            }
            if (("class java.sql.Timestamp").equals(classFullName)) {
                result = new Timestamp(((Date) result).getTime());
            }
        } else if (Cell.CELL_TYPE_NUMERIC == cell.getCellType() && DateUtil.isCellDateFormatted(cell)) {
            result = DateUtil.getJavaDate(cell.getNumericCellValue());
        } else {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                    result = cell.getRichStringCellValue() == null ? ""
                            : cell.getRichStringCellValue().getString();
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        if ("class java.lang.String".equals(classFullName)) {
                            result = formateDate(entity, cell.getDateCellValue());
                        }
                    } else {
                        result = readNumericCell(cell);
                    }
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    result = Boolean.toString(cell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_BLANK:
                    break;
                case Cell.CELL_TYPE_ERROR:
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    try {
                        result = readNumericCell(cell);
                    } catch (Exception e1) {
                        try {
                            result = cell.getRichStringCellValue() == null ? ""
                                    : cell.getRichStringCellValue().getString();
                        } catch (Exception e2) {
                            throw new RuntimeException("获取公式类型的单元格失败", e2);
                        }
                    }
                    break;
                default:
                    break;
            }
        }
        return result;
    }

    private Object readNumericCell(Cell cell) {
        Object result = null;
        double value = cell.getNumericCellValue();
        if (((int) value) == value) {
            result = (int) value;
        } else {
            result = value;
        }
        return result;
    }

    /**
     * 获取cell值
     * @param object
     * @param cellEntity
     * @param excelParams
     * @param titleString
     * @return
     */
    public Object getValue(Object object,
                           SaxReadCellEntity cellEntity, Map<String, ImportEntity> excelParams,
                           String titleString) {
        ImportEntity entity = excelParams.get(titleString);
        Method setMethod = entity.getMethods() != null && entity.getMethods().size() > 0
                ? entity.getMethods().get(entity.getMethods().size() - 1) : entity.getMethod();
        Type[] ts = setMethod.getGenericParameterTypes();
        String classFullName = ts[0].toString();
        Object result = cellEntity.getValue();
        return getValueByType(classFullName, result, entity, (Class) ts[0]);
    }

    /**
     * 根据返回类型获取返回值
     *
     * @param classFullName
     * @param result
     * @param entity
     * @param clazz
     * @return
     */
    private Object getValueByType(String classFullName, Object result, ImportEntity entity, Class clazz) {
        try {
            //过滤空和空字符串,如果基本类型null会在上层抛出,这里就不处理了
            if (result == null || StringUtils.isBlank(result.toString())) {
                return null;
            }
            if ("class java.util.Date".equals(classFullName) && result instanceof Date) {
                return result;
            } else if ("class java.util.Date".equals(classFullName) && result instanceof String) {
                return DateUtils.parseDate(result.toString(), entity.getFormat());
            }
            if ("class java.sql.Date".equals(classFullName) && result instanceof java.sql.Date) {
                return result;
            }
            if ("class java.lang.Boolean".equals(classFullName) || "boolean".equals(classFullName)) {
                return Boolean.valueOf(String.valueOf(result));
            }
            if ("class java.lang.Double".equals(classFullName) || "double".equals(classFullName)) {
                return Double.valueOf(String.valueOf(result));
            }
            if ("class java.lang.Long".equals(classFullName) || "long".equals(classFullName)) {
                try {
                    return Long.valueOf(String.valueOf(result));
                } catch (Exception e) {
                    //格式错误的时候,就用double,然后获取Int值
                    return Double.valueOf(String.valueOf(result)).longValue();
                }
            }
            if ("class java.lang.Float".equals(classFullName) || "float".equals(classFullName)) {
                return Float.valueOf(String.valueOf(result));
            }
            if ("class java.lang.Integer".equals(classFullName) || "int".equals(classFullName)) {
                try {
                    return Integer.valueOf(String.valueOf(result));
                } catch (Exception e) {
                    //格式错误的时候,就用double,然后获取Int值
                    return Double.valueOf(String.valueOf(result)).intValue();
                }
            }
            if ("class java.math.BigDecimal".equals(classFullName)) {
                return new BigDecimal(String.valueOf(result));
            }
            if ("class java.lang.String".equals(classFullName)) {
                //针对String 类型,但是Excel获取的数据却不是String,比如Double类型,防止科学计数法
                if (result instanceof String) {
                    return result;
                }
                // double类型防止科学计数法
                if (result instanceof Double) {
                    return PoiPublicUtil.doubleToString((Double) result);
                }
                return String.valueOf(result);
            }
            if (clazz != null && clazz.isEnum()) {
                if (StringUtils.isNotEmpty(entity.getEnumImportMethod())) {
                    return PoiReflectorUtil.fromCache(clazz).execEnumStaticMethod(entity.getEnumImportMethod(), result);
                } else {
                    return Enum.valueOf(clazz, result.toString());
                }
            }
            return result;
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new ExcelImportException(ExcelImportEnum.GET_VALUE_ERROR);
        }
    }


    /**
     * 获取日期类型数据
     *
     * @author JueYue
     *  2013年11月26日
     * @param entity
     * @param value
     * @return
     */
    private Date getDateData(ImportEntity entity, String value) {
        if (StringUtils.isNotEmpty(entity.getFormat()) && StringUtils.isNotEmpty(value)) {
            SimpleDateFormat format = new SimpleDateFormat(entity.getFormat());
            try {
                return format.parse(value);
            } catch (ParseException e) {
                LOGGER.error("时间格式化失败,格式化:{},值:{}", entity.getFormat(), value);
                throw new ExcelImportException(ExcelImportEnum.GET_VALUE_ERROR);
            }
        }
        return null;
    }

    private String formateDate(ImportEntity entity, Date value) {
        if (StringUtils.isNotEmpty(entity.getFormat()) && value != null) {
            SimpleDateFormat format = new SimpleDateFormat(entity.getFormat());
            return format.format(value);
        }
        return null;
    }



}
