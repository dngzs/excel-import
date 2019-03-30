package com.best.excel.imports.utils;


import javax.validation.ConstraintViolation;
import javax.validation.Validation;
import javax.validation.Validator;
import javax.validation.ValidatorFactory;
import java.util.Set;

/**
 * jsr303校验
 *
 * @author bg317957
 * @date 2019年3月29日19:59:09
 */
public class PoiValidationUtil {

    private final static Validator VALIDATOR;

    static {
        ValidatorFactory factory = Validation.buildDefaultValidatorFactory();
        VALIDATOR = factory.getValidator();
    }

    public static String validation(Object obj, Class[] verfiyGroup) {
        Set<ConstraintViolation<Object>> set = null;
        if (verfiyGroup != null) {
            set = VALIDATOR.validate(obj, verfiyGroup);
        } else {
            set = VALIDATOR.validate(obj);
        }
        if (set != null && set.size() > 0) {
            //获取到第一条信息
            for (ConstraintViolation<Object> constraintViolation : set) {
                return constraintViolation.getMessage();
            }
        }
        return null;
    }

}

