package beantoexcel.excel.annotation;

import java.lang.annotation.*;

/**
 * @author guoxy
 * @description Excel简单通用工具
 * @create 2018-05-18 18:28
 **/
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface SheetCol {
	String value();
}
