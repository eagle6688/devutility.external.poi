package devutility.external.poi.common;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;

/**
 * 
 * ExcelColumn
 * 
 * @author: Aldwin Su
 * @version: 2019-09-12 19:40:07
 */
@Retention(RUNTIME)
@Target(FIELD)
public @interface ExcelColumn {
	/**
	 * Index number of excel column, start from 0.
	 * @return int
	 */
	int index() default 0;
}