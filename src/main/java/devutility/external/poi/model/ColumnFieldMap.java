package devutility.external.poi.model;

import java.lang.annotation.Annotation;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;

import devutility.external.poi.common.ExcelColumnUtils;
import devutility.internal.annotations.Ignore;
import devutility.internal.lang.ClassUtils;
import devutility.internal.lang.models.EntityField;

/**
 * 
 * ColumnFieldMap, map between excel column and model field.
 * 
 * @author: Aldwin Su
 * @version: 2019-09-12 22:25:24
 */
public class ColumnFieldMap extends LinkedHashMap<Integer, EntityField> {
	/**
	 * @Fields serialVersionUID
	 */
	private static final long serialVersionUID = 768142973835436871L;

	/**
	 * Constructor
	 * @param clazz Class object.
	 */
	public ColumnFieldMap(Class<?> clazz) {
		List<Class<? extends Annotation>> annotaionClasses = Arrays.asList(Ignore.class);
		List<EntityField> entityFields = ClassUtils.getEntityFields(clazz, annotaionClasses);
		List<EntityField> nonExcelColumnEntityFields = new LinkedList<>();

		for (int i = 0; i < entityFields.size(); i++) {
			EntityField entityField = entityFields.get(i);
			int columnIndex = ExcelColumnUtils.getIndex(entityField.getField());

			if (this.containsKey(columnIndex)) {
				throw new IllegalArgumentException("Duplicated column index!");
			}

			if (columnIndex < 0) {
				nonExcelColumnEntityFields.add(entityField);
				continue;
			}

			put(columnIndex, entityField);
		}

		for (EntityField nonExcelColumnEntityField : nonExcelColumnEntityFields) {

		}
	}
}