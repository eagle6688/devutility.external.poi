package devutility.external.poi.model;

import java.lang.annotation.Annotation;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;

import devutility.external.poi.common.ExcelColumnUtils;
import devutility.internal.annotation.Ignore;
import devutility.internal.lang.ClassUtils;
import devutility.internal.model.ObjectField;
import devutility.internal.util.CollectionUtils;

/**
 * 
 * ColumnFieldMap, map between excel column and model field.
 * 
 * @author: Aldwin Su
 * @version: 2019-09-12 22:25:24
 */
public class ColumnFieldMap extends LinkedHashMap<Integer, ObjectField> {
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
		List<ObjectField> entityFields = ClassUtils.getNonAnnotationClassesEntityFields(annotaionClasses, clazz);
		List<ObjectField> nonExcelColumnEntityFields = new LinkedList<>();

		for (int i = 0; i < entityFields.size(); i++) {
			ObjectField entityField = entityFields.get(i);
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

		int startIndex = CollectionUtils.maxInt(this.keySet()) + 1;

		for (ObjectField nonExcelColumnEntityField : nonExcelColumnEntityFields) {
			put(startIndex++, nonExcelColumnEntityField);
		}
	}
}