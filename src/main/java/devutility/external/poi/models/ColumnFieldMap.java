package devutility.external.poi.models;

import java.util.LinkedHashMap;
import java.util.List;

import devutility.external.poi.common.ExcelColumnUtils;
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
	 */
	public ColumnFieldMap(Class<?> clazz) {
		List<EntityField> entityFields = ClassUtils.getEntityFields(clazz);

		for (int i = 0; i < entityFields.size(); i++) {
			EntityField entityField = entityFields.get(i);
			int columnIndex = ExcelColumnUtils.getIndex(entityField.getField());

			if (columnIndex < 0) {
				columnIndex = i;
			}

			if (this.containsKey(columnIndex)) {
				throw new IllegalArgumentException("Duplicated column index!");
			}

			put(columnIndex, entityField);
		}
	}
}