package devutility.external.poi.models;

import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;

import devutility.internal.lang.ClassHelper;
import devutility.internal.lang.models.EntityField;
import devutility.internal.util.CollectionUtils;

public class FieldColumnMap<T> {
	/**
	 * Type T Class object.
	 */
	private Class<T> clazz = null;

	/**
	 * Entries container
	 */
	private HashMap<String, FieldColumnEntry> map = new LinkedHashMap<>();

	/**
	 * EntityField objects for T.
	 */
	private List<EntityField> entityFields = new ArrayList<>();

	public FieldColumnMap() {
		Type type = getClass().getGenericSuperclass();

		if (type instanceof ParameterizedType) {
			ParameterizedType pType = (ParameterizedType) type;
			Type claz = pType.getActualTypeArguments()[0];

			if (claz instanceof Class) {
				this.clazz = (Class<T>) claz;
			}
		}

		entityFields = ClassHelper.getEntityFields(clazz);
	}

	/**
	 * Find EntityField from entityFields.
	 * @param field: Field name in model.
	 * @return EntityField
	 */
	private EntityField findEntityField(String field) {
		return CollectionUtils.find(entityFields, i -> field.equals(i.getField().getName()));
	}

	/**
	 * Put a new or existed FieldColumnEntry object in map container.
	 * @param modelField: Field name for model class.
	 * @param columnIndex: Column index in sheet.
	 * @return FieldColumnEntry
	 */
	public synchronized FieldColumnEntry put(String modelField, int columnIndex) {

		FieldColumnEntry entry = new FieldColumnEntry(modelField, columnIndex);
		map.put(modelField, entry);
		return entry;
	}

	/**
	 * Return all entries in map container.
	 * @return List<FieldColumnEntry>
	 */
	public List<FieldColumnEntry> getEntries() {
		return new ArrayList<FieldColumnEntry>(map.values());
	}
}