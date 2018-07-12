package devutility.external.poi.models;

import java.util.ArrayList;
import java.util.Comparator;
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

	/**
	 * Constructor
	 */
	public FieldColumnMap(Class<T> clazz) {
		this.clazz = clazz;
		entityFields = ClassHelper.getEntityFields(this.clazz);
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
	 * @param columnIndex: Column index in sheet.
	 * @param modelField: Field name for model class.
	 * @return FieldColumnEntry
	 */
	public synchronized FieldColumnEntry put(int columnIndex, String modelField) {
		FieldColumnEntry entry = new FieldColumnEntry(columnIndex, modelField);
		entry.setEntityField(findEntityField(modelField));
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

	/**
	 * Return all sorted entries in map container.
	 * @return List<FieldColumnEntry>
	 */
	public List<FieldColumnEntry> getSortedEntries() {
		List<FieldColumnEntry> list = getEntries();
		list.sort(Comparator.comparingInt(FieldColumnEntry::getColumnIndex));
		return list;
	}
}