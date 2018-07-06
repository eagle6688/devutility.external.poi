package devutility.external.poi.models;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;

public class FieldColumnMap {
	/**
	 * Entries container
	 */
	private HashMap<String, FieldColumnEntry> map = new LinkedHashMap<>();

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