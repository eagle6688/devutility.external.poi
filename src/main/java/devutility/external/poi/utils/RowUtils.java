package devutility.external.poi.utils;

import org.apache.poi.ss.usermodel.Row;

import devutility.external.poi.models.FieldColumnMap;
import devutility.internal.lang.ClassHelper;

public class RowUtils {
	public static <T> T toModel(Row row, FieldColumnMap<T> fieldColumnMap, Class<T> clazz) throws ReflectiveOperationException {
		T model = ClassHelper.newInstance(clazz);

		return model;
	}
}