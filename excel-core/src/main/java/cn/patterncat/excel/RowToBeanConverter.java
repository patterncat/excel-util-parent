package cn.patterncat.excel;

import org.apache.poi.ss.usermodel.Row;

/**
 * Created by patterncat on 2017-11-05.
 */
public interface RowToBeanConverter<T> {

    public T from(Row row);
}
