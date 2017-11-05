package cn.patterncat.excel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * Created by patterncat on 2017-11-05.
 */
public class ExcelReader {

    public static <T> List<T> read(InputStream inputStream, int sheetIndex, int dataRowIndex, RowToBeanConverter<T> rowToBeanConverter) throws IOException, InvalidFormatException {
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            //总行数=sheet.getLastRowNum() + 1
            List<T> result = IntStream.range(dataRowIndex, sheet.getLastRowNum() + 1)
                    .mapToObj(i -> {
                        try{
                            Row row = sheet.getRow(i);
                            //如果sheet中间有空行,这里会是null
                            if (row == null) {
                                return null;
                            }
                            return rowToBeanConverter.from(row);
                        }catch (Exception e){
                            e.printStackTrace();
                            return null;
                        }
                    })
                    .filter(e -> e != null)
                    .collect(Collectors.toList());
            return result;
        } finally {
            if (workbook != null) {
                IOUtils.closeQuietly(workbook);
            }
            IOUtils.closeQuietly(inputStream);
        }
    }

    public static String getCellStringValue(Cell cell) {
        String cellValue = "";
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                cellValue = cell.getRichStringCellValue().getString();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    //todo 日期格式化问题
                    cellValue = cell.getDateCellValue().toString();
                } else {
                    //todo 小数点处理以及科学计数的处理
                    cellValue = String.valueOf(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                cellValue = Boolean.toString(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                cellValue = cell.getCellFormula();
                break;
            default:
                //do noting
        }
        return cellValue;
    }
}
