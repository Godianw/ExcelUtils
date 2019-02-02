package com.godianw.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * excel工具类
 * @auther : Godianw
 * @version : v1.0
 */
public class ExcelUtils {

    private static final String SUFFIX_EXCEL_2003 = ".xls";
    private static final String SUFFIX_EXCEL_2007 = ".xlsx";

    private static final String DEFAULT_FONT = "微软雅黑";
    private static final short HEADER_FONT_SIZE = (short) 14;
    private static final short VALUE_FONT_SIZE = (short) 12;

    private static final String DEFAULT_DATE_PATTERN = "yyyy-MM-dd HH:mm:ss";
    private static final String DEFAULT_NUMERIC_PATTERN = "#.##";

    private ExcelUtils() {

    }

    /**
     * read excel file from inputStream.
     * @param in inputStream
     * @param header2ModelMap header and model field mapping.
     *                        <p>if the header field and the model field are consistent,
     *                        you can set parameter {@code header2ModelMap} to null.
     * @return mapping list of model fields and values
     * @throws IOException
     */
    public static List<Map<String, String>> read(
            InputStream in, Map<String, String> header2ModelMap) throws IOException {

        Workbook workbook = WorkbookFactory.create(in);
        DataFormatter formatter = new DataFormatter();
        List<Map<String, String>> resultList = new ArrayList<>();

        for (Sheet sheet : workbook) {
            boolean isHeaderRow = true;
            Map<Integer, String> modelMap = null;

            for (Row row : sheet) {
                // 解析头部行
                if (isHeaderRow) {
                    isHeaderRow = false;
                    modelMap = buildColumnMap(row);
                    continue;
                }

                // 解析数据行
                Map<String, String> resultMap = new HashMap<>();
                resultList.add(resultMap);
                for (Cell cell : row) {
                    String key = modelMap.get(cell.getColumnIndex());
                    if (key != null && header2ModelMap != null) {
                        key = header2ModelMap.get(key);
                    }
                    if (key == null) {
                        continue;
                    }
                    String value = getCellValue(cell, null);
                    resultMap.put(key, value);
                }
            }
        }
        return resultList;
    }

    /**
     * write excel file to outputStream
     * @param out outputStream
     * @param fileName the name of file to be written
     * @param model2HeaderMap model and header field mapping.
     *                        <p>if the model field and the header field are consistent,
     *                        you can set parameter {@code model2HeaderMap} to null.
     * @param sourceList data collection
     * @throws IllegalAccessException
     * @throws IOException
     */
    public static<T> void write(
            OutputStream out,
            String fileName,
            Map<String, String> model2HeaderMap,
            Collection<T> sourceList) throws IllegalAccessException, IOException {

        try (Workbook workbook = buildWorkBook(fileName)) {
            writeData2Sheet(workbook, model2HeaderMap, sourceList);
            workbook.write(out);
        }
    }

    private static<T> void writeData2Sheet(
            Workbook workbook,
            Map<String, String> model2HeaderMap,
            Collection<T> sourceList)
            throws IllegalAccessException {
        String safeName = WorkbookUtil.createSafeSheetName(UUID.randomUUID().toString());
        Sheet sheet = workbook.createSheet(safeName);

        T t;
        int currentRowIndex = 0; // 最近创建的行索引
        int columnCount = 0;
        Iterator iterator = sourceList.iterator();
        Row headerRow = sheet.createRow(currentRowIndex);
        Row valueRow = null;
        CellStyle headerCellStyle = buildHeaderCellStyle(workbook);
        CellStyle valueCellStyle = buildValueCellStyle(workbook);
        while (iterator.hasNext()) {
            t = (T) iterator.next();
            valueRow = sheet.createRow(++currentRowIndex);

            Field[] fields = t.getClass().getDeclaredFields();
            for (int i = 0; i < fields.length; ++i) {
                Field field = fields[i];
                // 创建第一行数据时填充表头
                if (currentRowIndex == 1) {
                    String name = field.getName();
                    if (model2HeaderMap != null && model2HeaderMap.get(name) != null) {
                        name = model2HeaderMap.get(name);
                    }
                    Cell cell = buildStyledCell(headerRow, i, headerCellStyle);
                    setCellValue(cell, name, null);
                    columnCount++;
                }
                field.setAccessible(true);
                Object value = field.get(t);
                Cell cell = buildStyledCell(valueRow, i, valueCellStyle);
                setCellValue(cell, value, null);
            }
        }

        // 所有列自适应大小
        for (int i = 0; i < columnCount; ++i) {
            sheet.autoSizeColumn(i);
        }
    }

    private static Map<Integer, String> buildColumnMap(Row row) {
        Map<Integer, String> modelMap = new HashMap<>();
        DataFormatter formatter = new DataFormatter();
        for (Cell cell : row) {
            modelMap.put(cell.getColumnIndex(), formatter.formatCellValue(cell).trim());
        }
        return modelMap;
    }

    private static Workbook buildWorkBook(String fileName) {
        if (fileName != null) {
            if (fileName.endsWith(SUFFIX_EXCEL_2003)) {
                return new HSSFWorkbook();
            } else if (fileName.endsWith(SUFFIX_EXCEL_2007)) {
                return  new XSSFWorkbook();
            }
        }
        return null;
    }

    private static Cell buildStyledCell(Row row, int cellIndex, CellStyle cellStyle) {
        Cell cell = row.createCell(cellIndex);
        cell.setCellStyle(cellStyle);
        return cell;
    }

    private static CellStyle buildHeaderCellStyle(Workbook workbook) {
        CellStyle cellStyle = buildCommonCellStyle(workbook);

        // 字体
        Font font = workbook.createFont();
        font.setFontName(DEFAULT_FONT);
        font.setFontHeightInPoints(HEADER_FONT_SIZE);
        font.setBold(true);
        font.setColor(IndexedColors.BLACK.getIndex());
        cellStyle.setFont(font);

        // 填充
        cellStyle.setFillForegroundColor(IndexedColors.SEA_GREEN.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        return cellStyle;
    }

    private static CellStyle buildValueCellStyle(Workbook workbook) {
        CellStyle cellStyle = buildCommonCellStyle(workbook);

        // 字体
        Font font = workbook.createFont();
        font.setFontName(DEFAULT_FONT);
        font.setFontHeightInPoints(VALUE_FONT_SIZE);
        font.setBold(false);
        font.setColor(IndexedColors.BLACK.getIndex());
        cellStyle.setFont(font);

        // 填充
        cellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());

        return cellStyle;
    }

    private static CellStyle buildCommonCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();

        // 对齐方式
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setWrapText(false); // 取消自动换行

        // 边框
        BorderStyle borderStyle = BorderStyle.THIN;
        cellStyle.setBorderTop(borderStyle);
        cellStyle.setBorderRight(borderStyle);
        cellStyle.setBorderBottom(borderStyle);
        cellStyle.setBorderLeft(borderStyle);
        short borderColor = IndexedColors.BLACK.getIndex();
        cellStyle.setTopBorderColor(borderColor);
        cellStyle.setRightBorderColor(borderColor);
        cellStyle.setBottomBorderColor(borderColor);
        cellStyle.setLeftBorderColor(borderColor);

        return cellStyle;
    }

    private static String getCellValue(Cell cell, String pattern) {
        String cellValue = "";
        if (cell == null) {
            return cellValue;
        }

        switch (cell.getCellType()) {
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    if (pattern == null) {
                        pattern = DEFAULT_DATE_PATTERN;
                    }
                    SimpleDateFormat sdf = new SimpleDateFormat(pattern);
                    Date date = cell.getDateCellValue();
                    cellValue = sdf.format(date);
                } else if (cell.getCellStyle().getDataFormat() == 0) { //处理数值格式
                    cell.setCellType(CellType.STRING);
                    cellValue = String.valueOf(cell.getRichStringCellValue().getString());
                }
                break;
            case STRING: // 字符串
                cellValue = String.valueOf(cell.getStringCellValue());
                break;
            case BOOLEAN: // Boolean
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case FORMULA: // 公式
                cellValue = String.valueOf(cell.getCellFormula());
                break;
            case BLANK: // 空值
                cellValue = null;
                break;
            case ERROR: // 故障
                cellValue = "非法字符";
                break;
            default:
                cellValue = "未知类型";
                break;
        }

        return cellValue;
    }

    /**
     * 根据值类型自动设置单元格的值
     * @param cell 单元格
     * @param value 单元格的值
     * @param pattern 格式
     */
    private static void setCellValue(Cell cell, Object value, String pattern) {
        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Float) {
            if (pattern == null) {
                pattern = DEFAULT_NUMERIC_PATTERN;
            }
            DecimalFormat df = new DecimalFormat(pattern);
            cell.setCellValue(df.format(value));
        } else if (value instanceof Double) {
            if (pattern == null) {
                pattern = DEFAULT_NUMERIC_PATTERN;
            }
            DecimalFormat df = new DecimalFormat(pattern);
            cell.setCellValue(df.format(value));
        } else if (value instanceof Long) {
            cell.setCellValue((Long) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Date) {
            if (pattern == null) {
                pattern = DEFAULT_DATE_PATTERN;
            }
            Date date = (Date) value;
            SimpleDateFormat df = new SimpleDateFormat(pattern);
            cell.setCellValue(df.format(date));
        } else if (value instanceof String[]) {
            StringBuilder stringBuilder = new StringBuilder();
            for (String s : (String []) value) {
                stringBuilder.append(s);
            }
            cell.setCellValue(stringBuilder.toString());
        } else if (value instanceof Double[]) {
            StringBuilder stringBuilder = new StringBuilder();
            for (Double d : (Double []) value) {
                stringBuilder.append(d);
            }
            cell.setCellValue(stringBuilder.toString());
        } else {
            cell.setCellValue((String) value);
        }
    }
}
