package excelutils;

import com.goldgov.kduck.dao.ParamMap;
import com.goldgov.kduck.service.ValueMapList;
import com.goldgov.kduck.utils.ValueMapUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.StringUtils;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;


public class ExcelUtilsRead {

    public static final String SHEETNAME = "excelSheetName";

    public static final String ROWNUMBER = "excelRowNumber";

    /**
     * 将Excel文件解析为Map结构
     * @param inputStream
     * @return
     */
    public static Map<String, List<Map<String, Object>>> excelConversionMap(InputStream inputStream, String version) {
        Map<String, List<Map<String, Object>>> excelMap = new HashMap<>(16);
        version= version == null ? "07" : version;
        try {
            Workbook workbook = null;
            if ("07".equals(version)) {
                workbook = new XSSFWorkbook(inputStream);
            } else {
                workbook = new HSSFWorkbook();
            }
            int sheetSize = workbook.getNumberOfSheets();
            //处理 sheet
            for (int i = 0; i < sheetSize; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                List<Map<String, Object>> sheetMap = new ArrayList<>();
                int rowSize = sheet.getLastRowNum()+1;
                int startRow = 0;
                //处理标题行
                List<String> keys = new ArrayList<>();
                for (int j = 0; j < rowSize; j++) {
                    Row row = sheet.getRow(j);
                    int cellSize = row.getLastCellNum();
                    boolean flag = false;
                    for (int x = 0; x < cellSize; x++) {
                        //处理本行内是否有合并单元格
                        Result mergedRegion = isMergedRegion(sheet, row.getRowNum(), x);
                        if (mergedRegion.merged) {
                            flag = true;
                        }
                    }
                    if (flag) {
                        continue;
                    }
                    //将标题行 设置到keys中
                    for (int k = 0; k < cellSize; k++) {
                        Cell cell = row.getCell(k);
                        if (cell != null) {
                            try{
                                keys.add(cell.getStringCellValue());
                            }catch (Exception e){
                                throw new RuntimeException(e);
                            }
                        }
                    }
                    startRow = j + 1;
                    break;
                }
                //处理每一行数据,封装为Map 放入List
                st:
                for (int j = startRow; j < rowSize; j++) {
                    Row row = sheet.getRow(j);
                    if(row == null|| row.getPhysicalNumberOfCells()<3){
                        continue st;
                    }
                    int cellSize = row.getLastCellNum();
                    Map<String, Object> data = new HashMap<>(32);
                    for (int k = 0; k < cellSize; k++) {
                        Cell cell = row.getCell(k);
                        if (cell == null) {
                            continue;
                        }
                        data.put(keys.get(k), getCellValue(cell,workbook));
                    }
                    //保留位置信息,方便提示
                    data.put(SHEETNAME, sheet.getSheetName());
                    data.put(ROWNUMBER, row.getRowNum() + 1);
                    sheetMap.add(data);
                }
                excelMap.put(sheet.getSheetName(), sheetMap);
            }
            return excelMap;
        } catch (IOException e) {
            throw new RuntimeException("解析文件异常！");
        }
    }

    public static Object getCellValue(Cell cell, Workbook workbook) {
        if (cell != null) {
            CellType type = cell.getCellType();
            switch (type) {
                case STRING:
                    return cell.getStringCellValue();
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue();
                    } else {
                        return cell.getNumericCellValue();
                    }
                case BOOLEAN:
                    return cell.getBooleanCellValue();
                case FORMULA:
                    //String cellFormula = cell.getCellFormula(); 获取公式
                    XSSFFormulaEvaluator evaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);
                    CellValue evaluate = evaluator.evaluate(cell);
                    CellType cellType = evaluate.getCellType();
                    switch (cellType) {
                        case STRING:
                            return evaluate.getStringValue();
                        case NUMERIC:
                            return evaluate.getNumberValue();
                        case BOOLEAN:
                            return evaluate.getBooleanValue();
                        case BLANK:
                        case ERROR:
                            return evaluate.getErrorValue();
                        default:
                            break;
                    }
                    break;
                case BLANK:
                    break;
                case ERROR:
                    return cell.getErrorCellValue();
                default:
                    break;
            }
        }
        return null;
    }
    /**
     * 将数据解析为workBook
     *
     * @param sheetName sheet名称
     * @param list      数据集
     * @param keys      对象Key  表头  name : 姓名   age : 年龄
     * @return Excel对象
     */
    public static Workbook dataConversionStream(String sheetName, ValueMapList list, LinkedHashMap<String, String> keys) {
        Workbook workbook = new SXSSFWorkbook();
        Sheet sheet = workbook.createSheet(sheetName);
        Row firstRow = sheet.createRow(0);
        int x = 0;
        for (String value : keys.values()) {
            Cell cell = firstRow.createCell(x);
            cell.setCellValue(value);
            x++;
        }
        //生成数据
        for (int i = 0; i < list.size(); i++) {
            Row row = sheet.createRow(i + 1);
            int j = 0;
            for (Map.Entry<String, String> node : keys.entrySet()) {
                Cell cell = row.createCell(j);
                cell.setCellValue(list.get(i).getValueAsString(node.getKey()));
                j++;
            }
        }
         /*OutputStream bos = new ByteArrayOutputStream();
         workbook.write(bos);*/
        return workbook;
    }

    public static String getValueAsString(Object object) {
        return ValueMapUtils.getValueAsString(ParamMap.create("key",object).toMap(),"key");
    }

    /**
     *
     * @param map
     * @param key
     * @param mandatory 是否必填
     * @param parentCode
     * @param errorList
     * @return
     */
    public static String getValueAsString(Map<String, Object> map, String key, boolean mandatory, String parentCode, List<String> errorList) {
        if (isNotNull(map, key, mandatory, errorList)) {
            try {
                return ValueMapUtils.getValueAsString(getValue(map, key, mandatory, parentCode, errorList), key);
            } catch (Exception e) {
                errorList.add(map.get(SHEETNAME) + "第" + map.get(ROWNUMBER) + "行" + key + "格式不正确");
            }
        }
        return null;
    }

    public static Integer getValueAsInteger(Map<String, Object> map, String key, boolean mandatory, String parentCode, List<String> errorList) {
        if (isNotNull(map, key, mandatory, errorList)) {
            try {
                return ValueMapUtils.getValueAsInteger(getValue(map, key, mandatory, parentCode, errorList), key);
            } catch (Exception e) {
                errorList.add(map.get(SHEETNAME) + "第" + map.get(ROWNUMBER) + "行" + key + "格式不正确");
            }
        }
        return null;
    }

    public static Date getDateValue(Map<String, Object> map, String key, boolean mandatory, String parentCode, List<String> errorList) {
        if (isNotNull(map, key, mandatory, errorList)) {
            //格式转换
            try {
                return ValueMapUtils.getValueAsDate(map, key);
            } catch (Exception e) {
                errorList.add(map.get(SHEETNAME) + "第" + map.get(ROWNUMBER) + "行" + key + "格式不正确");
            }
        }
        return null;
    }

    private static Map<String, Object> getValue(Map<String, Object> map, String key, boolean mandatory, String parentCode, List<String> errorList) {
        if (parentCode != null) {
            //TODO 基础数据转换
            return null;
        } else {
            return map;
        }
    }

    private static boolean isNotNull(Map<String, Object> map, String key, boolean mandatory, List<String> errorList) {
        if (mandatory) {
            if (StringUtils.isEmpty(ValueMapUtils.getValueAsString(map,key))) {
                errorList.add(map.get(SHEETNAME) + "第" + map.get(ROWNUMBER) + "行" + key + "不能为空");
                return false;
            }
        }
        return true;
    }

    /**
     * 校验指定位置的单元格是否是一个合并过的单元格
     *
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    private static Result isMergedRegion(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return new Result(true, firstRow + 1, lastRow + 1, firstColumn + 1, lastColumn + 1);
                }
            }
        }
        return new Result(false, 0, 0, 0, 0);
    }
}

class Result {
    public boolean merged;
    public int startRow;
    public int endRow;
    public int startCol;
    public int endCol;

    public Result(boolean merged, int startRow, int endRow, int startCol, int endCol) {
        this.merged = merged;
        this.startRow = startRow;
        this.endRow = endRow;
        this.startCol = startCol;
        this.endCol = endCol;
    }
}
