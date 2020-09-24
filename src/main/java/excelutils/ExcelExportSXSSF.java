package excelutils;

import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;
import java.util.function.Consumer;
import java.util.stream.Collectors;

public class ExcelExportSXSSF{

    private LinkedHashMap<String, List<List<ExcelCell>>> dataMap=new LinkedHashMap<>(16);
    private SXSSFWorkbook workbook;
    private Map<Integer, Set<Integer>> indexMap=new HashMap<>(16);
    private Map<String,ExcelSheet> sheetMap = new HashMap<>(16);
    //用于存放缓存
    private Map<String, Object> cache = new HashMap<>();

    private List<Consumer> borderList = new ArrayList<>();
    public ExcelExportSXSSF(){
        this.dataMap = dataMap;
        workbook=new SXSSFWorkbook();
    }

    public class ExcelSheet{
        private ExcelSheet(){}
        List<List<ExcelCell>> dataList = new ArrayList<>(16);
        public ExcelSheet(String sheetName,LinkedHashMap<String, List<List<ExcelCell>>> dataMap) {
            dataMap.put(sheetName, dataList);
            sheetMap.put(sheetName, this);
        }
        public void remove(int index){dataList.remove(index);}

        public List<ExcelCell> createRow(){
            List<ExcelCell> cellList = new ArrayList<>(16);
            dataList.add(cellList);
            return cellList;
        }
        /** 获取列头 */
        public List<String> keyList(){
            return dataList.get(0).stream().map(ExcelCell::getValueStr).collect(Collectors.toList());
        }

    }

    public Object getCache(String key){return cache.get(key);}

    public void putCache(String key,Object value){cache.put(key,value);}

    public ExcelSheet getSheetDataList(String sheetName){
        return sheetMap.get(sheetName);
    }
    public ExcelSheet creatSheet(String sheetName){
        return new ExcelSheet(sheetName,dataMap);
    }

    public SXSSFWorkbook getWorkbook(){
        return workbook;
    }
    public Workbook buildWorkbook() {
        for(String sheetName  : dataMap.keySet()){
            List<List<ExcelCell>> sheetDataList = dataMap.get(sheetName);
            SXSSFSheet sheet = workbook.createSheet(sheetName);
            sheet.setRandomAccessWindowSize(-1);
            indexMap.clear();
            //行是一行一行固定递增
            for (int rowNumber = 0; rowNumber < sheetDataList.size(); rowNumber++) {
                SXSSFRow row = rowNumber!=0 && sheet.getLastRowNum()>=rowNumber ? sheet.getRow(rowNumber): sheet.createRow(rowNumber);
                List<ExcelCell> rowData = sheetDataList.get(rowNumber);
                //列可能会因为合并单元格二跳过
                for (int dataIndex=0; dataIndex < rowData.size(); dataIndex++) {
                    SXSSFCell cell = row.createCell(getCellNumber(rowNumber, dataIndex,rowData.get(dataIndex),sheet));
                    writeCellData(cell, rowData.get(dataIndex));
                }
            }
        }
        if(borderList.size()>0){
            borderList.forEach(consumer -> consumer.accept(null));
        }
        return workbook;
    }

    /**
     * 单个单元格写出方式
     * @param cell
     * @param excelCell
     */
    private void writeCellData(SXSSFCell cell, ExcelCell excelCell) {
        if (excelCell != null) {
            CellStyle cellStyle = workbook.createCellStyle();
            if(excelCell.getVerticalAlignment()!=null){
                cellStyle.setVerticalAlignment(excelCell.getVerticalAlignment());
            }
            if(excelCell.getHorizontalAlignment()!=null){
                cellStyle.setAlignment(excelCell.getHorizontalAlignment());
            }
            if(excelCell.getFormat()!=null){
                DataFormat format= workbook.createDataFormat();
                cellStyle.setDataFormat(format.getFormat(excelCell.getFormat()));
            }
            if(excelCell.getAutoWidth()){
                cell.getSheet().autoSizeColumn(cell.getColumnIndex());
            }else if(excelCell.getColWidth()!=null){
                cell.getSheet().setColumnWidth(cell.getColumnIndex(), 256*excelCell.getColWidth()+184);
            }
            if(excelCell.getRowHight()!=null){
                cell.getRow().setHeight((short)(20*excelCell.getRowHight()));
            }
            //合并单元格
            if(excelCell.getFloatCol()!=null||excelCell.getFloatRow()!=null){
                int rowIndex = cell.getRowIndex();
                int columnIndex = cell.getColumnIndex();
                CellRangeAddress userRegion = new CellRangeAddress(rowIndex,rowIndex+(excelCell.getFloatRow()==null?0:excelCell.getFloatRow()),columnIndex, columnIndex+(excelCell.getFloatCol()==null?0:excelCell.getFloatCol()));
                if(excelCell.getBorderStyle()!=null){
                    borderList.add(o -> {
                        RegionUtil.setBorderBottom(excelCell.getBorderStyle(), userRegion, cell.getSheet());
                        RegionUtil.setBorderLeft(excelCell.getBorderStyle(), userRegion, cell.getSheet());
                        RegionUtil.setBorderRight(excelCell.getBorderStyle(), userRegion, cell.getSheet());
                        RegionUtil.setBorderTop(excelCell.getBorderStyle(), userRegion, cell.getSheet());
                    });
                }
                cell.getSheet().addMergedRegion(userRegion);
            }else if(excelCell.getBorderStyle()!=null){
                cellStyle.setBorderTop(excelCell.getBorderStyle());
                cellStyle.setBorderBottom(excelCell.getBorderStyle());
                cellStyle.setBorderLeft(excelCell.getBorderStyle());
                cellStyle.setBorderRight(excelCell.getBorderStyle());
            }
            if (excelCell.getStyleFunction() != null) {
                excelCell.getStyleFunction().accept(cellStyle);
            }
            cell.setCellStyle(cellStyle);
            if (excelCell.getValue() != null) {
                if(excelCell.getValue().getClass().equals(String.class)){
                    RichTextString str = new HSSFRichTextString((String)excelCell.getValue());
                    cell.setCellValue(str);
                }else if(excelCell.getValue().getClass().equals(Integer.class)){
                    cell.setCellValue((Integer)excelCell.getValue());
                }else if(excelCell.getValue().getClass().equals(Date.class)){
                    cell.setCellValue((Date)excelCell.getValue());
                }else if(excelCell.getValue().getClass().equals(Double.class)){
                    cell.setCellValue((Double)excelCell.getValue());
                }
            }
        }
    }

    /**
     * 获取当前cell的起始位置
     * @param rowNumber
     * @param dataIndex
     * @param rowData
     * @return
     */
    private int getCellNumber(Integer rowNumber, Integer dataIndex, ExcelCell rowData, SXSSFSheet sheet){
        //列位置可能会因为合并单元格而向后平移
        for(;;dataIndex++ ){
            if(!indexMap.containsKey(rowNumber) || !indexMap.get(rowNumber).contains(dataIndex)){
                Set<Integer> cellNumSet = indexMap.containsKey(rowNumber) ? indexMap.get(rowNumber) : new HashSet<>();
                cellNumSet.add(dataIndex);
                addFloatCell(rowData,cellNumSet,dataIndex);
                indexMap.put(rowNumber, cellNumSet);
                if(rowData.getFloatRow()!=null){
                    for(int fr=1;fr<=rowData.getFloatRow();fr++){
                        Integer fri=rowNumber+fr;
                        sheet.createRow(fri);
                        Set<Integer> cellNumSetfr = indexMap.containsKey(fri) ? indexMap.get(fri) : new HashSet<>();
                        cellNumSetfr.add(dataIndex);
                        addFloatCell(rowData,cellNumSetfr,dataIndex);
                        indexMap.put(fri, cellNumSetfr);
                    }
                }
                return dataIndex;
            }
        }
    }
    private void addFloatCell(ExcelCell rowData,Set<Integer> cellNumSet,Integer dataIndex){
        if(rowData.getFloatCol()!=null){
            for(int f=1;f<=rowData.getFloatCol();f++){
                cellNumSet.add(dataIndex+f);
            }
        }
    }


    public Font createFont(){
        return workbook.createFont();
    }

    public static void main(String[]strs) throws Exception {
        ExcelExportSXSSF excelExportSXSSF = new ExcelExportSXSSF();
        ExcelSheet sheet1 = excelExportSXSSF.creatSheet("sheet1");
        List<ExcelCell> row1 = sheet1.createRow();

        ExcelCell excelCell = new ExcelCell("行列合并测试", 1, 1, BorderStyle.THIN, HorizontalAlignment.LEFT, VerticalAlignment.CENTER, null, false,null,null);
        row1.add(excelCell);
        row1.add( new ExcelCell("一般属性", null, null, BorderStyle.THIN, HorizontalAlignment.CENTER, VerticalAlignment.CENTER, null, false,null,Short.valueOf("100")));
        row1.add( new ExcelCell("一般属性", null, null, BorderStyle.THIN, HorizontalAlignment.RIGHT, VerticalAlignment.CENTER, null, false,null,Short.valueOf("100")));

        List<ExcelCell> row2 = sheet1.createRow();
        row2.add( new ExcelCell("一般属性", null, null, BorderStyle.THIN, HorizontalAlignment.CENTER, VerticalAlignment.CENTER, null, false,null,Short.valueOf("100")));

        row2.add(new ExcelCell("一般属性", null, null, BorderStyle.THIN, HorizontalAlignment.CENTER, VerticalAlignment.CENTER, null, false,null,Short.valueOf("100"), cellStyle -> {
            Font font = excelExportSXSSF.createFont();
            font.setFontName("楷体");
            font.setBold(true);
            font.setFontHeightInPoints((short) 14);
            cellStyle.setFont(font);
        }));
        //controller下载 ExcelDownload.download(excelExportSXSSF,fileName,response,request);
        excelExportSXSSF.buildWorkbook().write(new FileOutputStream(new File("D:\\demo.xlsx")));
    }
}