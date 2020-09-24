package excelutils;

import com.goldgov.kduck.service.ValueMap;
import com.goldgov.kduck.service.ValueMapList;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class ExcelDownload {

    /**
     * 下载excel方法
     * @param excelExportSXSSF 导出构建类
     * @param fileName 不需要加后缀
     * @param response response
     * @param request request
     * @throws Exception
     */
    public static void download(ExcelExportSXSSF excelExportSXSSF, String fileName, HttpServletResponse response, HttpServletRequest request) throws Exception {
        excelExportSXSSF.buildWorkbook();
        setHeader(fileName,response,request);
        excelExportSXSSF.getWorkbook().write(response.getOutputStream());
    }

    /**
     * 简易版本 excel下载
     * @param exportFieldMap key 为要导出数据的属性   value {"name":"党组织名称","property":"orgName","width":20}
     * @param valueMapList 要导出的数据 key value
     * @param fileName 不带后缀
     * @param request request
     * @param response response
     * @throws Exception
     */
    public void download(Map<String, ValueMap> exportFieldMap, ValueMapList valueMapList, String fileName, HttpServletRequest request, HttpServletResponse response)throws Exception{
        ExcelExportSXSSF excelExportSXSSF = new ExcelExportSXSSF();
        ExcelExportSXSSF.ExcelSheet excelSheet = excelExportSXSSF.creatSheet("列表");


        List<ExcelCell> listHead = excelSheet.createRow();
        for (String field : exportFieldMap.keySet()) {
            Integer width = exportFieldMap.get(field).getValueAsInt("width");
            listHead.add(new ExcelCell(exportFieldMap.get(field).getValueAsString("name"),null,null, BorderStyle.THIN, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,null,width==null,width,null));
        }
        if (valueMapList != null && valueMapList.size() > 0) {
            for (ValueMap valueMap : valueMapList) {
                List<ExcelCell> listrow = excelSheet.createRow();
                for (String field : exportFieldMap.keySet()) {
                    Integer width = exportFieldMap.get(field).getValueAsInt("width");
                    listrow.add(new ExcelCell(valueMap.getValueAsString(field),null,null, BorderStyle.THIN, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,null,width==null,width,null));
                }
            }
        }
        this.download(excelExportSXSSF,fileName,response,request);
    }

    public static void setHeader(String fileName, HttpServletResponse response, HttpServletRequest request) throws Exception {
        response.setContentType("application/octet-stream");
        if(fileName.indexOf(".")==-1){
            fileName = fileName + ".xlsx";
        }
        if("FF".equals(getBrowser(request))){
            //如果是火狐,解决火狐中文名乱码问题
            response.setHeader("Content-Disposition", "attachment;filename="+ new String(fileName.getBytes("UTF-8"),"iso-8859-1"));
        }else{
            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
        }
    }

    public static String getBrowser(HttpServletRequest request) {
        String userAgent = request.getHeader("USER-AGENT");
        if (userAgent != null) {
            userAgent = userAgent.toLowerCase();
            if (userAgent.indexOf("msie") >= 0)
                return "IE";
            if (userAgent.indexOf("firefox") >= 0)
                return "FF";
            if (userAgent.indexOf("safari") >= 0)
                return "SF";
        }
        return null;
    }

    /**
     * 导入excel 通用
     * @author Lvxin
     **/
    public static List<Account> importExcel(InputStream inputStream) throws IOException, InvalidFormatException {
        List<Account> accounts=new ArrayList<>();
        Workbook hssfWorkbook = WorkbookFactory.create(inputStream);//将输入流对象存到工作簿对象里面
        // 循环工作表Sheet
        for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
            Sheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
            if (hssfSheet == null) {
                continue;
            }

            // 循环行Row
            for (int rowNum = 1; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
                Row hssfRow = hssfSheet.getRow(rowNum);
                if (hssfRow == null) {
                    continue;
                }
                String accountName= getValue(hssfRow.getCell(2));
                if (accountName==null || "".equals(accountName)){
                    continue;
                }
                Account account=new Account();
                // 循环列Cell
                /*	account.setAccountId(getValue(hssfRow.getCell(0)));*/
                account.setDisplayName(getValue(hssfRow.getCell(1)));
                account.setAccountName(getValue(hssfRow.getCell(2)));
                account.setPhone(getValue(hssfRow.getCell(3)));
                account.setEmail(getValue(hssfRow.getCell(4)));
                account.setAccountState(getValue(hssfRow.getCell(5)).equals("启用")?1:0);
                account.setRelationName(getValue(hssfRow.getCell(6)));
                accounts.add(account);
            }
        }
        return accounts;
    }

    /**
     * 得到Excel表中的值
     *
     * @param hssfCell
     *            Excel中的每一个格子
     * @return Excel中每一个格子中的值
     */
    @SuppressWarnings("static-access")
    private static String getValue(Cell hssfCell) {
        if (hssfCell.getCellType() == CellType.BOOLEAN) {
            // 返回布尔类型的值
            return String.valueOf(hssfCell.getBooleanCellValue());
        } else if (hssfCell.getCellType() == CellType.NUMERIC) {
            // 返回数值类型的值
            return String.valueOf((int) hssfCell.getNumericCellValue());
        } else {
            // 返回字符串类型的值
            return String.valueOf(hssfCell.getStringCellValue());
        }

    }


}
