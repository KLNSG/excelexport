package excelutils.export;

import excelutils.ExcelDownload;
import excelutils.ExcelExportSXSSF;

import java.util.ArrayList;
import java.util.List;

/**
 * excel导出工具类
 */
public abstract class ExcelExport{
    private ExcelExportSXSSF excelExportSXSSF = new ExcelExportSXSSF();
    private List<ExcelSheetExport> sheetList = new ArrayList<>();

    /** 设置一些通用缓存类数据  在构建的 时候可以获取；excelExportSXSSF.putCache */
    public abstract void buildCache(ExcelExportSXSSF excelExportSXSSF);
    /** 添加sheet数据 */
    public abstract void setSheetList(ExcelExportSXSSF excelExportSXSSF,List<ExcelSheetExport> sheetList);

    public ExcelExport(){
        buildCache(excelExportSXSSF);
        setSheetList(excelExportSXSSF,sheetList);
    }

    /**
     * controller下载
     * @param fileName 文件名，不需要后缀
     * @param response
     * @param request
     */
    public void download(String fileName, HttpServletResponse response, HttpServletRequest request) throws Exception {
        ExcelDownload.download(excelExportSXSSF,fileName,response,request);
    }

}
