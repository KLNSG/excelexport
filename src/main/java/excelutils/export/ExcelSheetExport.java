package excelutils.export;


import excelutils.ExcelCell;
import excelutils.ExcelExportSXSSF;

import java.util.List;

/**
 * 导出exceSheet
 */
public abstract class ExcelSheetExport <T>{
    private ExcelExportSXSSF excelExportSXSSF;
    private ExcelExportSXSSF.ExcelSheet excelSheet;

    public ExcelSheetExport(String sheetName, ExcelExportSXSSF excelExportSXSSF) {
        this.excelExportSXSSF = excelExportSXSSF;
        excelSheet = this.excelExportSXSSF.creatSheet(sheetName);

        List<ExcelCell> titleRow = excelSheet.createRow();
        setTitle(titleRow);
        if(titleRow.size()==0){
            excelSheet.remove(0);
        }

        getDataList().forEach(t -> {
            List<ExcelCell> row = excelSheet.createRow();
            buildData(t,row);
        });
    }
    public abstract  void setTitle(List<ExcelCell> row);

    public abstract List<T> getDataList();

    public abstract void buildData(T data, List<ExcelCell> row);

}
