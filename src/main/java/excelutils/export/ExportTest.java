package excelutils.export;

import excelutils.ExcelCell;
import excelutils.ExcelExportSXSSF;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.function.Consumer;

public class ExportTest {
    public static void main(String[] args) {
        ExcelExport excelExport = new ExcelExport() {
            @Override
            public void setSheetList(ExcelExportSXSSF excelExportSXSSF, List<ExcelSheetExport> sheetList) {
                sheetList.add(new ExcelSheetExport<String>("测试sheet名称",excelExportSXSSF) {
                    @Override
                    public void setTitle(List<ExcelCell> row) {
                        Consumer<CellStyle> styleFunction = cellStyle -> {
                            Font font = excelExportSXSSF.createFont();
                            font.setBold(true);
                            cellStyle.setFont(font);
                            cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                        };
                        row.add(new ExcelCell("机构名称", null, null, BorderStyle.THIN, HorizontalAlignment.CENTER, VerticalAlignment.CENTER, null, false,18,(short)25,styleFunction));
                        row.add(new ExcelCell("机构编码", null, null, BorderStyle.THIN, HorizontalAlignment.CENTER, VerticalAlignment.CENTER, null, false,15,(short)25,styleFunction));
                        row.add(new ExcelCell("机构简称", null, null, BorderStyle.THIN, HorizontalAlignment.CENTER, VerticalAlignment.CENTER, null, false,15,(short)25,styleFunction));
                        row.add(new ExcelCell("上级机构名称", null, null, BorderStyle.THIN, HorizontalAlignment.CENTER, VerticalAlignment.CENTER, null, false,45,(short)25,styleFunction));
                        row.add(new ExcelCell("机构类型", null, null, BorderStyle.THIN, HorizontalAlignment.CENTER, VerticalAlignment.CENTER, null, false,30,(short)25,styleFunction));
                        row.add(new ExcelCell("机构性质", null, null, BorderStyle.THIN, HorizontalAlignment.CENTER, VerticalAlignment.CENTER, null, false,30,(short)25,styleFunction));
                    }

                    @Override
                    public List<String> getDataList() {
                        List<String> s = new ArrayList();
                        Collections.addAll(s, "1", "2", "3");
                        return s;
                    }

                    @Override
                    public void buildData(String data, List<ExcelCell> row) {
                        row.add(new ExcelCell("机构名称", null, null, BorderStyle.THIN, HorizontalAlignment.CENTER, VerticalAlignment.CENTER, null, false,null,null));
                        row.add(new ExcelCell(data, null, null, BorderStyle.THIN, HorizontalAlignment.CENTER, VerticalAlignment.CENTER, null, false,null,null));
                    }
                });
            }

            @Override
            public void buildCache(ExcelExportSXSSF excelExportSXSSF) {

            }
        };
    }
}
