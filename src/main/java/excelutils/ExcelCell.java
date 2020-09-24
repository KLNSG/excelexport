package excelutils;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.util.function.Consumer;


public class ExcelCell {
    private Object value;
    private Integer floatRow;
    private Integer floatCol;
    private BorderStyle borderStyle;
    private HorizontalAlignment horizontalAlignment;
    private VerticalAlignment verticalAlignment;
    private String format;
    private Boolean autoWidth;
    private Integer colWidth;
    private Short rowHight;
    private Consumer<CellStyle> styleFunction;

    /**
     * 构造单元格对象 默认 带边框 居中
     * @param value 值
     */
    public ExcelCell(Object value){
        this.value=value;
        this.borderStyle= BorderStyle.THIN;
        this.horizontalAlignment= HorizontalAlignment.CENTER;
        this.verticalAlignment= VerticalAlignment.CENTER;
    }

    /**
     * 构造单元格对象
     * @param value 值
     * @param floatRow 向下合并行数（最小1）不合并则传null
     * @param floatCol 向右合并列数（最小1）不合并则传null
     * @param borderStyle 边框样式
     * @param horizontalAlignment 水平对齐
     * @param verticalAlignment 垂直对齐
     * @param format 字段格式化 参考excel 格式例如 "¥#.##00"
     * @param autoWidth 是否自动宽度
     * @param colWidth 宽 数值参考excel设置
     * @param rowHight 高 数值参考excel设置
     */
    public ExcelCell(Object value, Integer floatRow, Integer floatCol, BorderStyle borderStyle, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment, String format, Boolean autoWidth, Integer colWidth, Short rowHight){
        this.value=value;
        this.floatRow=floatRow;
        this.floatCol=floatCol;
        this.borderStyle=borderStyle;
        this.horizontalAlignment=horizontalAlignment;
        this.verticalAlignment=verticalAlignment;
        this.format = format;
        this.autoWidth = autoWidth;
        this.colWidth=colWidth;
        this.rowHight=rowHight;
    }
    /**
     * 构造单元格对象
     * @param value 值
     * @param floatRow 向下合并行数（最小1）不合并则传null
     * @param floatCol 向右合并列数（最小1）不合并则传null
     * @param borderStyle 边框样式
     * @param horizontalAlignment 水平对齐
     * @param verticalAlignment 垂直对齐
     * @param format 字段格式化 参考excel 格式例如 "¥#.##00"
     * @param autoWidth 是否自动宽度
     * @param colWidth 宽 数值参考excel设置
     * @param rowHight 高 数值参考excel设置
     * @param styleFunction 自定义单元格格式
     */
    public ExcelCell(Object value, Integer floatRow, Integer floatCol, BorderStyle borderStyle, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment, String format, Boolean autoWidth, Integer colWidth, Short rowHight, Consumer<CellStyle> styleFunction){
        this.value=value;
        this.floatRow=floatRow;
        this.floatCol=floatCol;
        this.borderStyle=borderStyle;
        this.horizontalAlignment=horizontalAlignment;
        this.verticalAlignment=verticalAlignment;
        this.format = format;
        this.autoWidth = autoWidth;
        this.colWidth=colWidth;
        this.rowHight=rowHight;
        this.styleFunction = styleFunction;
    }

    public Consumer<CellStyle> getStyleFunction() {
        return styleFunction;
    }


    public Object getValue() {
        return value;
    }
    public String getValueStr(){return value.toString();}

    public Integer getFloatRow() {
        return floatRow;
    }

    public Integer getFloatCol() {
        return floatCol;
    }


    public BorderStyle getBorderStyle() {
        return borderStyle;
    }


    public HorizontalAlignment getHorizontalAlignment() {
        return horizontalAlignment;
    }

    public VerticalAlignment getVerticalAlignment() {
        return verticalAlignment;
    }


    public String getFormat() {
        return format;
    }


    public Boolean getAutoWidth() {
        return autoWidth!=null&&autoWidth;
    }

    public Integer getColWidth() {
        return colWidth;
    }

    public Short getRowHight() {
        return rowHight;
    }

    public ExcelCell setColWidth(Integer colWidth) {
        this.colWidth = colWidth;
        return this;
    }

    public ExcelCell setValue(Object value) {
        this.value = value;
        return this;
    }

    public ExcelCell setStyleFunction(Consumer<CellStyle> styleFunction) {
        this.styleFunction = styleFunction;
        return this;
    }

    public ExcelCell setBorderStyle(BorderStyle borderStyle) {
        this.borderStyle = borderStyle;
        return this;
    }

    public ExcelCell setHorizontalAlignment(HorizontalAlignment horizontalAlignment) {
        this.horizontalAlignment = horizontalAlignment;
        return this;
    }

    public ExcelCell setVerticalAlignment(VerticalAlignment verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
        return this;
    }
}
