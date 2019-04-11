package ec.comun.reporte.util;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;

public class EstiloExcelUtil {
    private Map<CellStyle, XSSFCellStyle> ESTILO_LIBRO = new HashMap<>();
    public static final String ARGB_PETRO = "FF044866";
    private XSSFWorkbook workbook;
    private XSSFCellStyle cellStyle;
    private XSSFFont font;

    private static EstiloExcelUtil instancia;

    private EstiloExcelUtil() {

    }

    public static final synchronized EstiloExcelUtil getInstancia(XSSFWorkbook workbook) {
        if (instancia == null) {
            instancia = new EstiloExcelUtil();
        }
        instancia.iniciarEstiloCelda(workbook);
        return instancia;
    }

    private void iniciarEstiloCelda(XSSFWorkbook workbook) {
        this.workbook = workbook;
        this.cellStyle = this.workbook.createCellStyle();
        this.font = this.workbook.createFont();

    }

    public EstiloExcelUtil aplicarNegrita() {
        this.font.setBold(true);
        return this;
    }

    public EstiloExcelUtil alinearHorizontalmente(HorizontalAlignment alignment) {
        this.cellStyle.setAlignment(alignment);
        return this;
    }

    public EstiloExcelUtil alinearVerticalmente(VerticalAlignment alignment) {
        this.cellStyle.setVerticalAlignment(alignment);
        return this;
    }

    public EstiloExcelUtil aplicarTamanioTexto(short tamanioTexto) {
        this.font.setFontHeightInPoints(tamanioTexto);
        return this;
    }

    public EstiloExcelUtil aplicarColorTexto(String hexColor) {
        XSSFColor color = obtenerXSSFColor(hexColor);
        this.font.setColor(color);
        return this;
    }

    public EstiloExcelUtil aplicarBordes(BorderStyle border) {
        this.cellStyle.setBorderBottom(border);
        this.cellStyle.setBorderLeft(border);
        this.cellStyle.setBorderRight(border);
        this.cellStyle.setBorderTop(border);
        return this;
    }

    public EstiloExcelUtil aplicarBordes(String hexColor, BorderStyle border) {
        aplicarBordes(border);
        XSSFColor color = obtenerXSSFColor(hexColor);
        this.cellStyle.setBorderColor(BorderSide.BOTTOM, color);
        this.cellStyle.setBorderColor(BorderSide.LEFT, color);
        this.cellStyle.setBorderColor(BorderSide.TOP, color);
        this.cellStyle.setBorderColor(BorderSide.RIGHT, color);
        return this;
    }

    public EstiloExcelUtil aplicarColorFondo(String hexColor) {
        XSSFColor color = obtenerXSSFColor(hexColor);
        this.cellStyle.setFillBackgroundColor(color);
        this.cellStyle.setFillForegroundColor(color);
        return this;
    }

    public EstiloExcelUtil aplicarPatroColorFondo(FillPatternType pattern) {
        this.cellStyle.setFillPattern(pattern);
        return this;
    }

    public EstiloExcelUtil aplicarFormato(String formato) {
        final short index = workbook.createDataFormat().getFormat(formato);
        this.cellStyle.setDataFormat(index);
        return this;
    }

    public EstiloExcelUtil aplicarWrapTexto() {
        this.cellStyle.setWrapText(true);
        return this;
    }

    public XSSFColor obtenerXSSFColor(String hexColor) {
        XSSFColor color = workbook.getCreationHelper().createExtendedColor();
        color.setARGBHex(hexColor);
        return color;
    }

    public XSSFCellStyle obtenerEstilo() {
        this.cellStyle.setFont(this.font);
        if (ESTILO_LIBRO.containsKey(cellStyle)) {
            return ESTILO_LIBRO.get(cellStyle);
        } else {
            ESTILO_LIBRO.put(cellStyle, cellStyle);
        }
        return this.cellStyle;
    }

}
