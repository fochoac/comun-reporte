package ec.comun.reporte.excel;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.helpers.HeaderFooterHelper;

import ec.comun.reporte.modelo.CeldaTo;
import ec.comun.reporte.modelo.FilaTo;
import ec.comun.reporte.modelo.HojaExcelTo;
import ec.comun.reporte.modelo.LibroExcelTo;
import ec.comun.reporte.util.EstiloExcelUtil;

public class ReporteExcel {
    private static final SpreadsheetVersion VERSION = SpreadsheetVersion.EXCEL2007;
    private static int NUMERO_FILA = 0;
    private static final String FORMATO_FECHA_NOMBRE_EXCEL = "ddMMyyyyHHmmssss";
    private static final String EXTENSION_EXCEL = ".xlsx";
    private static ReporteExcel instancia;
    private LibroExcelTo libroExcelTo;
    private XSSFWorkbook workbook;
    private Map<String, XSSFCellStyle> estilos = new HashMap<>();

    private ReporteExcel() {
        super();
    }

    public static synchronized ReporteExcel getInstancia(LibroExcelTo libroExcelTo) {

        if (instancia == null) {
            instancia = new ReporteExcel();
        }
        instancia.libroExcelTo = libroExcelTo;
        instancia.workbook = new XSSFWorkbook();
        return instancia;
    }

    public ReporteExcel construir() {
        validar();
        generarEstilosCeldas();
        for (HojaExcelTo hoja : libroExcelTo.getHojas()) {
            NUMERO_FILA = 0;
            Sheet sheet = workbook.createSheet(hoja.getNombre());
            armarFilas(sheet, hoja.getContenido());
            resizarTamanioColumnas(sheet, getNumeroCeldasMaximo(hoja));
            aplicarEstilosHoja(hoja, sheet);
            generarFooter(sheet, hoja.getTextoCabecera(), hoja.getTamanioLetraCabecera(), hoja.isNegritaCabecera());
        }

        return this;
    }

    private void generarEstilosCeldas() {

        for (HojaExcelTo hoja : libroExcelTo.getHojas()) {
            for (FilaTo fila : hoja.getContenido()) {
                for (CeldaTo celda : fila.getCeldas()) {
                    if (estilos.containsKey(celda.getEstilo())) {

                    } else {
                        estilos.put(celda.getEstilo(), obtenerEstiloCelda(celda));
                    }

                }

            }
        }
    }

    private int getNumeroCeldasMaximo(HojaExcelTo hoja) {
        int numero = 0;
        for (FilaTo fila : hoja.getContenido()) {
            if (fila.getCeldas().size() >= numero) {
                numero = fila.getCeldas().size();
            }
        }
        return numero;
    }

    public byte[] exportarArray() throws IOException {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        workbook.write(os);
        workbook.close();
        return os.toByteArray();
    }

    public void exportarArchivo(File archivo) throws FileNotFoundException, IOException {
        try (FileOutputStream os = new FileOutputStream(archivo)) {
            workbook.write(os);
            workbook.close();
        }
    }

    public File exportarArchivo() throws IOException {
        String nombre = "";
        SimpleDateFormat format = new SimpleDateFormat(FORMATO_FECHA_NOMBRE_EXCEL);
        if (this.libroExcelTo.getNombreArchivo() == null
                || this.libroExcelTo.getNombreArchivo().isEmpty()) {
            nombre = format.format(new Date());
        } else {
            nombre = this.libroExcelTo.getNombreArchivo()
                    + "-"
                    + format.format(new Date());
        }
        File file = File.createTempFile(nombre, EXTENSION_EXCEL);
        try (FileOutputStream os = new FileOutputStream(file)) {
            workbook.write(os);
            workbook.close();
        }
        return file;

    }

    private void armarFilas(Sheet hoja, List<FilaTo> filas) {
        if (filas == null
                || filas.isEmpty()) {
            return;
        }
        Collections.sort(filas);
        for (FilaTo fila : filas) {
            Row filaExcel = hoja.createRow(NUMERO_FILA);
            armarFila(filaExcel, fila);
            if (fila.getNumeroFilasUnir() > 0) {

                fusionarFilas(filaExcel, fila);

            }
            NUMERO_FILA++;
        }
    }

    private void fusionarCeldas(Cell celdaExcel, CeldaTo celda, int row) {
        final Sheet sheet = celdaExcel.getSheet();
        int firstRow = row;
        int lastRow = firstRow
                + celda.getNumeroFilasUnir();
        int firstCol = celdaExcel.getColumnIndex();
        int lastCol = firstCol
                + celda.getNumeroColumnasUnir();
        CellRangeAddress ca = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        sheet.addMergedRegion(ca);

    }

    private void fusionarFilas(Row filaExcel, FilaTo fila) {
        final Sheet sheet = filaExcel.getSheet();
        int firstRow = NUMERO_FILA;
        int lastRow = firstRow
                + fila.getNumeroFilasUnir();

        int firstCol = filaExcel.getFirstCellNum();
        int lastCol = filaExcel.getLastCellNum();

        CellRangeAddress ca = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        sheet.addMergedRegion(ca);
        NUMERO_FILA = lastRow;
    }

    private void armarFila(Row filaExcel, FilaTo datosFila) {
        List<CeldaTo> celdas = datosFila.getCeldas();
        int filas = 0;
        Collections.sort(celdas);
        for (CeldaTo celda : celdas) {
            Cell celdaExcel = filaExcel.createCell(celda.getIndice());
            if (celda.getDato() instanceof Number) {
                Number numero = (Number) celda.getDato();
                celdaExcel.setCellValue((double) numero);
                celdaExcel.setCellStyle(estilos.get(celda.getEstilo()));

            } else if (celda.getDato() instanceof Date) {
                Date fecha = (Date) celda.getDato();
                celdaExcel.setCellValue(fecha);
                celdaExcel.setCellStyle(estilos.get(celda.getEstilo()));
            } else if (celda.getDato() instanceof Calendar) {
                Calendar fecha = (Calendar) celda.getDato();
                celdaExcel.setCellValue(fecha);
                celdaExcel.setCellStyle(estilos.get(celda.getEstilo()));
            } else if (celda.getDato() instanceof String) {

                if (celda.isTextoEnriquecido()) {
                    CreationHelper createHelper = workbook.getCreationHelper();
                    RichTextString texto = createHelper.createRichTextString(celda.getDato().toString());
                    celdaExcel.setCellValue(texto);
                } else {
                    celdaExcel.setCellValue(celda.getDato().toString());
                }
                celdaExcel.setCellStyle(estilos.get(celda.getEstilo()));
            }
            if (celda.getNumeroColumnasUnir() > 0
                    || celda.getNumeroFilasUnir() > 0) {
                if (celda.getNumeroFilasUnir() > filas) {
                    filas = celda.getNumeroFilasUnir();
                }
                fusionarCeldas(celdaExcel, celda, NUMERO_FILA);
            }

        }
        NUMERO_FILA += filas;

    }

    private void resizarTamanioColumnas(Sheet hojaExcel, int numeroColumnas) {

        for (int i = 0; i < numeroColumnas; i++) {

            hojaExcel.autoSizeColumn(i, true);

        }
    }

    private void aplicarEstilosHoja(HojaExcelTo hojaExcelTo, Sheet sheet) {
        sheet.setDisplayRowColHeadings(true);
        if (hojaExcelTo.isFitToPage()) {
            sheet.setFitToPage(true);
        } else {
            sheet.setFitToPage(false);
        }
        if (hojaExcelTo.isImprimirGrilla()) {
            sheet.setPrintGridlines(true);
        } else {
            sheet.setPrintGridlines(false);
        }
        if (hojaExcelTo.isMostrarGrilla()) {
            sheet.setDisplayGridlines(true);
        } else {
            sheet.setDisplayGridlines(false);
        }

    }

    private void generarFooter(Sheet sheet, String textoCabecera, Short tamanioTexto, boolean negrita) {
        String texto = "";
        HeaderFooterHelper helper = new HeaderFooterHelper();
        texto = helper.setCenterSection(texto, textoCabecera);
        if (tamanioTexto != null) {
            texto = HeaderFooter.fontSize(tamanioTexto)
                    + texto;
        }
        if (negrita) {
            texto = HeaderFooter.startBold()
                    + texto
                    + HeaderFooter.endBold();
        }

        sheet.getHeader().setCenter(texto);
        Footer footer = sheet.getFooter();

        footer.setCenter("Página "
                + HeaderFooter.page()
                + " de "
                + HeaderFooter.numPages());
        footer.setRight(HeaderFooter.date()
                + " "
                + HeaderFooter.time());
    }

    private XSSFCellStyle obtenerEstiloCelda(CeldaTo celda) {
        EstiloExcelUtil estilo = EstiloExcelUtil.getInstancia(workbook);
        boolean entro = false;
        if (celda.isNegrita()) {
            estilo.aplicarNegrita();
            entro = true;
        }
        if (celda.getAlinearHorizontalmente() != null) {
            estilo.alinearHorizontalmente(celda.getAlinearHorizontalmente());
            entro = true;
        }
        if (celda.getAlinearVerticalmente() != null) {
            estilo.alinearVerticalmente(celda.getAlinearVerticalmente());
            entro = true;
        }
        if (celda.getFormato() != null
                && !celda.getFormato().isEmpty()) {
            estilo.aplicarFormato(celda.getFormato());
            entro = true;
        }
        if (celda.getColorTextoHexadecimal() != null
                && !celda.getColorTextoHexadecimal().isEmpty()) {
            estilo.aplicarColorTexto(celda.getColorTextoHexadecimal());
            entro = true;
        }
        if (celda.getTamanioLetra() > 0) {
            estilo.aplicarTamanioTexto(celda.getTamanioLetra());
            entro = true;
        }
        if (celda.isWrapTexto()) {
            estilo.aplicarWrapTexto();
            entro = true;
        }
        if (celda.getColorRellenoHexadecimal() != null
                && !celda.getColorRellenoHexadecimal().isEmpty()) {
            estilo.aplicarColorFondo(celda.getColorRellenoHexadecimal());
            entro = true;
        }
        if (celda.getPatronRelleno() != null) {
            estilo.aplicarPatroColorFondo(celda.getPatronRelleno());
            entro = true;
        }
        if (celda.getColorBordesHexadecimal() != null
                && !celda.getColorBordesHexadecimal().isEmpty()) {
            BorderStyle borderStyle = BorderStyle.NONE;
            if (celda.getEstiloBorde() != null) {
                borderStyle = celda.getEstiloBorde();
            }
            estilo.aplicarBordes(celda.getColorBordesHexadecimal(), borderStyle);
            entro = true;
        }
        if (entro) {
            return estilo.obtenerEstilo();
        }
        return null;
    }

    private void validar() {
        VERSION.getMaxColumns();
        VERSION.getMaxRows();
        int numeroFilas = 0;
        int numeroColumnas = 0;

        for (HojaExcelTo hoja : libroExcelTo.getHojas()) {
            numeroFilas += hoja.getContenido().size();

            if (hoja.getContenido() != null
                    && hoja.getContenido() != null
                    && !hoja.getContenido().isEmpty()) {
                List<CeldaTo> celdas = hoja.getContenido().get(0).getCeldas();
                if (celdas != null
                        && !celdas.isEmpty()) {
                    numeroColumnas = celdas.size();
                }

                numeroFilas += hoja.getContenido().size();
            }

        }
        if (numeroFilas > VERSION.getMaxRows()) {
            throw new ExceptionInInitializerError("Numero de filas supera el máximo permitido: "
                    + VERSION.getMaxRows()
                    + ". Filas Actuales: "
                    + numeroFilas);
        }
        if (numeroColumnas > VERSION.getMaxColumns()) {
            throw new ExceptionInInitializerError("Numero de columnas supera el máximo permitido: "
                    + VERSION.getMaxColumns());
        }

    }
}
