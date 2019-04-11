package ec.comun.reporte.modelo;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import ec.comun.reporte.util.EstiloExcelUtil;

public class CeldaTo implements Comparable<CeldaTo> {
    private Integer indice;
    private Object dato;
    private boolean textoEnriquecido = false;
    private boolean negrita;
    private VerticalAlignment alinearVerticalmente = VerticalAlignment.BOTTOM;
    private HorizontalAlignment alinearHorizontalmente = HorizontalAlignment.GENERAL;
    private short tamanioLetra;
    private String colorTextoHexadecimal;
    private String formato;
    private boolean wrapTexto = false;
    private int numeroColumnasUnir = 0;
    private int numeroFilasUnir = 0;
    private String colorRellenoHexadecimal;
    private FillPatternType patronRelleno = FillPatternType.NO_FILL;
    private String colorBordesHexadecimal;
    private BorderStyle estiloBorde = BorderStyle.NONE;

    public CeldaTo() {
        super();
    }

    public CeldaTo(int index, Object dato) {
        this.indice = index;
        this.dato = dato;

    }

    public CeldaTo(int index, Object dato, String formato) {
        this(index, dato);
        this.formato = formato;

    }

    public String getFormato() {
        return formato;
    }

    public void setFormato(String formato) {
        this.formato = formato;
    }

    public Object getDato() {
        return dato;
    }

    public void setDato(Object dato) {
        this.dato = dato;
    }

    public Integer getIndice() {
        return indice;
    }

    public void setIndice(int indice) {
        this.indice = indice;
    }

    public int getNumeroColumnasUnir() {
        return numeroColumnasUnir;
    }

    public void setNumeroColumnasUnir(int numeroColumnasUnir) {
        this.numeroColumnasUnir = numeroColumnasUnir;
    }

    public int getNumeroFilasUnir() {
        return numeroFilasUnir;
    }

    public void setNumeroFilasUnir(int numeroFilasUnir) {
        this.numeroFilasUnir = numeroFilasUnir;
    }

    @Override
    public int compareTo(CeldaTo o) {
        return this.getIndice().compareTo(o.getIndice());
    }

    public boolean isNegrita() {
        return negrita;
    }

    public void setNegrita(boolean negrita) {
        this.negrita = negrita;
    }

    public VerticalAlignment getAlinearVerticalmente() {
        return alinearVerticalmente;
    }

    public void setAlinearVerticalmente(VerticalAlignment alinearVerticalmente) {
        this.alinearVerticalmente = alinearVerticalmente;
    }

    public HorizontalAlignment getAlinearHorizontalmente() {
        return alinearHorizontalmente;
    }

    public void setAlinearHorizontalmente(HorizontalAlignment alinearHorizontalmente) {
        this.alinearHorizontalmente = alinearHorizontalmente;
    }

    public short getTamanioLetra() {
        return tamanioLetra;
    }

    public void setTamanioLetra(short tamanioLetra) {
        this.tamanioLetra = tamanioLetra;
    }

    public void setIndice(Integer indice) {
        this.indice = indice;
    }

    public String getColorTextoHexadecimal() {
        return colorTextoHexadecimal;
    }

    public void setColorTextoHexadecimal(String colorTextoHexadecimal) {

        this.colorTextoHexadecimal = colorTextoHexadecimal.replace("#", "");
    }

    public boolean isTextoEnriquecido() {
        return textoEnriquecido;
    }

    public void setTextoEnriquecido(boolean textoEnriquecido) {
        this.textoEnriquecido = textoEnriquecido;
    }

    public void aplicarColorPetro() {
        setColorTextoHexadecimal(EstiloExcelUtil.ARGB_PETRO);
    }

    public boolean isWrapTexto() {
        return wrapTexto;
    }

    public void setWrapTexto(boolean wrapTexto) {
        this.wrapTexto = wrapTexto;
    }

    public String getColorRellenoHexadecimal() {
        return colorRellenoHexadecimal;
    }

    public void setColorRellenoHexadecimal(String colorRellenoHexadecimal) {
        this.colorRellenoHexadecimal = colorRellenoHexadecimal.replace("#", "");
        ;
    }

    public FillPatternType getPatronRelleno() {
        return patronRelleno;
    }

    public void setPatronRelleno(FillPatternType patronRelleno) {
        this.patronRelleno = patronRelleno;
    }

    public void centrarDatoCelda() {
        this.alinearHorizontalmente = HorizontalAlignment.CENTER;
        this.alinearVerticalmente = VerticalAlignment.CENTER;
    }

    public String getColorBordesHexadecimal() {
        return colorBordesHexadecimal;
    }

    public void setColorBordesHexadecimal(String colorBordesHexadecimal) {
        this.colorBordesHexadecimal = colorBordesHexadecimal.replace("#", "");
        ;
    }

    public BorderStyle getEstiloBorde() {
        return estiloBorde;
    }

    public void setEstiloBorde(BorderStyle estiloBorde) {
        this.estiloBorde = estiloBorde;
    }

    public String getEstilo() {
        final String estilo = "N:%s, VA:%d, HA:%d, TL:%d, CT:%s, F:%s, WT:%s, CRC:%s, PRC:%d, CB:%s, BS: %d";
        return String.format(estilo, negrita ? "S" : "N", alinearVerticalmente.getCode(), alinearHorizontalmente.getCode(), tamanioLetra, colorTextoHexadecimal, formato,
                wrapTexto ? "S" : "N", colorRellenoHexadecimal, patronRelleno.getCode(), colorBordesHexadecimal, estiloBorde.getCode());
    }
}
