package ec.comun.reporte.modelo;

import java.util.List;

public class HojaExcelTo {
    private String nombre;
    private CabeceraTo cabecera;
    private List<FilaTo> contenido;
    private List<FilaTo> resumenContenido;
    private boolean mostrarGrilla = false;
    private boolean imprimirGrilla = true;
    private boolean fitToPage = true;
    private String textoCabecera = "Cabecera de prueba";
    private Short tamanioLetraCabecera = 14;
    private boolean negritaCabecera = true;

    public HojaExcelTo(String nombre, List<FilaTo> contenido) {
        this.nombre = nombre;
        this.setContenido(contenido);
    }

    public HojaExcelTo(String nombre, List<FilaTo> contenido, List<FilaTo> resumenContenido) {
        this(nombre, contenido);
        this.resumenContenido = resumenContenido;
    }

    public HojaExcelTo(String nombre, List<FilaTo> contenido, List<FilaTo> resumenContenido, String[] cabecera) {
        this(nombre, contenido, resumenContenido);
        CabeceraTo cabeceraTo = new CabeceraTo(cabecera);
        this.cabecera = cabeceraTo;
    }

    public HojaExcelTo(String nombre, List<FilaTo> contenido, List<FilaTo> resumenContenido,
            List<CeldaTo> datosCabecera) {
        this(nombre, contenido, resumenContenido);

        CabeceraTo cabeceraTo = new CabeceraTo(datosCabecera.toArray(new CeldaTo[] {}));
        this.cabecera = cabeceraTo;
    }

    public HojaExcelTo(final String nombre, List<FilaTo> contenido, List<FilaTo> resumenContenido,
            CabeceraTo datosCabecera) {
        this(nombre, contenido, resumenContenido);

        this.cabecera = datosCabecera;
    }

    public HojaExcelTo(String nombre, List<FilaTo> contenido, List<FilaTo> resumenContenido, CeldaTo... datosCabecera) {
        this(nombre, contenido, resumenContenido);

        CabeceraTo cabeceraTo = new CabeceraTo(datosCabecera);
        this.cabecera = cabeceraTo;
    }

    public HojaExcelTo(String nombre, List<FilaTo> contenido, String[] cabecera) {
        this(nombre, contenido);
        CabeceraTo cabeceraTo = new CabeceraTo(cabecera);
        this.cabecera = cabeceraTo;
    }

    public HojaExcelTo(String nombre, List<FilaTo> contenido, CeldaTo... datosCabecera) {
        this(nombre, contenido);

        CabeceraTo cabeceraTo = new CabeceraTo(datosCabecera);
        this.cabecera = cabeceraTo;
    }

    public CabeceraTo getCabecera() {
        return cabecera;
    }

    public void setCabecera(CabeceraTo cabecera) {
        this.cabecera = cabecera;
    }

    public String getNombre() {
        return nombre;
    }

    public void setNombreArchivo(String nombreHoja) {
        this.nombre = nombreHoja;
    }

    public List<FilaTo> getContenido() {
        return contenido;
    }

    public void setContenido(List<FilaTo> contenido) {
        this.contenido = contenido;
    }

    public List<FilaTo> getResumenContenido() {
        return resumenContenido;
    }

    public void setResumenContenido(List<FilaTo> resumenContenido) {
        this.resumenContenido = resumenContenido;
    }

    public boolean isMostrarGrilla() {
        return mostrarGrilla;
    }

    public void setMostrarGrilla(boolean mostrarGrilla) {
        this.mostrarGrilla = mostrarGrilla;
    }

    public boolean isImprimirGrilla() {
        return imprimirGrilla;
    }

    public void setImprimirGrilla(boolean imprimirGrilla) {
        this.imprimirGrilla = imprimirGrilla;
    }

    public boolean isFitToPage() {
        return fitToPage;
    }

    public void setFitToPage(boolean fitToPage) {
        this.fitToPage = fitToPage;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }

    public String getTextoCabecera() {
        return textoCabecera;
    }

    public void setTextoCabecera(String textoCabecera) {
        this.textoCabecera = textoCabecera;
    }

    public Short getTamanioLetraCabecera() {
        return tamanioLetraCabecera;
    }

    public void setTamanioLetraCabecera(Short tamanioLetraCabecera) {
        this.tamanioLetraCabecera = tamanioLetraCabecera;
    }

    public boolean isNegritaCabecera() {
        return negritaCabecera;
    }

    public void setNegritaCabecera(boolean negritaCabecera) {
        this.negritaCabecera = negritaCabecera;
    }
}
