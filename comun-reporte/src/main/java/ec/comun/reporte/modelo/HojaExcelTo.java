package ec.comun.reporte.modelo;

import java.util.List;

public class HojaExcelTo {
    private String nombre;

    private List<FilaTo> contenido;

    private boolean mostrarGrilla = true;
    private boolean imprimirGrilla = true;
    private boolean fitToPage = true;
    private String textoCabecera = "EP PETROECUADOR";
    private Short tamanioLetraCabecera = 14;
    private boolean negritaCabecera = true;

    public HojaExcelTo(String nombre, List<FilaTo> contenido) {
        this.nombre = nombre;
        this.setContenido(contenido);
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
