package ec.comun.reporte.modelo;

import java.util.List;

public class LibroExcelTo {
    
    private List<HojaExcelTo> hojas;
    private String nombreArchivo;

    public LibroExcelTo() {
        super();
        
    }

    public LibroExcelTo(String nombreArchivo, List<HojaExcelTo> hojas) {

        this.hojas = hojas;
        this.nombreArchivo = nombreArchivo;
    }

    public List<HojaExcelTo> getHojas() {
        return hojas;
    }

    public void setHojas(List<HojaExcelTo> hojas) {
        this.hojas = hojas;
    }

    public String getNombreArchivo() {
        return nombreArchivo;
    }

    public void setNombreArchivo(String nombreArchivo) {
        this.nombreArchivo = nombreArchivo;
    }

}
