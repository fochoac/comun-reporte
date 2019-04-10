package ec.comun.reporte.modelo;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class CabeceraTo {

    private List<FilaTo> filas;

    public CabeceraTo(List<FilaTo> filas) {

        this.setFilas(filas);
    }

    public CabeceraTo(String... datosCabecera) {
        super();
        List<CeldaTo> celdasCabecera = new ArrayList<>();
        for (int i = 0; i < datosCabecera.length; i++) {
            celdasCabecera.add(generarCeldas(i, datosCabecera[i]));
        }
        FilaTo fila = new FilaTo(celdasCabecera);
        this.setFilas(Arrays.asList(fila));
    }

    private CeldaTo generarCeldas(int index, final String dato) {
        return new CeldaTo(index, dato);
    }

    public CabeceraTo(CeldaTo... celdas) {
        FilaTo fila = new FilaTo(Arrays.asList(celdas));
        this.setFilas(Arrays.asList(fila));
    }
    
    public List<FilaTo> getFilas() {
        return filas;
    }

    public void setFilas(List<FilaTo> filas) {
        this.filas = filas;
    }

}
