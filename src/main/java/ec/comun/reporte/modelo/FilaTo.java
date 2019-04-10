package ec.comun.reporte.modelo;

import java.util.ArrayList;
import java.util.List;

public class FilaTo {

    private List<CeldaTo> celdas;
    private int numeroFilasUnir = 0;

    public static final FilaTo crearFilaVacia(int numeroColumnas) {
        List<CeldaTo> celdas = new ArrayList<>();
        for (int i = 0; i < numeroColumnas; i++) {
            CeldaTo to = new CeldaTo(i, "");
            celdas.add(to);
        }
        return new FilaTo(celdas);
    }

    public FilaTo() {
        super();
    }

    public FilaTo(List<CeldaTo> celdas) {

        this.celdas = celdas;
    }

    public List<CeldaTo> getCeldas() {
        return celdas;
    }

    public void setCeldas(List<CeldaTo> celdas) {
        this.celdas = celdas;
    }

    public int getNumeroFilasUnir() {
        return numeroFilasUnir;
    }

    public void setNumeroFilasUnir(int numeroFilasUnir) {
        this.numeroFilasUnir = numeroFilasUnir;
    }

}
