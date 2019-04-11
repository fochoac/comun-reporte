package ec.comun.reporte.modelo;

import java.util.ArrayList;
import java.util.List;

public class FilaTo implements Comparable<FilaTo> {

    private List<CeldaTo> celdas;
    private Integer index;
    private int numeroFilasUnir = 0;
    

    public static final FilaTo crearFilaVacia(int numeroColumnas) {
        List<CeldaTo> celdas = new ArrayList<>();
        for (int i = 0; i < numeroColumnas; i++) {
            CeldaTo to = new CeldaTo(i, "");
            celdas.add(to);
        }
        return new FilaTo(1, celdas);
    }

    public FilaTo() {
        super();
    }

    public FilaTo(int index, List<CeldaTo> celdas) {
        
        this.index = index;
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

    @Override
    public int compareTo(FilaTo o) {

        return getIndex().compareTo(o.getIndex());
    }

    public Integer getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

}
