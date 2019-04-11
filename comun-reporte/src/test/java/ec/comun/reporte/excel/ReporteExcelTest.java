package ec.comun.reporte.excel;

import java.io.File;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.junit.Test;

import ec.comun.reporte.modelo.CeldaTo;
import ec.comun.reporte.modelo.FilaTo;
import ec.comun.reporte.modelo.HojaExcelTo;
import ec.comun.reporte.modelo.LibroExcelTo;
import junit.framework.TestCase;

public class ReporteExcelTest extends TestCase {
  
    private static int index = 0;
    private static String[] columns = { "Name", "Email", "Date Of Birth", "Salary" };
    private static List<Employee> employees = new ArrayList<>();
    static {
        Calendar dateOfBirth = Calendar.getInstance();
        dateOfBirth.set(1992, 7, 21);
        employees.add(new Employee("Rajeev Singh", "rajeev@example.com", dateOfBirth.getTime(), 1200000.0));

        dateOfBirth.set(1965, 10, 15);
        employees.add(new Employee("Thomas cook", "thomas@example.com", dateOfBirth.getTime(), 1500000.25));

        dateOfBirth.set(1987, 4, 18);
        employees.add(new Employee("Steve Maiden", "steve@example.com", dateOfBirth.getTime(), 1800000.100));

    }

   
    @Test
    public final void testGetInstancia() throws Exception {

        File file = ReporteExcel.getInstancia(generarDatosReporte()).construir().exportarArchivo();
        assertNotNull(file);

    }

    private LibroExcelTo generarDatosReporte() {
        List<FilaTo> contenido = new ArrayList<>();
        contenido.addAll(generarDatosGenerales());
        contenido.addAll(generarContenido());
        contenido.addAll(generarResumen());
        HojaExcelTo excel = new HojaExcelTo("Empleados", contenido);
        excel.setTextoCabecera("Reporte de Empleados");
        excel.setNegritaCabecera(true);
        excel.setNegritaCabecera(true);
        excel.setTamanioLetraCabecera((short) 16);
        excel.setTextoCabecera("Reporte de prueba");
        HojaExcelTo excel2 = new HojaExcelTo("Empleados2", contenido);

        return new LibroExcelTo("Reporte de Prueba", Arrays.asList(excel, excel2));

    }

    private List<FilaTo> generarDatosGenerales() {

        List<FilaTo> filas = new ArrayList<>();

        CeldaTo fechaCorteLabel = new CeldaTo(0, "Fecha Corte:");
        fechaCorteLabel.setNegrita(true);
        CeldaTo fechaCorteTexto = new CeldaTo(1, new Date(), "dd-MM-yyyy");
        fechaCorteTexto.setNumeroColumnasUnir(2);
        fechaCorteTexto.setAlinearHorizontalmente(HorizontalAlignment.LEFT);
        CeldaTo fechaPeriodoDesdeLabel = new CeldaTo(0, "Periodo Desde:");
        fechaPeriodoDesdeLabel.setNegrita(true);
        CeldaTo fechaPeriodoDesdeTexto = new CeldaTo(1, new Date(), "dd-MM-yyyy");
        fechaPeriodoDesdeTexto.setAlinearHorizontalmente(HorizontalAlignment.LEFT);
        CeldaTo fechaPeriodoHastaLabel = new CeldaTo(2, "Periodo Desde:");
        fechaPeriodoHastaLabel.setNegrita(true);
        CeldaTo fechaPeriodoHastaTexto = new CeldaTo(3, new Date(), "dd-MM-yyyy");
        fechaPeriodoHastaTexto.setAlinearHorizontalmente(HorizontalAlignment.LEFT);
        FilaTo fechaCorte = new FilaTo(index++, Arrays.asList(fechaCorteLabel, fechaCorteTexto));
        FilaTo periodo = new FilaTo(index++, Arrays.asList(fechaPeriodoDesdeLabel, fechaPeriodoDesdeTexto, fechaPeriodoHastaLabel, fechaPeriodoHastaTexto));
        FilaTo filaVacia = FilaTo.crearFilaVacia(1);
        filaVacia.getCeldas().get(0).setNumeroColumnasUnir(3);
        filaVacia.setIndex(index++);
        filas.addAll(Arrays.asList(fechaCorte, periodo, filaVacia));
        return filas;
    }

    private List<FilaTo> generarContenido() {
        List<FilaTo> filas = new ArrayList<>();
        List<CeldaTo> cabecera = new ArrayList<>();
        for (int i = 0; i < columns.length; i++) {
            CeldaTo to = new CeldaTo(i, columns[i]);
            to.aplicarColorPetro();
            to.setAlinearVerticalmente(VerticalAlignment.CENTER);
            to.setAlinearHorizontalmente(HorizontalAlignment.CENTER);
            to.setNegrita(true);
            to.setTamanioLetra((short) 14);
            cabecera.add(to);
        }
        FilaTo cabeceraFila = new FilaTo(index++, cabecera);
        filas.add(cabeceraFila);
        for (Employee employee : employees) {
            CeldaTo celdaNombre = new CeldaTo(0, employee.getName());
            celdaNombre.centrarDatoCelda();

            CeldaTo celdaEmail = new CeldaTo(1, employee.getEmail());
            celdaEmail.centrarDatoCelda();
            CeldaTo celdaFechaNacimiento = new CeldaTo(2, employee.getDateOfBirth(), "dd-MM-yyyy");
            celdaFechaNacimiento.centrarDatoCelda();
            CeldaTo celdaSalario = new CeldaTo(3, employee.getSalary(), "$#,##0.00");
            FilaTo fila = new FilaTo(index++, Arrays.asList(celdaNombre, celdaEmail, celdaFechaNacimiento, celdaSalario));

            filas.add(fila);
        }
        int tamanio = SpreadsheetVersion.EXCEL97.getMaxRows();
        FilaTo fila = filas.get(filas.size()
                - 1);
        for (int i = 0; i < tamanio; i++) {
            filas.add(fila);
        }
        return filas;
    }

    private static List<FilaTo> generarResumen() {
        List<FilaTo> resumen = new ArrayList<>();
        BigDecimal suma = BigDecimal.ZERO;

        for (Employee e : employees) {
            suma = suma.add(BigDecimal.valueOf(e.getSalary()));
        }
        CeldaTo sumaSalario = new CeldaTo(0, "Suma Total:");
        sumaSalario.setAlinearHorizontalmente(HorizontalAlignment.RIGHT);
        sumaSalario.setNegrita(true);
        sumaSalario.setAlinearVerticalmente(VerticalAlignment.CENTER);
        sumaSalario.setNumeroColumnasUnir(2);
        sumaSalario.setNumeroFilasUnir(1);
        CeldaTo celdaValorSuma = new CeldaTo(3, suma.setScale(BigDecimal.ROUND_HALF_DOWN, 2).doubleValue(), "$#,##0.00");
        celdaValorSuma.setAlinearHorizontalmente(HorizontalAlignment.RIGHT);
        celdaValorSuma.setNegrita(true);
        celdaValorSuma.setNumeroFilasUnir(1);
        celdaValorSuma.setAlinearVerticalmente(VerticalAlignment.CENTER);
        celdaValorSuma.setColorTextoHexadecimal("#FF0000");
        celdaValorSuma.setColorRellenoHexadecimal("#dba129");
        celdaValorSuma.setPatronRelleno(FillPatternType.SOLID_FOREGROUND);
        celdaValorSuma.setColorBordesHexadecimal("#A9A9A9");
        celdaValorSuma.setEstiloBorde(BorderStyle.DOUBLE);
        FilaTo fila = new FilaTo(index++, Arrays.asList(sumaSalario, celdaValorSuma));
        resumen.add(fila);
        return resumen;
    }

}
