package ec.comun.reporte.excel;

import java.io.File;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import ec.comun.reporte.excel.ReporteExcel;
import ec.comun.reporte.modelo.CeldaTo;
import ec.comun.reporte.modelo.FilaTo;
import ec.comun.reporte.modelo.HojaExcelTo;
import ec.comun.reporte.modelo.LibroExcelTo;
import junit.framework.TestCase;

public class ReporteExcelTest extends TestCase {
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

    private static final List<FilaTo> armarResumenContenido(List<Employee> contenido) {
        List<FilaTo> resumen = new ArrayList<>();
        BigDecimal suma = BigDecimal.ZERO;
        for (Employee e : contenido) {
            suma = suma.add(BigDecimal.valueOf(e.getSalary()));
        }
        CeldaTo sumaSalario = new CeldaTo(0, "Suma Total:");
        sumaSalario.setAlinearHorizontalmente(HorizontalAlignment.RIGHT);
        sumaSalario.setNegrita(true);
        sumaSalario.setAlinearVerticalmente(VerticalAlignment.CENTER);
        sumaSalario.setNumeroColumnasUnir(2);
        sumaSalario.setNumeroFilasUnir(1);
        CeldaTo celdaValorSuma = new CeldaTo(3, suma.setScale(BigDecimal.ROUND_HALF_DOWN, 2).doubleValue(),
                "$#,##0.00");
        celdaValorSuma.setAlinearHorizontalmente(HorizontalAlignment.RIGHT);
        celdaValorSuma.setNegrita(true);
        celdaValorSuma.setNumeroFilasUnir(1);
        celdaValorSuma.setAlinearVerticalmente(VerticalAlignment.CENTER);
        celdaValorSuma.setColorTextoHexadecimal("#FF0000");
        celdaValorSuma.setColorRellenoHexadecimal("#dba129");
        celdaValorSuma.setPatronRelleno(FillPatternType.SOLID_FOREGROUND);
        celdaValorSuma.setColorBordesHexadecimal("#A9A9A9");
        celdaValorSuma.setEstiloBorde(BorderStyle.DOUBLE);
        FilaTo fila = new FilaTo(Arrays.asList(sumaSalario, celdaValorSuma));
        FilaTo filaVacia = FilaTo.crearFilaVacia(3);
        filaVacia.setNumeroFilasUnir(3);
        resumen.add(filaVacia);

        resumen.add(fila);
        resumen.add(filaVacia);
        return resumen;
    }

    protected void setUp() throws Exception {
        super.setUp();
    }

    public final void testGetInstancia() throws Exception {
        File file = ReporteExcel.getInstancia(generarDatos()).construir().exportarArchivo();
        assertNotNull(file);

    }

    private LibroExcelTo generarDatos() {
        List<FilaTo> filas = new ArrayList<>();
        for (Employee employee : employees) {
            CeldaTo celdaNombre = new CeldaTo(0, employee.getName());
            CeldaTo celdaEmail = new CeldaTo(1, employee.getEmail());
            CeldaTo celdaFechaNacimiento = new CeldaTo(2, employee.getDateOfBirth(), "dd-MM-yyyy");
            CeldaTo celdaSalario = new CeldaTo(3, employee.getSalary(), "$#,##0.00");
            FilaTo fila = new FilaTo(Arrays.asList(celdaNombre, celdaEmail, celdaFechaNacimiento, celdaSalario));

            filas.add(fila);
        }

        HojaExcelTo excel = new HojaExcelTo("Empleados", filas, armarResumenContenido(employees), generarCabecera());
        excel.setNegritaCabecera(true);
        excel.setTamanioLetraCabecera((short) 16);
        excel.setTextoCabecera("Reporte de prueba");
        HojaExcelTo excel2 = new HojaExcelTo("Empleados2", filas, columns);

        return new LibroExcelTo("poi-generated-file", Arrays.asList(excel, excel2));

    }

    private List<CeldaTo> generarCabecera() {
        List<CeldaTo> celdas = new ArrayList<>();

        for (int i = 0; i < columns.length; i++) {
            CeldaTo to = new CeldaTo(i, columns[i]);
            to.setAlinearHorizontalmente(HorizontalAlignment.CENTER);
            to.setAlinearVerticalmente(VerticalAlignment.CENTER);
            to.setTamanioLetra((short) 14);
            to.aplicarColorPetro();
            to.setNegrita(true);
            celdas.add(to);
        }
        return celdas;
    }
}
