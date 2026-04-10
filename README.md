# -_Excel_Oracle_19c_Java_Swing_Apache_POI_- :.
📊 Excel → Oracle 19c (Java + Swing + Apache POI):

<img width="1024" height="1024" alt="image" src="https://github.com/user-attachments/assets/0940d57b-814a-488c-822e-be80d0e711c7" />    

<img width="2551" height="1076" alt="image" src="https://github.com/user-attachments/assets/25029ab7-c537-4566-8946-842951b8769e" />    

```
Solución completa y funcional con:

✔ Interfaz gráfica (Swing)
✔ Generación y descarga de Excel (.xlsx)
✔ Archivo de ejemplo automático
✔ Inserción en Oracle 19c vía JDBC
✔ Arquitectura simple (Vista + Servicio + DAO) .

🧩 1. Dependencias (Maven)
<dependencies>
    <!-- Apache POI para Excel -->
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml</artifactId>
        <version>5.2.5</version>
    </dependency>

    <!-- Oracle JDBC -->
    <dependency>
        <groupId>com.oracle.database.jdbc</groupId>
        <artifactId>ojdbc8</artifactId>
        <version>19.3.0.0</version>
    </dependency>
</dependencies>

🧩 2. Modelo BD (Oracle 19c)
CREATE TABLE EXCEL_DATOS (
    ID NUMBER GENERATED ALWAYS AS IDENTITY,
    NOMBRE VARCHAR2(100),
    VALOR NUMBER
);

🧩 3. Clase DAO (Oracle)
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;

public class ExcelDAO {

    private final String URL = "jdbc:oracle:thin:@localhost:1521:XE";
    private final String USER = "system";
    private final String PASS = "oracle";

    public void guardarDato(String nombre, double valor) {
        String sql = "INSERT INTO EXCEL_DATOS (NOMBRE, VALOR) VALUES (?, ?)";

        try (Connection conn = DriverManager.getConnection(URL, USER, PASS);
             PreparedStatement ps = conn.prepareStatement(sql)) {

            ps.setString(1, nombre);
            ps.setDouble(2, valor);
            ps.executeUpdate();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

🧩 4. Servicio Excel (Generar + Leer)
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileInputStream;

public class ExcelService {

    private ExcelDAO dao = new ExcelDAO();

    // Generar Excel de ejemplo
    public File generarExcel() {
        File file = new File("datos.xlsx");

        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Datos");

            String[] nombres = {"A", "B", "C", "D", "E"};
            double[] valores = {10, 20, 30, 40, 50};

            for (int i = 0; i < nombres.length; i++) {
                Row row = sheet.createRow(i);
                row.createCell(0).setCellValue(nombres[i]);
                row.createCell(1).setCellValue(valores[i]);
            }

            FileOutputStream fos = new FileOutputStream(file);
            wb.write(fos);
            fos.close();

        } catch (Exception e) {
            e.printStackTrace();
        }

        return file;
    }

    // Leer Excel e insertar en BD
    public void procesarExcel(File file) {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheetAt(0);

            for (Row row : sheet) {
                String nombre = row.getCell(0).getStringCellValue();
                double valor = row.getCell(1).getNumericCellValue();

                dao.guardarDato(nombre, valor);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

🧩 5. Interfaz Gráfica (Swing)
import javax.swing.*;
import java.awt.event.ActionEvent;
import java.io.File;

public class Vista extends JFrame {

    private JButton btnGenerar = new JButton("Generar Excel");
    private JButton btnCargar = new JButton("Enviar a Oracle");

    private ExcelService service = new ExcelService();
    private File archivo;

    public Vista() {
        setTitle("Excel → Oracle 19c");
        setSize(400, 200);
        setLayout(null);
        setDefaultCloseOperation(EXIT_ON_CLOSE);

        btnGenerar.setBounds(100, 30, 200, 30);
        btnCargar.setBounds(100, 80, 200, 30);

        add(btnGenerar);
        add(btnCargar);

        btnGenerar.addActionListener(this::generarExcel);
        btnCargar.addActionListener(this::cargarExcel);
    }

    private void generarExcel(ActionEvent e) {
        archivo = service.generarExcel();
        JOptionPane.showMessageDialog(this, "Excel generado: " + archivo.getAbsolutePath());
    }

    private void cargarExcel(ActionEvent e) {
        if (archivo != null) {
            service.procesarExcel(archivo);
            JOptionPane.showMessageDialog(this, "Datos enviados a Oracle");
        } else {
            JOptionPane.showMessageDialog(this, "Primero genera el Excel");
        }
    }
}

🧩 6. Main
public class Main {
    public static void main(String[] args) {
        new Vista().setVisible(true);
    }
}
⚙️ Flujo Completo
Click "Generar Excel"
→ crea datos.xlsx automáticamente

Click "Enviar a Oracle"
→ lee el archivo
→ inserta cada fila en Oracle

📌 Resultado Esperado
ID	NOMBRE	VALOR
1	A	10
2	B	20
…	…	…

🚀 Posibles Mejoras
Usar JFileChooser para elegir ruta de descarga
Mostrar datos en JTable
Validaciones (celdas vacías, tipos de datos)
Uso de Stored Procedure (SP) en lugar de INSERT directo
Implementar pool de conexiones (HikariCP) :. . / . 
