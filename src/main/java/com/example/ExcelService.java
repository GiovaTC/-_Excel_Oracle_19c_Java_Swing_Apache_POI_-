package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class ExcelService {

    private ExcelDAO dao = new ExcelDAO();

    // generar excel de ejemplo .
    public File generarExcel() {
        File file = new File("datos.xlsx");

        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Datos");

            String[] nombres = {"A","B","C","D","E"};
            double[] valores = {10, 20, 30, 40, 50};

            for (int i = 0; i < nombres.length; i++){
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

    // leer excel a insertar en bd .
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
