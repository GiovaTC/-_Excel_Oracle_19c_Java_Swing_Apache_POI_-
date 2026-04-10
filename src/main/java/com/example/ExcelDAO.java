package com.example;

import java.sql.Connection;
import java.sql.Driver;
import java.sql.DriverManager;
import java.sql.PreparedStatement;

public class ExcelDAO {

    private final String URL = "jdbc:oracle:thin:@//localhost:1521/orcl";
    private final String USER = "system";
    private final String PASS = "Tapiero123";

    public void guardarDato(String nombre, double valor) {
        String sql = "INSERT INTO EXCEL_DATOS (NOMBRE, VALOR) VALUES (?, ?)";

        try ( Connection conn = DriverManager.getConnection(URL, USER, PASS);
              PreparedStatement ps = conn.prepareStatement(sql)) {

            ps.setString(1, nombre);
            ps.setDouble(2, valor);
            ps.executeUpdate();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
