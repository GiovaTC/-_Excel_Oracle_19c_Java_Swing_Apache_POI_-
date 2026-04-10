package com.example;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.io.File;

public class Vista extends JFrame {

    private JButton btnGenerar = new JButton("generar excel");
    private JButton btnCargar = new JButton("enviar a oracle");

    private ExcelService service = new ExcelService();
    private File archivo;

    public Vista() {
        setTitle("EXCEL -> ORACLE 19C");
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
            JOptionPane.showMessageDialog(this, "Datos enviados a oracle");
        } else {
            JOptionPane.showMessageDialog(this, "Primero genera el excel");
        }
    }
}   
