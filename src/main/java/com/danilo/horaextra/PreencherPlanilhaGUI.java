package com.danilo.horaextra;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class PreencherPlanilhaGUI extends JFrame {
    private JTextField caminhoArquivoField;
    private JTextField nomeField;
    private JTextField mesField;
    private JTextField entradaNormalField;
    private JTextField saidaNormalField;
    private JTextField salarioField;
    private JTextField dataField;
    private JTextField entradaField;
    private JTextField saidaField;
    private JTextField observacaoField;
    private JButton gerarButton;

    public PreencherPlanilhaGUI() {
        setTitle("Preencher Planilha");
        setSize(400, 500);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLocationRelativeTo(null);

        JPanel panel = new JPanel();
        panel.setLayout(new GridLayout(11, 2));

        panel.add(new JLabel("Caminho do arquivo:"));
        caminhoArquivoField = new JTextField();
        panel.add(caminhoArquivoField);

        panel.add(new JLabel("Nome:"));
        nomeField = new JTextField();
        panel.add(nomeField);

        panel.add(new JLabel("Mês:"));
        mesField = new JTextField();
        panel.add(mesField);

        panel.add(new JLabel("Horário de entrada normal (HH:mm):"));
        entradaNormalField = new JTextField();
        panel.add(entradaNormalField);

        panel.add(new JLabel("Horário de saída normal (HH:mm):"));
        saidaNormalField = new JTextField();
        panel.add(saidaNormalField);

        panel.add(new JLabel("Salário:"));
        salarioField = new JTextField();
        panel.add(salarioField);

        panel.add(new JLabel("Data e dia da semana (01 domingo):"));
        dataField = new JTextField();
        panel.add(dataField);

        panel.add(new JLabel("Horário de entrada (HH:mm):"));
        entradaField = new JTextField();
        panel.add(entradaField);

        panel.add(new JLabel("Horário de saída (HH:mm):"));
        saidaField = new JTextField();
        panel.add(saidaField);

        panel.add(new JLabel("Observação:"));
        observacaoField = new JTextField();
        panel.add(observacaoField);

        gerarButton = new JButton("Gerar Planilha");
        panel.add(gerarButton);

        add(panel);

        gerarButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    preencherPlanilha();
                } catch (IOException ex) {
                    ex.printStackTrace();
                }
            }
        });
    }

    private void preencherPlanilha() throws IOException {
        String caminhoArquivo = caminhoArquivoField.getText();
        String nome = nomeField.getText();
        String mes = mesField.getText();
        String entradaNormal = entradaNormalField.getText();
        String saidaNormal = saidaNormalField.getText();
        double salario = Double.parseDouble(salarioField.getText());

        String data = dataField.getText();
        String[] partesData = data.split(" ");
        int dia = Integer.parseInt(partesData[0]);
        String diaSemana = partesData[1];

        String entrada = entradaField.getText();
        String saida = saidaField.getText();
        String observacao = observacaoField.getText();

        FileInputStream fileInputStream = new FileInputStream(caminhoArquivo);
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

        // Preencher as células com os dados fornecidos
        sheet.getRow(0).getCell(3).setCellValue(nome); // Nome: D1
        sheet.getRow(0).getCell(3).setCellValue(mes); // Mês: D1
        sheet.getRow(1).getCell(17).setCellValue(entradaNormal); // Horário de entrada normal: R2
        sheet.getRow(1).getCell(18).setCellValue(saidaNormal); // Horário de saída normal: S2
        sheet.getRow(2).getCell(4).setCellValue(salario); // Salário: E3

        // Calcular a linha correta com base no dia
        int linha = 5 + (dia - 1) * 2;

        // Preencher as células da linha calculada
        sheet.getRow(linha).getCell(2).setCellValue(dia + " " + diaSemana); // Dia e dia da semana: C(linha)
        sheet.getRow(linha).getCell(3).setCellValue(entrada); // Horário entrada: D(linha)
        sheet.getRow(linha).getCell(4).setCellValue(saida); // Horário saída: E(linha)
        sheet.getRow(linha).getCell(21).setCellValue(observacao); // Observação: V(linha)

        // Salvar o arquivo atualizado
        fileInputStream.close();
        try (FileOutputStream fileOut = new FileOutputStream(caminhoArquivo)) {
            workbook.write(fileOut);
        }
        workbook.close();

        abrirArquivo(caminhoArquivo);
    }

    private void abrirArquivo(String caminhoArquivo) {
        try {
            File arquivo = new File(caminhoArquivo);
            if (arquivo.exists()) {
                if (Desktop.isDesktopSupported()) {
                    Desktop.getDesktop().open(arquivo);
                } else {
                    JOptionPane.showMessageDialog(this, "Abertura de arquivos não suportada no sistema.");
                }
            } else {
                JOptionPane.showMessageDialog(this, "Arquivo não encontrado.");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            PreencherPlanilhaGUI frame = new PreencherPlanilhaGUI();
            frame.setVisible(true);
        });
    }
}
