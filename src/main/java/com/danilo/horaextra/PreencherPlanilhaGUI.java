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
    private JComboBox<Integer> diaComboBox;
    private JComboBox<String> diaSemanaComboBox;
    private JTextField entradaField;
    private JTextField saidaField;
    private JTextField observacaoField;
    private JButton gerarButton;
    private JButton adicionarButton;
    private JButton procurarButton;

    private Workbook workbook;
    private Sheet sheet;

    public PreencherPlanilhaGUI() {
        setTitle("Preencher Planilha");
        setSize(400, 500);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLocationRelativeTo(null);

        JPanel panel = new JPanel();
        panel.setLayout(new GridLayout(13, 2));

        panel.add(new JLabel("Caminho do arquivo:"));

        JPanel filePanel = new JPanel(new BorderLayout());
        caminhoArquivoField = new JTextField("C:\\Users\\danilo.silva\\Desktop\\horasextra");
        caminhoArquivoField.setEditable(false);
        filePanel.add(caminhoArquivoField, BorderLayout.CENTER);
        procurarButton = new JButton("Procurar");
        procurarButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                int returnValue = fileChooser.showOpenDialog(null);
                if (returnValue == JFileChooser.APPROVE_OPTION) {
                    File selectedFile = fileChooser.getSelectedFile();
                    caminhoArquivoField.setText(selectedFile.getAbsolutePath());
                    abrirArquivoParaEdicao(selectedFile.getAbsolutePath());
                }
            }
        });
        filePanel.add(procurarButton, BorderLayout.EAST);
        panel.add(filePanel);

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

        panel.add(new JLabel("Dia:"));
        Integer[] dias = new Integer[31];
        for (int i = 1; i <= 31; i++) {
            dias[i - 1] = i;
        }
        diaComboBox = new JComboBox<>(dias);
        panel.add(diaComboBox);

        panel.add(new JLabel("Dia da semana:"));
        String[] diasSemana = {"domingo", "sábado", "feriado"};
        diaSemanaComboBox = new JComboBox<>(diasSemana);
        panel.add(diaSemanaComboBox);

        panel.add(new JLabel("Horário de entrada (HH:mm):"));
        entradaField = new JTextField();
        panel.add(entradaField);

        panel.add(new JLabel("Horário de saída (HH:mm):"));
        saidaField = new JTextField();
        panel.add(saidaField);

        panel.add(new JLabel("Observação:"));
        observacaoField = new JTextField();
        panel.add(observacaoField);

        adicionarButton = new JButton("Adicionar");
        panel.add(adicionarButton);

        gerarButton = new JButton("Gerar Planilha");
        panel.add(gerarButton);

        add(panel);

        adicionarButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    adicionarInformacoes();
                    limparCampos();
                } catch (IOException ex) {
                    ex.printStackTrace();
                }
            }
        });

        gerarButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    if (workbook == null || sheet == null) {
                        abrirArquivoParaEdicao(caminhoArquivoField.getText());
                    }
                    // Adicionar informações atuais se existirem antes de gerar a planilha
                    if (!entradasEstaoVazias()) {
                        adicionarInformacoes();
                    }
                    salvarEFecharArquivo();
                } catch (IOException ex) {
                    ex.printStackTrace();
                }
            }
        });
    }

    private void abrirArquivoParaEdicao(String caminhoArquivo) {
        try {
            FileInputStream fileInputStream = new FileInputStream(caminhoArquivo);
            workbook = new XSSFWorkbook(fileInputStream);
            sheet = workbook.getSheetAt(0);
            fileInputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void adicionarInformacoes() throws IOException {
        if (workbook == null || sheet == null) {
            abrirArquivoParaEdicao(caminhoArquivoField.getText());
        }

        String nome = nomeField.getText();
        String mes = mesField.getText();
        String entradaNormal = entradaNormalField.getText();
        String saidaNormal = saidaNormalField.getText();
        double salario = Double.parseDouble(salarioField.getText());

        int dia = (int) diaComboBox.getSelectedItem();
        String diaSemana = (String) diaSemanaComboBox.getSelectedItem();
        String data = String.format("%02d %s", dia, diaSemana);

        String entrada = entradaField.getText();
        String saida = saidaField.getText();
        String observacao = observacaoField.getText();

        // Preencher as células com os dados fornecidos
        updateCell(sheet, 0, 3, nome); // Nome: D1
        updateCell(sheet, 1, 3, mes);  // Mês: D2
        updateCell(sheet, 1, 17, entradaNormal); // Horário de entrada normal: R2
        updateCell(sheet, 1, 18, saidaNormal);   // Horário de saída normal: S2
        updateCell(sheet, 2, 4, salario);        // Salário: E3

        // Calcular a linha correta com base no dia
        int linha = 5 + (dia - 1) * 2;

        // Preencher as células da linha calculada
        updateCell(sheet, linha, 2, data);       // Dia e dia da semana: C(linha)
        updateCell(sheet, linha, 3, entrada);    // Horário entrada: D(linha)
        updateCell(sheet, linha, 4, saida);      // Horário saída: E(linha)
        updateCell(sheet, linha, 21, observacao); // Observação: V(linha)
    }

    private boolean entradasEstaoVazias() {
        return entradaField.getText().isEmpty() &&
                saidaField.getText().isEmpty() &&
                observacaoField.getText().isEmpty();
    }

    private void salvarEFecharArquivo() throws IOException {
        // Forçar a recalculação de todas as fórmulas na planilha
        workbook.setForceFormulaRecalculation(true);
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        evaluator.evaluateAll();

        // Salvar o arquivo atualizado
        String caminhoArquivo = caminhoArquivoField.getText();
        try (FileOutputStream fileOut = new FileOutputStream(caminhoArquivo)) {
            workbook.write(fileOut);
        }

        abrirArquivo(caminhoArquivo);
    }

    private void limparCampos() {
        diaComboBox.setSelectedIndex(0);
        diaSemanaComboBox.setSelectedIndex(0);
        entradaField.setText("");
        saidaField.setText("");
        observacaoField.setText("");
    }

    private void updateCell(Sheet sheet, int rowIndex, int colIndex, String value) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        cell.setCellValue(value);
    }

    private void updateCell(Sheet sheet, int rowIndex, int colIndex, double value) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        cell.setCellValue(value);
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
