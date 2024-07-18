package com.danilo.horaextra;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.text.NumberFormat;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.Locale;
import java.util.prefs.Preferences;

public class PreencherPlanilhaGUI extends JFrame {
    private JTextField caminhoArquivoField;
    private JTextField nomeField;
    private JComboBox<String> mesComboBox;
    private JFormattedTextField salarioField;
    private JComboBox<Integer> diaComboBox;
    private JComboBox<String> diaSemanaComboBox;
    private JButton entradaNormalButton;
    private JButton saidaNormalButton;
    private JButton entradaButton;
    private JButton saidaButton;
    private JTextField observacaoField;
    private JButton gerarButton;
    private JButton adicionarButton;
    private JButton procurarButton;
    private JButton resetarButton;
    private JButton imprimirButton;

    private LocalTime entradaNormal;
    private LocalTime saidaNormal;
    private LocalTime entrada;
    private LocalTime saida;

    private Workbook workbook;
    private Sheet sheet;

    private Preferences prefs;

    public PreencherPlanilhaGUI() {
        prefs = Preferences.userNodeForPackage(PreencherPlanilhaGUI.class);

        setTitle("Preencher Planilha");
        setSize(600, 600);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLocationRelativeTo(null);

        JPanel panel = new JPanel();
        panel.setLayout(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(5, 5, 5, 5);

        // Font and colors
        Font buttonFont = new Font("Arial", Font.PLAIN, 14);
        Color buttonBackgroundColor = Color.LIGHT_GRAY;
        Color buttonBackgroundColor1 = Color.GRAY;
        Color buttonForegroundColor = Color.BLACK;

        // Labels
        JLabel caminhoLabel = new JLabel("Arquivo");
        caminhoLabel.setFont(new Font("Arial", Font.BOLD, 15));
        addComponent(panel, caminhoLabel, gbc, 0, 0);

        caminhoArquivoField = new JTextField(prefs.get("caminhoArquivo", ""));
        caminhoArquivoField.setEditable(false);
        caminhoArquivoField.setPreferredSize(new Dimension(200, 25));
        addComponent(panel, caminhoArquivoField, gbc, 1, 0, 2, 1);

        procurarButton = new JButton("Procurar");
        procurarButton.setFont(buttonFont);
        procurarButton.setBackground(buttonBackgroundColor);
        procurarButton.setForeground(buttonForegroundColor);
        procurarButton.setPreferredSize(new Dimension(100, 30));
        procurarButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                int returnValue = fileChooser.showOpenDialog(null);
                if (returnValue == JFileChooser.APPROVE_OPTION) {
                    File selectedFile = fileChooser.getSelectedFile();
                    caminhoArquivoField.setText(selectedFile.getAbsolutePath());
                    prefs.put("caminhoArquivo", selectedFile.getAbsolutePath());
                    abrirArquivoParaEdicao(selectedFile.getAbsolutePath());
                }
            }
        });
        addComponent(panel, procurarButton, gbc, 3, 0);

        JLabel nomeLabel = new JLabel("Nome");
        nomeLabel.setFont(new Font("Arial", Font.BOLD, 15));
        addComponent(panel, nomeLabel, gbc, 0, 1);

        nomeField = new JTextField(prefs.get("nome", ""));
        nomeField.setPreferredSize(new Dimension(200, 25));
        addComponent(panel, nomeField, gbc, 1, 1, 3, 1);

        JLabel mesLabel = new JLabel("Mês");
        mesLabel.setFont(new Font("Arial", Font.BOLD, 15));
        addComponent(panel, mesLabel, gbc, 0, 2);

        String[] meses = {"janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"};
        mesComboBox = new JComboBox<>(meses);
        mesComboBox.setPreferredSize(new Dimension(200, 25));
        addComponent(panel, mesComboBox, gbc, 1, 2, 3, 1);

        JLabel entradaNormalLabel = new JLabel("Horário de trabalho");
        entradaNormalLabel.setFont(new Font("Arial", Font.BOLD, 14));
        addComponent(panel, entradaNormalLabel, gbc, 0, 3);

        entradaNormalButton = new JButton("Entrada");
        entradaNormalButton.setFont(buttonFont);
        entradaNormalButton.setBackground(buttonBackgroundColor1);
        entradaNormalButton.setForeground(buttonForegroundColor);
        entradaNormalButton.setPreferredSize(new Dimension(200, 30));
        entradaNormalButton.addActionListener(e -> selecionarHora(entradaNormalButton, true));
        addComponent(panel, entradaNormalButton, gbc, 1, 3, 3, 1);

        JLabel saidaNormalLabel = new JLabel("Horário de trabalho");
        saidaNormalLabel.setFont(new Font("Arial", Font.BOLD, 15));
        addComponent(panel, saidaNormalLabel, gbc, 0, 4);

        saidaNormalButton = new JButton("Saída");
        saidaNormalButton.setFont(buttonFont);
        saidaNormalButton.setBackground(buttonBackgroundColor1);
        saidaNormalButton.setForeground(buttonForegroundColor);
        saidaNormalButton.setPreferredSize(new Dimension(200, 30));
        saidaNormalButton.addActionListener(e -> selecionarHora(saidaNormalButton, false));
        addComponent(panel, saidaNormalButton, gbc, 1, 4, 3, 1);

        JLabel salarioLabel = new JLabel("Salário (R$)");
        salarioLabel.setFont(new Font("Arial", Font.BOLD, 15));
        addComponent(panel, salarioLabel, gbc, 0, 5);

        NumberFormat currencyFormat = NumberFormat.getCurrencyInstance(new Locale("pt", "BR"));
        salarioField = new JFormattedTextField(currencyFormat);
        salarioField.setText(prefs.get("salario", ""));
        salarioField.setPreferredSize(new Dimension(200, 25));
        addComponent(panel, salarioField, gbc, 1, 5, 3, 1);

        JLabel diaLabel = new JLabel("Dia");
        diaLabel.setFont(new Font("Arial", Font.BOLD, 15));
        addComponent(panel, diaLabel, gbc, 0, 6);

        Integer[] dias = new Integer[31];
        for (int i = 1; i <= 31; i++) {
            dias[i - 1] = i;
        }
        diaComboBox = new JComboBox<>(dias);
        diaComboBox.setPreferredSize(new Dimension(200, 25));
        addComponent(panel, diaComboBox, gbc, 1, 6);

        JLabel diaSemanaLabel = new JLabel("Dia da semana");
        diaSemanaLabel.setFont(new Font("Arial", Font.BOLD, 15));
        addComponent(panel, diaSemanaLabel, gbc, 0, 7);

        String[] diasSemana = {"domingo", "sábado", "feriado"};
        diaSemanaComboBox = new JComboBox<>(diasSemana);
        diaSemanaComboBox.setPreferredSize(new Dimension(200, 25));
        addComponent(panel, diaSemanaComboBox, gbc, 1, 7, 3, 1);

        JLabel entradaLabel = new JLabel("Hora Extra");
        entradaLabel.setFont(new Font("Arial", Font.BOLD, 15));
        addComponent(panel, entradaLabel, gbc, 0, 8);

        entradaButton = new JButton("Entrada");
        entradaButton.setFont(buttonFont);
        entradaButton.setBackground(buttonBackgroundColor1);
        entradaButton.setForeground(buttonForegroundColor);
        entradaButton.setPreferredSize(new Dimension(200, 30));
        entradaButton.addActionListener(e -> selecionarHora(entradaButton, true));
        addComponent(panel, entradaButton, gbc, 1, 8, 3, 1);

        JLabel saidaLabel = new JLabel("Hora Extra");
        saidaLabel.setFont(new Font("Arial", Font.BOLD, 15));
        addComponent(panel, saidaLabel, gbc, 0, 9);

        saidaButton = new JButton("Saída");
        saidaButton.setFont(buttonFont);
        saidaButton.setBackground(buttonBackgroundColor1);
        saidaButton.setForeground(buttonForegroundColor);
        saidaButton.setPreferredSize(new Dimension(200, 30));
        saidaButton.addActionListener(e -> selecionarHora(saidaButton, false));
        addComponent(panel, saidaButton, gbc, 1, 9, 3, 1);

        JLabel observacaoLabel = new JLabel("Observação");
        observacaoLabel.setFont(new Font("Arial", Font.BOLD, 15));
        addComponent(panel, observacaoLabel, gbc, 0, 10);

        observacaoField = new JTextField();
        observacaoField.setPreferredSize(new Dimension(200, 25));
        addComponent(panel, observacaoField, gbc, 1, 10, 3, 1);

        adicionarButton = new JButton("Adicionar");
        adicionarButton.setFont(buttonFont);
        adicionarButton.setBackground(buttonBackgroundColor);
        adicionarButton.setForeground(buttonForegroundColor);
        adicionarButton.setPreferredSize(new Dimension(150, 30));
        addComponent(panel, adicionarButton, gbc, 0, 11, 2, 1);

        gerarButton = new JButton("Gerar Planilha");
        gerarButton.setFont(buttonFont);
        gerarButton.setBackground(buttonBackgroundColor);
        gerarButton.setForeground(buttonForegroundColor);
        gerarButton.setPreferredSize(new Dimension(150, 30));
        addComponent(panel, gerarButton, gbc, 2, 11, 2, 1);

        resetarButton = new JButton("Resetar");
        resetarButton.setFont(buttonFont);
        resetarButton.setBackground(buttonBackgroundColor);
        resetarButton.setForeground(buttonForegroundColor);
        resetarButton.setPreferredSize(new Dimension(150, 30));
        addComponent(panel, resetarButton, gbc, 0, 12, 2, 1);

        imprimirButton = new JButton("Imprimir Planilha");
        imprimirButton.setFont(buttonFont);
        imprimirButton.setBackground(buttonBackgroundColor);
        imprimirButton.setForeground(buttonForegroundColor);
        imprimirButton.setPreferredSize(new Dimension(150, 30));
        addComponent(panel, imprimirButton, gbc, 2, 12, 2, 1);

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

        resetarButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    resetarPlanilha();
                } catch (IOException ex) {
                    ex.printStackTrace();
                }
            }
        });

        imprimirButton.addActionListener(new ActionListener() {
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
                    salvarEImprimirArquivo();
                } catch (IOException ex) {
                    ex.printStackTrace();
                }
            }
        });

        // Salvar as preferências ao fechar o programa
        addWindowListener(new java.awt.event.WindowAdapter() {
            @Override
            public void windowClosing(java.awt.event.WindowEvent windowEvent) {
                salvarPreferencias();
            }
        });
    }

    private void addComponent(JPanel panel, Component component, GridBagConstraints gbc, int x, int y) {
        addComponent(panel, component, gbc, x, y, 1, 1);
    }

    private void addComponent(JPanel panel, Component component, GridBagConstraints gbc, int x, int y, int width, int height) {
        gbc.gridx = x;
        gbc.gridy = y;
        gbc.gridwidth = width;
        gbc.gridheight = height;
        panel.add(component, gbc);
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

    private void selecionarHora(JButton button, boolean isEntrada) {
        SpinnerDateModel model = new SpinnerDateModel();
        JSpinner spinner = new JSpinner(model);
        JSpinner.DateEditor editor = new JSpinner.DateEditor(spinner, "HH:mm");
        spinner.setEditor(editor);

        int result = JOptionPane.showOptionDialog(null, spinner, "Selecione a Hora",
                JOptionPane.OK_CANCEL_OPTION, JOptionPane.QUESTION_MESSAGE, null, null, null);

        if (result == JOptionPane.OK_OPTION) {
            LocalTime time = LocalTime.parse(editor.getFormat().format(spinner.getValue()), DateTimeFormatter.ofPattern("HH:mm"));
            button.setText(time.toString());
            if (isEntrada) {
                if (button == entradaNormalButton) {
                    entradaNormal = time;
                } else {
                    entrada = time;
                }
            } else {
                if (button == saidaNormalButton) {
                    saidaNormal = time;
                } else {
                    saida = time;
                }
            }
        }
    }

    private void adicionarInformacoes() throws IOException {
        if (workbook == null || sheet == null) {
            abrirArquivoParaEdicao(caminhoArquivoField.getText());
        }

        String nome = nomeField.getText();
        String mes = (String) mesComboBox.getSelectedItem();
        String entradaNormalStr = entradaNormal != null ? entradaNormal.toString() : "";
        String saidaNormalStr = saidaNormal != null ? saidaNormal.toString() : "";
        String salarioStr = salarioField.getText().replace("R$", "").replaceAll("[,.]", "").trim();
        double salario = Double.parseDouble(salarioStr) / 100.0;

        int dia = (int) diaComboBox.getSelectedItem();
        String diaSemana = (String) diaSemanaComboBox.getSelectedItem();
        String data = String.format("%02d %s", dia, diaSemana);

        String entradaStr = entrada != null ? entrada.toString() : "";
        String saidaStr = saida != null ? saida.toString() : "";
        String observacao = observacaoField.getText();

        // Preencher as células com os dados fornecidos
        updateCell(sheet, 0, 3, nome); // Nome: D1
        updateCell(sheet, 1, 3, mes);  // Mês: D2
        updateCell(sheet, 1, 17, entradaNormalStr); // Horário de entrada normal: R2
        updateCell(sheet, 1, 18, saidaNormalStr);   // Horário de saída normal: S2
        updateCell(sheet, 2, 4, salario);        // Salário: E3

        // Calcular a linha correta com base no dia
        int linha = 5 + (dia - 1) * 2;

        // Preencher as células da linha calculada
        updateCell(sheet, linha, 2, data);       // Dia e dia da semana: C(linha)
        updateCell(sheet, linha, 3, entradaStr);    // Horário entrada: D(linha)
        updateCell(sheet, linha, 4, saidaStr);      // Horário saída: E(linha)
        updateCell(sheet, linha, 21, observacao); // Observação: V(linha)
    }

    private boolean entradasEstaoVazias() {
        return entradaButton.getText().equals("Selecione") &&
                saidaButton.getText().equals("Selecione") &&
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

    private void salvarEImprimirArquivo() throws IOException {
        // Forçar a recalculação de todas as fórmulas na planilha
        workbook.setForceFormulaRecalculation(true);
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        evaluator.evaluateAll();

        // Salvar o arquivo atualizado
        String caminhoArquivo = caminhoArquivoField.getText();
        File arquivo = new File(caminhoArquivo);
        try (FileOutputStream fileOut = new FileOutputStream(arquivo)) {
            workbook.write(fileOut);
        }

        // Imprimir o arquivo
        imprimirArquivo(arquivo);
    }

    private void limparCampos() {
        diaComboBox.setSelectedIndex(0);
        diaSemanaComboBox.setSelectedIndex(0);
        entradaButton.setText("Selecione");
        saidaButton.setText("Selecione");
        observacaoField.setText("");
    }

    private void resetarPlanilha() throws IOException {
        if (workbook == null || sheet == null) {
            abrirArquivoParaEdicao(caminhoArquivoField.getText());
        }

        for (int i = 5; i <= 65; i += 2) {
            clearCell(sheet, i, 2); // Limpar Dia e dia da semana: C(linha)
            clearCell(sheet, i, 3); // Limpar Horário entrada: D(linha)
            clearCell(sheet, i, 4); // Limpar Horário saída: E(linha)
            clearCell(sheet, i, 21); // Limpar Observação: V(linha)
        }

        // Salvar as mudanças após resetar, sem abrir o arquivo
        String caminhoArquivo = caminhoArquivoField.getText();
        try (FileOutputStream fileOut = new FileOutputStream(caminhoArquivo)) {
            workbook.write(fileOut);
        }
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

    private void clearCell(Sheet sheet, int rowIndex, int colIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row != null) {
            Cell cell = row.getCell(colIndex);
            if (cell != null) {
                cell.setCellType(CellType.BLANK);
            }
        }
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

    private void imprimirArquivo(File arquivo) {
        try {
            if (Desktop.isDesktopSupported()) {
                Desktop desktop = Desktop.getDesktop();
                if (desktop.isSupported(Desktop.Action.PRINT)) {
                    desktop.print(arquivo);
                } else {
                    JOptionPane.showMessageDialog(this, "Impressão não suportada no sistema.");
                }
            } else {
                JOptionPane.showMessageDialog(this, "Impressão não suportada no sistema.");
            }
        } catch (IOException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(this, "Erro ao imprimir o arquivo: " + e.getMessage());
        }
    }

    private void salvarPreferencias() {
        prefs.put("nome", nomeField.getText());
        prefs.put("entradaNormal", entradaNormal != null ? entradaNormal.toString() : "");
        prefs.put("saidaNormal", saidaNormal != null ? saidaNormal.toString() : "");
        prefs.put("salario", salarioField.getText());
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            PreencherPlanilhaGUI frame = new PreencherPlanilhaGUI();
            frame.setVisible(true);
        });
    }
}
