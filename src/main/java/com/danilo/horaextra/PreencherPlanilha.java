package com.danilo.horaextra;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.io.*;
import java.sql.SQLOutput;
import java.util.Scanner;

public class PreencherPlanilha {
    public static void main(String[] args) {
        Scanner sc = new Scanner(System.in);

        System.out.println("Caminho: ");
        String caminho = sc.nextLine();

        System.out.println("Nome: ");
        String nome = sc.nextLine();

        System.out.println("Salário: ");
        Double salario = sc.nextDouble();

        sc.nextLine();

        System.out.println("Horário normal entrada (HH:mm): ");
        String entradaNormal = sc.nextLine();

        System.out.println("Horário normal saída (HH:mm): ");
        String saidaNormal = sc.nextLine();

        System.out.println("Mês: ");
        String mes = sc.nextLine();

        System.out.println("Dia e Dia da semana: ");
        String data = sc.nextLine();
        String[] partesData = data.split(" ");
        Integer dia = Integer.parseInt(partesData[0]);
        String diaSemana = partesData[1];
        System.out.println("Entrada: ");
        String entrada = sc.nextLine();

        System.out.println("Saída: ");
        String saida = sc.nextLine();

        System.out.println("Observação: ");
        String observacao = sc.nextLine();

        try {
            preencherPlanilha(caminho,
                    nome,
                    salario,
                    entradaNormal,
                    saidaNormal,
                    mes,
                    dia,
                    diaSemana,
                    entrada,
                    saida,
                    observacao);
            abrirArquivo(caminho);
        } catch (Exception e) {
            e.printStackTrace();
        }

        sc.close();
    }

    private static void preencherPlanilha(String caminho, String nome, Double salario, String entradaNormal, String saidaNormal, String mes, Integer dia, String diaSemana, String entrada, String saida, String observacao) {
        FileInputStream fileInputStream;
        Workbook workbook;
        Sheet sheet;
        try {
            fileInputStream = new FileInputStream(caminho);
            workbook = new XSSFWorkbook(fileInputStream);
            sheet = workbook.getSheetAt(0);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        sheet.getRow(0).getCell(3).setCellValue(nome);
        sheet.getRow(1).getCell(3).setCellValue(mes);
        sheet.getRow(1).getCell(17).setCellValue(entradaNormal);
        sheet.getRow(1).getCell(18).setCellValue(saidaNormal);
        sheet.getRow(2).getCell(4).setCellValue(salario);

        int linha = 5 + (dia - 1) * 2;

        sheet.getRow(linha).getCell(2).setCellValue(dia + " " + diaSemana);
        sheet.getRow(linha).getCell(3).setCellValue(entrada);
        sheet.getRow(linha).getCell(4).setCellValue(saida);
        sheet.getRow(linha).getCell(21).setCellValue(observacao);

        try (FileOutputStream fileOutputStream = new FileOutputStream(caminho)) {
            fileInputStream.close();
            workbook.write(fileOutputStream);
            workbook.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

    private static void abrirArquivo(String caminho) {

        try {
            File arquivo = new File((caminho));
            if (arquivo.exists()) {
                if (Desktop.isDesktopSupported()) {
                    Desktop.getDesktop().open(arquivo);
                }
                else {
                    System.out.println("Abertura de arquivos não suportada no sistema.");
                }
            }
            else {
                System.out.println("Arquivo não encontrado");
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

}
