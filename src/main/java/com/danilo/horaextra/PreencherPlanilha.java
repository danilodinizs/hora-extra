package com.danilo.horaextra;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
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
                    data,
                    entrada,
                    saida,
                    observacao);
        } catch (Exception e) {
            e.printStackTrace();
        }

        sc.close();
    }

    private static void preencherPlanilha(String caminho, String nome, Double salario, String entradaNormal, String saidaNormal, String mes, String data, Integer dia, String entrada, String saida, String observacao) {
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

        sheet.getRow(linha).getCell(2).setCellValue(nome);
        sheet.getRow(linha).getCell(3).setCellValue(nome);
        sheet.getRow(linha).getCell(4).setCellValue(nome);
        sheet.getRow(linha).getCell(21).setCellValue(nome);

        try(FileOutputStream fileOutputStream = new FileOutputStream(caminho)) {
            fileInputStream.close();
            workbook.write(fileOutputStream);
            workbook.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }
}
