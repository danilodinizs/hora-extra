package com.danilo.horaextra;

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
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
