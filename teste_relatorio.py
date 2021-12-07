# -*- coding: utf-8 -*-
from relatorio import Relatorio

if __name__ == "__main__":
    num_laudo = "165/2021"
    num_sei = "00000.00000-23"
    solicitante = 'Hygor Costa'
    sistema = 'Adutora Agreste'
    tipo = 1
    area = 200
    valor = 'R$ 1.000,00'
    solicitante = [num_laudo, num_sei, solicitante, sistema, tipo, area, valor]
    nome = 'Pedro'
    cpf = '079.938.334-32'
    proprietario = [nome, cpf]
    endereco = 'Rua 01'
    municipio = 'Recife'
    cep = '21030-494'
    imovel = [endereco, municipio, cep]
    desc_regiao = 'rica'
    desc_area = 'rural'
    desempenho = ['médio', 'baixo', 'médio', 'baixo' ]
    fundamentacao = [3, 2, 3, 2, 2, 2]
    precisao = 10.14
    num_variaveis = 2
    avaliacao = [desc_regiao, desc_area, desempenho, fundamentacao, precisao, num_variaveis]

    rel = Relatorio(solicitante, proprietario, imovel, avaliacao)
    rel.gerar_relatorio()