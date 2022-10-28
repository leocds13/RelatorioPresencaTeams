# Unificador de Relatório de presença do Teams

Projeto em Python para unificar relatórios de presença do Teams.
Tive a necessidade de gerar uma planilha com a presença dos participantes em uma serie de chamadas de um treinamento,
e o Relatório gerado pelo Teams tráz somente informações de uma unica chamada e não do treinamento inteiro.

Criei então este programinha que unifica tudo pra mim em uma planilha, bem mais facil de manipular depois caso precise de mais informações.

## Requisitos
Python
bibliotecas adicionais:
  pandas
  openpyxl

## Como Usar
1º Baixe todos os relatórios de presença do Teams em uma unica pasta
2º Baixe o código
3º Execute via terminal passando os parametros como no exemplo abaixo:
  python .\UnificadorDeRelatorioPresencaTeams.py -r <Diretório onde estão os relatórios> -o <Nome do arquivo xlsx gerado ao final>

---------

usage: UnificadorDeRelatorioPresencaTeams.py [-h] [-r REPORTS] [-p PATTERN] [-o OUTPUT]

Programa que junta varios relatórios de presença do teams em um unico excel

options:
  -h, --help            show this help message and exit
  -r REPORTS, --Reports REPORTS
                        Diretório que está localizado os relatórios de presença do teams
  -p PATTERN, --Pattern PATTERN
                        Regex usado para localizar os arquivos de relatórios no diretório
                        especificado
  -o OUTPUT, --Output OUTPUT
                        Nome do arquivo xlsx gerado ao final
