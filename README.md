# Sistema de Endereçamento Logístico

Sistema desenvolvido em Python para gerar, cadastrar, consultar e controlar endereços de armazenagem.

## Objetivo

O projeto tem como objetivo apoiar operações logísticas no controle de posições de armazenagem, utilizando buffers, ruas e canalizações padronizadas.

## Estrutura dos endereços

Os endereços seguem o seguinte padrão:

B01-R01-BA02

Exemplo:

- B01: Buffer 01
- R01: Rua 01
- BA02: Canalização BA02

## Regras do sistema

- Buffers de B01 até B05
- Ruas de R01 até R20
- Canalizações com numeração par
- 500 endereços gerados automaticamente
- Endereços disponíveis e ocupados
- Controle da quantidade armazenada
- Proteção contra estoque negativo

## Funcionalidades

- Geração automática de endereços
- Cadastro de produtos
- Consulta por endereço ou produto
- Filtros por buffer, rua e status
- Entrada de estoque
- Saída de estoque
- Liberação manual de posições
- Liberação automática quando o estoque chega a zero
- Histórico de movimentações
- Dashboard logístico
- Exportação de relatórios para Excel

## Tecnologias utilizadas

- Python
- Streamlit
- Pandas
- OpenPyXL
- Plotly
- Excel
- Git
- GitHub

## Arquivos principais

- `app.py`: aplicação e interface do sistema
- `gerador_enderecos.py`: geração dos endereços logísticos
- `enderecos_logisticos.xlsx`: base atual dos endereços
- `historico_movimentacoes.xlsx`: histórico das operações
- `requirements.txt`: dependências do projeto

## Como executar

### 1. Instale as dependências

```bash
python -m pip install -r requirements.txt

## Imagens do sistema

### Tela inicial

imagens/tela_inicial.png

### Cadastro de produto

imagens/cadastro_produto.png

### Consulta de endereços

imagens/consulta_enderecos.png

### Movimentação de estoque

imagens/movimentacao_estoque.png

### Liberação de endereço

imagens/liberar_endereco.png

### Histórico de movimentações

![Hists/historico.png

### Dashboard logístico

imagens/dashboard.png

### Relatórios

imagens/relatorios.png

## Sobre o desenvolvedor

### Nicolas Henrique Lima Iansen

Estudante do Ensino Médio integrado ao curso técnico em Logística, com interesse em tecnologia aplicada às operações logísticas.

Atualmente, estudo e desenvolvo conhecimentos nas seguintes áreas:

- Python
- Excel
- Inteligência Artificial Generativa
- Cibersegurança
- Automação de processos
- Análise e organização de dados
- Controle de estoque
- Endereçamento logístico

Este projeto foi desenvolvido com o objetivo de unir conhecimentos de logística e tecnologia, criando uma solução para geração, controle e consulta de endereços de armazenagem.

Durante o desenvolvimento, pratiquei conceitos de programação, manipulação de planilhas, criação de interfaces, validação de dados, movimentação de estoque, construção de dashboards e geração de relatórios.

Busco continuar aprimorando meus conhecimentos e desenvolvendo projetos que possam contribuir para operações logísticas, análise de dados e tecnologia da informação.

## Competências demonstradas neste projeto

- Desenvolvimento de aplicações em Python
- Criação de interfaces com Streamlit
- Manipulação de dados com Pandas
- Criação e edição de planilhas com OpenPyXL
- Desenvolvimento de dashboards e indicadores
- Organização de endereços logísticos
- Controle de entradas e saídas de estoque
- Validação de informações
- Registro de histórico de movimentações
- Exportação de relatórios em Excel
- Documentação de projetos
- Uso do GitHub para portfólio

## Objetivo profissional

Desenvolver experiência nas áreas de tecnologia, logística, automação e análise de dados, aplicando conhecimentos de programação para solucionar problemas reais e melhorar processos operacionais.
