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
