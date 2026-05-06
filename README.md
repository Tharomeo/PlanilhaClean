<div align="center">

<table>
  <tr>
    <td><img src="https://img.icons8.com/fluency/64/microsoft-excel-2019.png" width="60"/></td>
    <td><h1>Planilha Clean</h1></td>
  </tr>
</table>

**Ferramenta desktop para limpeza e deduplicação de planilhas Excel e CSV**

![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white)
![Tkinter](https://img.shields.io/badge/Tkinter-007acc?style=for-the-badge&logoColor=white)
![Status](https://img.shields.io/badge/Status-Concluído-28a745?style=for-the-badge)

</div>

---

## 📌 Sobre o projeto

O **Planilha Clean** é uma aplicação desktop desenvolvida em Python que permite remover linhas duplicadas de arquivos Excel e CSV de forma visual, simples e sem precisar de conhecimento em programação.

O arquivo original **nunca é alterado** — o programa sempre gera uma cópia nova e limpa.

---

## ✨ Funcionalidades

- 📂 **Drag and drop** — arraste o arquivo direto para a janela
- 📑 **Suporte a múltiplas abas** — processa cada aba do Excel de forma independente
- ☑️ **Seleção de colunas** — escolha quais colunas usar como critério de duplicidade
- ⚙️ **Dois modos de limpeza** — Rigoroso e Flexível (detalhes abaixo)
- 🔍 **Simulação** — veja o relatório de quantas linhas serão removidas antes de salvar
- 💾 **Exportação segura** — gera um novo arquivo `.xlsx` sem modificar o original
- 🌙 **Interface dark mode** moderna

---

## 🚀 Como usar

### 1. Abrir o arquivo

Ao iniciar o app, você verá uma tela limpa com duas opções:

- **Arrastar e soltar** — arraste o arquivo Excel ou CSV direto para o quadrado central
- **Clicar para buscar** — clique no ícone 📂 para abrir o explorador de arquivos

**Formatos suportados:** `.xlsx`, `.xls`, `.csv`

---

### 2. Navegar pelas abas

Após o carregamento, o programa abre a tela de edição organizada em **abas** — uma para cada planilha do seu arquivo.

> ⚠️ A configuração é **individual por aba**. O que você define em "Clientes" não afeta "Vendas".

---

### 3. Configurar a limpeza

Para cada aba, siga dois passos:

**A) Escolha a Regra de Duplicidade**

| Modo | Como funciona | Ideal para |
|---|---|---|
| 🔒 **Rigoroso (E)** | Remove apenas se **todas** as colunas marcadas forem iguais | Linhas que são cópias exatas |
| 🔓 **Flexível (OU)** | Remove se **qualquer** coluna marcada for igual | Cadastros com mesmo CPF/e-mail mas nome diferente |

**B) Selecione as Colunas Critério**

Marque as colunas que o programa deve verificar (ex: Email, CPF, Telefone). Use os botões **Marcar Tudo** ou **Desmarcar Tudo** para agilizar.

> 💡 Colunas comuns como `email`, `cpf`, `cnpj` e `tel` são marcadas automaticamente.

---

### 4. Simular antes de apagar

Clique em **🔍 SIMULAR** para ver o relatório antes de confirmar:

```
📊 SIMULAÇÃO: Clientes
--------------------------
Total antes:   1.500
Removidas:       320
Restantes:     1.180
```

---

### 5. Salvar o arquivo final

Clique em **💾 SALVAR**, escolha o destino e pronto. O programa gera um **novo arquivo Excel** com todas as abas:

- **Abas configuradas** → limpas conforme suas escolhas
- **Abas sem seleção** → copiadas sem alteração

---

### 6. Recomeçar

Clique em **⬅ Voltar** para retornar à tela inicial e carregar outro arquivo.

---

### Resumo rápido

```
1. Arraste o arquivo
2. Navegue pelas abas
3. Defina Rigoroso ou Flexível
4. Marque as colunas-chave
5. Simule → Salve
```

---

## 🛠️ Instalação

**Pré-requisito:** Python 3.8 ou superior

```bash
# Clone o repositório
git clone https://github.com/Tharomeo/planilha-clean.git
cd planilha-clean

# Instale as dependências
pip install pandas openpyxl tkinterdnd2
```

```bash
# Execute
python planilha_clean.py
```

---

## 📦 Dependências

| Biblioteca | Uso |
|---|---|
| `pandas` | Leitura, processamento e exportação de planilhas |
| `openpyxl` | Escrita de arquivos `.xlsx` |
| `tkinter` | Interface gráfica (nativa do Python) |
| `tkinterdnd2` | Suporte a drag and drop de arquivos |

---

## 📁 Estrutura

```
planilha-clean/
├── planilha_clean.py   # Aplicação principal
└── README.md
```
