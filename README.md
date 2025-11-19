# 🔒 Excel Unprotector - Sistema de Remoção de Proteção do Microsoft Excel

[![Licença MIT](https://img.shields.io/badge/Licen%C3%A7a-MIT-blue.svg)](LICENSE)
[![Feito com Python](https://img.shields.io/badge/Feito%20com-Python-blue)](https://www.python.org/)

Um script poderoso e de código aberto desenvolvido em Python para automatizar a remoção de proteções comuns em arquivos do Microsoft Excel (.xlsx e .xlsm). Projetado para desenvolvedores e usuários que precisam gerenciar seus próprios arquivos protegidos ou esqueceram senhas em contextos legítimos.

## ✨ Funcionalidades em Destaque

| Tipo de Arquivo | Proteção Removida | Mecanismo | Saída |
| :---: | :---: | :--- | :--- |
| **.xlsm** | Proteção de Módulos VBA (Projetos de Código) | Manipulação do `vbaProject.bin` (Byte Substitution) | Arquivo original modificado, Backup criado. |
| **.xlsx** | Proteção de Planilhas e Bloqueio de Células | API da biblioteca `openpyxl` | Novo arquivo `_desprotegido.xlsx` criado, Backup criado. |

## 📐 Arquitetura do Sistema

O sistema é construído em torno de uma arquitetura modular, utilizando a força de bibliotecas padrão do Python para manipulação de arquivos binários e ZIP, e bibliotecas de terceiros para manipulação de formatos XML complexos como o Office Open XML.

### 1. Seleção e Dispatch
* A função `escolher_e_processar_arquivo()` inicializa uma interface gráfica minimalista (`tkinter`) para selecionar o arquivo.
* O fluxo é direcionado dinamicamente com base na extensão (`.xlsm` ou `.xlsx`), garantindo que apenas a rotina necessária seja executada.

### 2. Rotina VBA (.xlsm)
* Arquivos `.xlsm` são tratados como um contêiner ZIP (`zipfile`).
* A proteção é removida através de uma **substituição binária** (Byte Substitution): a *tag* de proteção `DPB=` (Designated Protected Binary) dentro do arquivo interno `xl/vbaProject.bin` é substituída pela *tag* neutra `DPx=`.
* Esta técnica é eficaz porque o Excel espera `DPB=` para iniciar a verificação de senha; ao encontrar `DPx=`, ele simplesmente ignora a proteção.

### 3. Rotina Planilha (.xlsx)
* A biblioteca **`openpyxl`** é usada para carregar a estrutura do arquivo.
* **Remoção de Proteção de Planilha:** O atributo booleano `sheet.protection.sheet` é explicitamente definido como `False` para cada folha de trabalho.
* **Remoção de Bloqueio de Célula:** O script itera sobre *todas* as células, definindo `cell.protection.locked = False`, o que remove a formatação que impede a edição após a desproteção da planilha.

## 🛠️ Instalação e Requisitos

### Pré-requisitos
Você precisa ter o **Python 3.x** instalado.

### Instalação das Dependências

O projeto depende da biblioteca `openpyxl`. Você pode instalá-la via `pip`:

```bash
pip install openpyxl