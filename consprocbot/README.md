Consprocbot Documentation
=========================
<!--ts-->
* [Getting started with ConsprocBot](#getting-started-with-consprocbot)
  * [Requirements](#requirements)
* [Setup](#setup)
  * [Webdriver](#webdriver)
  * [Libraries](#libraries)
  * [Data](#data)
  * [Running](#running)
* [Features](#features)
  * [Consprocbot.v.2.2.ipynb](#consprocbot.v.2.2.ipynb)
  * [ProcessosConsproc.xlsx](#processosconsproc.xlsx)
* [References](#references)
* [Updates](#updates)
  * [21 de maio](#21-de-maio)
  * [26 de dezembro](#26-de-dezembro)
  * [23 de dezembro](#23-de-dezembro)
  * [18 de dezembro](#18-de-dezembro)
<!--te-->

---

<h3 align="center"> 
<img alt="ConsprocBot banner" src="./assets/consprocb-banner.png" width="1500" height="150">
</h3>

<p align="center">
  <a href="https://www.linkedin.com/in/vitoria-pecanha/">
    <img alt="Author" src="https://img.shields.io/badge/made%20by-Vitória Peçanha-white">
  </a>
  
  <img alt="Top language" src="https://img.shields.io/github/languages/top/vitoriape/ConsprocBot?color=orange">
  
  <img alt="Browser laucher" src="https://img.shields.io/badge/browser-Edge-cyan">

  <a href="https://github.com/vitoriape/ConsprocBot/blob/main/LICENSE">
    <img alt="License: GPL" src="https://img.shields.io/github/license/vitoriape/ConsprocBot?color=red">
  </a>
  
  <a href="https://github.com/vitoriape/ConsprocBot/commits/main">
    <img alt="Last commit" src="https://img.shields.io/github/last-commit/vitoriape/ConsprocBot?color=lightgrey">
  </a>
	
  <a href="https://github.com/vitoriape/ConsprocBot/archive/refs/heads/main.zip">
    <img alt="Repo size" src="https://img.shields.io/github/repo-size/vitoriape/ConsprocBot?color=yellow">
  </a>

  <a href="https://github.com/vitoriape/ConsprocBot/releases">
    <img alt="Repo updates" src="https://img.shields.io/github/release-date/vitoriape/ConsprocBot?color=green">
  </a>
	
  <a href="https://github.com/vitoriape/ConsprocBot/releases">
    <img alt="Repo last release" src="https://img.shields.io/github/v/release/vitoriape/ConsprocBot?color=brightgreen">
  </a>

  <a href="https://github.com/vitoriape/ConsprocBot/releases">
    <img alt="Status" src="https://img.shields.io/badge/status-finish-blue">
  </a>
</p>

---
---

## **Getting started with ConsprocBot**
ConsprocBot é um software de automação criado com Python. Ele foi desenvolvido para realizar webscraping em sites municipais de Andamento de Processos, anexando os mesmos em bases de dados mantidas com Excel.

### **Requirements**
<table><thead><tr><th>Minimum</th><th>Recommended</th></tr></thead><tbody><tr><td>Requer um processador e um sistema operacional Windows ou Linux</td><td>Requer um processador e um sistema operacional Windows ou Linux</td></tr><tr><td>Kernel: <a href="https://www.python.org/downloads/release/python-370/" target="_blank" rel="noopener noreferrer">Python 3.7.0</a></td><td>Kernel: <a href="https://www.python.org/downloads/release/python-397/" target="_blank" rel="noopener noreferrer">Python 3.10.4</a></td></tr><tr><td>Libraries: pyautogui, selenium, openpyxl e os</td><td>Libraries: pyautogui, selenium, openpyxl e os</td></tr><tr><td>Drivers: any webdriver</td><td>Drivers: Edge webdriver</td></tr></tbody></table>

---

#### **Guia de Instalação das Bibliotecas**
<!--ts-->
* [Instalando pyautogui](#instalando-pyautogui)
* [Instalando selenium](#instalando-selenium)
* [Instalando openpyxl](#instalando-openpyxl)
<!--te-->

---

## **Setup**
O ConsprocBot é executável através do Python, sendo necessário identificar ao software utilizado em sua execução (seja o `Jupyter Notebook`, `Jupyter Lab`, `Visual Studio Code` ou afins) o local onde o kernel Python se encontra, ou seja, o arquivo `python.exe`.

**Para consultar a documentação dos softwares requeridos ou instalar os mesmos, confira a seção [References](#references).**

### **Webdriver**
Para que o software seja capaz de controlar o **navegador web** do usuário, o script precisa de um webdriver correspondente com o mesmo. Até o momento, o script deste projeto contempla **apenas** o navegador **Edge**.  

Para que a aplicação reconheça o _webdriver_ escolhido, após o download, basta descompactar o arquivo `edgedriver.zip` e copiar o executável `msedgedriver.exe` para dentro da pasta do **Anaconda** no sistema.

### **Libraries**
O ConsprocBot utiliza as bibliotecas Python listadas na coluna **Recommended** na seção [Requeriments](#Requeriments). Caso não possua alguma instalada, siga os passos nos tutoriais abaixo:

#### **Instalando pyautogui**
A biblioteca `pyautogui` é um módulo de automação de GUI para Python2 e Pyhon3.

```cmd
py -m pip install pyautogui
```

ou

```cmd
pip install pyautogui
```

---

#### **Instalando selenium**
A biblioteca `selenium` é um `framework` para teste de aplicativos web. 

```cmd
py -m pip install selenium
```

ou

```cmd
pip install selenium
```

---

#### **Instalando openpyxl**
A biblioteca `openpyxl` é uma biblioteca que permite a leitura e escrita em arquivos do `Excel 2010`.

```cmd
py -m pip install openpyxl
```

ou

```cmd
pip install openpyxl
```
---

### **Data**
Antes da execução do RPA, é necessário preencher os dados na guia Processos da planilha [ProcessosConsproc.xlsx](#processosconsproc.xlsx):

<table><thead><tr><th>Prefixo</th><th>Ano</th></tr></thead><tbody><tr><td></td><td></td></tr></tbody></table>

---

### **Running**
<p align="center">
<img alt="Python Badge" src="https://img.shields.io/badge/Run on Python-002750?style=for-the-badge&logo=python&logoColor=yellow" />
</p>

Using Windows CMD
```cmd
python consprocbot.v.3.1.py
```

---

## **Features**
<p align="left"> 
	<a href="https://progress-bar.dev/100/">
  <img alt="Progress status" src="https://progress-bar.dev/100/"></a>
<p>

### **consprocbot.v.3.1.py**
RPA que minera dados de processos municipais e armazena as informações em arquivos de extensão `xlsx`, o utilitário [ProcessosConsproc.xlsx](#processosconsproc.xlsx) do projeto. O software varre as linhas armazenadas na sheet `Processos`, que precisa ser preenchida **antes** de rodar a aplicação:

- [x] **Mapeamento dos processos em guia**
- [x] **Possibilidade de adaptação para outros sites e navegadores**


### **ProcessosConsproc.xlsx**
Data source para **input e output** dos dados minerados pelo [consprocbot.v.3.1.py](#consprocbotv31py):

- [x] **Sheet utilitária de data source para o bot (`Processos`)**
- [x] **Sheet de armazenamento de dados minerados pelo bot (`Andamento`)**

---

## **References**
<table><thead><tr><th>Software Docs</th><th>Download Softwares</th></tr></thead><tbody><tr><td><a href="https://docs.python.org/3/" target="_blank" rel="noopener noreferrer">Python 3.10.4 documentation</a></td><td><a href="https://www.python.org/downloads/" target="_blank" rel="noopener noreferrer">Python</a></td></tr><tr><td><a href="https://docs.microsoft.com/pt-br/microsoft-edge/webdriver-chromium/" target="_blank" rel="noopener noreferrer">Use o WebDriver para automatizar o Microsoft Edge</a></td><td><a href="https://docs.microsoft.com/pt-br/microsoft-edge/webdriver-chromium/?tabs=c-sharp#download-microsoft-edge-driver" target="_blank" rel="noopener noreferrer">Download Microsoft Edge Webdriver</a></td></tr></tbody></table>

---
---

## **Updates**
### **Maio de 2022:**
### **21 de maio:**
- `cpb.v3.1:` otimização de script eliminando depreciações no código. Migração das dependências do projeto para executáveis diretamente em `Python`.

### **Dezembro de 2021:**
#### **26 de dezembro:**
- `cpb.v2.2:` incorporação do guia de funcionamento ao próprio código fonte do script. Pequena otimização do projeto escalando as tarefas em células distintas.

#### **23 de dezembro:**
- `cpb.v2.1:` atualização de código-fonte trazendo refinamento ao software. Os processos não precisam mais serem alocados dentro do próprio _script_, utilizando um sistema otimizado de retroalimentação de dados. Ele utiliza as seguintes bibliotecas: `pyautogui`, `selenium` e `openpyxl`

#### **18 de dezembro:**
- `cpb.v1.1:` em sua primeira versão estável, o consprocbot é capaz de minerar dados _web_ com auxílio de um `webdriver` do **Edge** a partir dos processos identificados dentro do seu próprio código fonte, e usar os mesmos para atualizar um arquivo `xlsx`. Ele utiliza as seguintes bibliotecas: `pyautogui`, `selenium` e `xlsxwriter`
