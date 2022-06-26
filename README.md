Scrapybots Documentation
=========================
<!--ts-->
* [Maintained Projects](#maintained-projects)
* [Setup](#setup)
* [Webdriver](#webdriver)
    * [Selenium Webdriver Problems](#selenium-webdriver-problems)
<!--te-->

---

<h3 align="center"> 
<img alt="scrapybots banner" src="./assets/scrapybots-logo.PNG" width="1500" height="250">
</h3>

<p align="center">
  <a href="https://www.linkedin.com/in/vitoria-pecanha/">
    <img alt="Author" src="https://img.shields.io/badge/made%20by-Vitória Peçanha-white">
  </a>

  <img alt="Top language" src="https://img.shields.io/github/languages/count/vitoriape/scrapybots?color=violet">
  
  <img alt="Top language" src="https://img.shields.io/github/languages/top/vitoriape/scrapybots?color=002750">
  
  <img alt="Browser laucher" src="https://img.shields.io/badge/browser-Edge-cyan">

  <a href="https://github.com/vitoriape/scrapybots/blob/main/LICENSE">
    <img alt="License: GPL" src="https://img.shields.io/github/license/vitoriape/scrapybots?color=red">
  </a>
  
  <a href="https://github.com/vitoriape/scrapybots/commits/main">
    <img alt="Last commit" src="https://img.shields.io/github/last-commit/vitoriape/scrapybots?color=lightgrey">
  </a>
	
  <a href="https://github.com/vitoriape/scrapybots/archive/refs/heads/main.zip">
    <img alt="Repo size" src="https://img.shields.io/github/repo-size/vitoriape/scrapybots?color=yellow">
  </a>

  <a href="https://github.com/vitoriape/ConsprocBot/releases">
    <img alt="Status" src="https://img.shields.io/badge/status-ongoing-black">
  </a>
</p>

---
---

## **Maintained Projects**
<p align="left"> 
	<a href="https://progress-bar.dev/100/">
  <img alt="Progress status" src="https://progress-bar.dev/100/"></a>
<p>
	
- [Consprocbot](./consprocbot/): Webscraping em sites municipais de Andamento de Processos, anexando os mesmos em bases de dados mantidas com Excel
- [Dwloopbot](./dwloopbot/): Download em massa de arquivos mantidos em páginas de cursos online

---

### **Setup**
Afim de facilitar o uso dos RPA's, adicione o executável Python que você utiliza dentro da variável `PATH` do seu sistema.

<table><thead><tr><th>Minimum</th><th>Recommended</th></tr></thead><tbody><tr><td>Requer um processador e um sistema operacional Windows ou Linux</td><td>Requer um processador e um sistema operacional Windows ou Linux</td></tr><tr><td>Kernel: <a href="https://www.python.org/downloads/release/python-370/" target="_blank" rel="noopener noreferrer">Python 3.7.0</a></td><td>Kernel: <a href="https://www.python.org/downloads/release/python-397/" target="_blank" rel="noopener noreferrer">Python 3.10.4</a></td></tr><tr><td>Libraries: pyautogui, selenium, openpyxl e os</td><td>Libraries: pyautogui, selenium, msedge.selenium_tools, openpyxl e os</td></tr><tr><td>Drivers: any webdriver</td><td>Drivers: Edge webdriver</td></tr></tbody></table>

<p align="center">
<img alt="Python Badge" src="https://img.shields.io/badge/Running on Windows-002750?style=for-the-badge&logo=python&logoColor=yellow" />
</p>

```cmd
python filename.py
```

---

### **Webdriver**
Faça o download do webdriver de acordo com o navegador de sua preferência e, para executar o mesmo através da biblioteca `selenium`, também adicione o mesmo na variável `PATH` do sistema, conforme orientado na seção [Setup](#setup). 

#### **Selenium Webdriver Problems**
Caso as configurações anteriores não funcionem e você comece a receber erros na execução do webdriver, faça o seguinte:

**I. Instale a biblioteca msedge.selenium_tools**
```cmd
pip install msedge.selenium_tools
```

**II. Importe as bibliotecas necessárias**
```python
from selenium import webdriver
from msedge.selenium_tools import EdgeOptions
from msedge.selenium_tools import Edge
```

**III. Reconfigure o local com o webdriver executável**
```python
edge_options = EdgeOptions()
edge_options.use_chromium = True
nav = Edge(executable_path=r'C:\localondeoarquivoestasalvo\msedgedriver.exe', 
            options=edge_options)
```

---
---