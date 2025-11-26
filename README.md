# 📊 Excel Data Appender (xlwings)

Uma ferramenta de automação em Python projetada para consolidar relatórios de ordens de serviço, copiando dados dinâmicos de planilhas diárias para um relatório mestre acumulado (`.xlsm`).

![Python](https://img.shields.io/badge/python-3.8%2B-blue)
![Library](https://img.shields.io/badge/lib-xlwings-orange)
![Status](https://img.shields.io/badge/status-production-green)

## 🎯 O Problema

Processos manuais de "Copiar e Colar" entre planilhas corporativas são propensos a erros humanos, consomem tempo e podem deslocar referências de células. Este script automatiza a transferência de dados garantindo que o "append" (anexação) ocorra sempre na próxima linha vazia disponível.

## ⚙️ Funcionalidades

* **Detecção Dinâmica de Intervalos:** O script calcula automaticamente a última linha preenchida tanto na planilha de origem (para saber o que copiar) quanto na de destino (para saber onde colar).
* **Manipulação de .XLSM:** Compatível com arquivos habilitados para macro do Excel.
* **Segurança de Processo:** Utiliza instâncias controladas do Excel para evitar processos "zumbis" ou conflitos com planilhas abertas pelo usuário.
* **Logs Detalhados:** Feedback visual no console sobre o progresso da operação.

## 🛠️ Pré-requisitos

* Microsoft Excel instalado (O `xlwings` requer uma instalação local do Excel para funcionar).
* Python 3.x.

### Instalação das Dependências

```bash
pip install xlwings
