# 📊 Analisador de Parcelamentos - RPA Fiscal

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![Tkinter](https://img.shields.io/badge/Tkinter-GUI-orange.svg)](https://docs.python.org/3/library/tkinter.html)
[![Pandas](https://img.shields.io/badge/Pandas-Data%20Analysis-green.svg)](https://pandas.pydata.org/)
[![PDFPlumber](https://img.shields.io/badge/PDFPlumber-PDF%20Processing-red.svg)](https://github.com/jsvine/pdfplumber)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Uma aplicação desktop robusta desenvolvida em Python para automatizar a análise de relatórios fiscais da Receita Federal. Processa PDFs de parcelamentos e débitos, extrai dados estruturados, gera relatórios detalhados e fornece dashboards interativos para tomada de decisão.

## ✨ Funcionalidades Principais

### 🔍 Processamento Inteligente de PDFs
- **Extração automática** de dados de parcelamentos (PARCMEI, PARCSN, SIEFPAR, SISPAR, SICOB, SIPADE)
- **Identificação de débitos pendentes** com valores e status
- **Reconhecimento de CNPJ e nomes de empresas** via regex avançado
- **Suporte a múltiplos tipos de parcelamento** da Receita Federal e PGFN

### 📊 Análise e Filtragem Avançada
- **Filtros dinâmicos** por empresa, tipo, status, valor mínimo e atrasos
- **Ordenação inteligente** por qualquer coluna da tabela
- **Busca em tempo real** para localização rápida de registros
- **Agrupamento opcional** por empresa nos relatórios

### 📈 Dashboard Interativo
- **Estatísticas visuais** com cards coloridos (total empresas, parcelamentos, valores)
- **Ranking de empresas críticas** com mais parcelas em atraso
- **Top 10 maiores valores** de débitos suspensos
- **Distribuição por tipo** com gráficos visuais e percentuais

### 💾 Exportação e Backup
- **Relatórios em Excel** com formatação profissional
- **Backup automático em JSON** para auditoria
- **Exportação de dados filtrados** para análise externa
- **Cópia para clipboard** de registros selecionados

### 🎯 Interface Amigável
- **GUI moderna** com abas organizadas (Configuração, Resultados, Dashboard)
- **Progress bar** e logs detalhados durante processamento
- **Menus contextuais** para ações rápidas (copiar CNPJ, abrir PDF)
- **Configurações salvas** automaticamente

## 🛠️ Pré-requisitos

- **Python 3.8 ou superior**
- **Bibliotecas Python:**
  - `tkinter` (incluído no Python padrão)
  - `pandas`
  - `pdfplumber`
  - `openpyxl` (para exportação Excel)

## 📦 Instalação

1. **Clone ou baixe** o repositório:
   ```bash
   git clone https://github.com/seu-usuario/analisador-parcelamentos.git
   cd analisador-parcelamentos
   ```

2. **Instale as dependências**:
   ```bash
   pip install pandas pdfplumber openpyxl
   ```

3. **Execute a aplicação**:
   ```bash
   python AnalyizeV1.0.py
   ```

   Ou use o arquivo batch incluído:
   ```bash
   run_analyize.bat
   ```

## 🚀 Como Usar

### 1. Configuração Inicial
- **Pasta dos PDFs**: Selecione o diretório contendo os relatórios fiscais em PDF
- **Excel de Empresas** (opcional): Arquivo Excel com lista de CNPJs para filtrar apenas empresas específicas
- **Pasta de Saída**: Local onde serão salvos os relatórios gerados

### 2. Opções Avançadas
- ✅ **Incluir detalhes de débitos**: Adiciona informações sobre débitos pendentes
- ✅ **Agrupar por empresa**: Organiza resultados agrupados por empresa
- ✅ **Salvar backup JSON**: Mantém cópia dos dados em formato JSON

### 3. Processamento
- Clique em **"🔄 Processar PDFs"** para iniciar a análise
- Acompanhe o progresso na barra de status e logs em tempo real
- Resultados são automaticamente carregados nas abas de Resultados e Dashboard

### 4. Análise de Resultados
- **Aba Resultados**: Tabela interativa com filtros e ordenação
- **Aba Dashboard**: Visão geral com estatísticas e rankings
- **Exportar**: Use os botões para gerar relatórios Excel

## 📁 Estrutura do Projeto

```
AnalyizeV1.0.py          # Aplicação principal
run_analyize.bat         # Script de execução para Windows
README.md               # Documentação do projeto
requirements.txt        # Dependências Python (opcional)
```

## 🔧 Arquitetura Técnica

### Classes Principais
- **`AnalisadorParcelamentos`**: Classe principal que gerencia toda a aplicação
  - Interface GUI com Tkinter
  - Processamento de PDFs com pdfplumber
  - Análise de dados com pandas
  - Exportação e backup de resultados

### Algoritmos de Extração
- **Regex avançado** para identificação de padrões em PDFs
- **Normalização de CNPJ** para consistência de dados
- **Extração de valores monetários** com tratamento de formatos diversos
- **Mapeamento de tipos** de parcelamento (MEI, Simples Nacional, etc.)

### Interface Gráfica
- **Notebook com abas** para organização modular
- **Treeview widgets** para tabelas interativas
- **Canvas com scroll** para dashboards extensos
- **Threading** para processamento assíncrono sem travar a interface

## 📊 Tipos de Parcelamento Suportados

| Tipo | Descrição | Origem |
|------|-----------|--------|
| **PARCMEI** | Microempreendedor Individual | Receita Federal |
| **PARCSN** | Simples Nacional | Receita Federal |
| **SIEFPAR** | Exigibilidade Suspensa | Receita Federal |
| **SISPAR** | Exigibilidade Suspensa | PGFN |
| **SICOB** | Débito Suspenso | Receita Federal |
| **SIPADE** | Parcelamento SIPADE | Receita Federal |
| **DÉBITO** | Débitos Pendentes | Receita Federal |

## 🤝 Contribuição

Contribuições são bem-vindas! Para contribuir:

1. Fork o projeto
2. Crie uma branch para sua feature (`git checkout -b feature/AmazingFeature`)
3. Commit suas mudanças (`git commit -m 'Add some AmazingFeature'`)
4. Push para a branch (`git push origin feature/AmazingFeature`)
5. Abra um Pull Request

## 📝 Licença

Este projeto está licenciado sob a Licença MIT - veja o arquivo [LICENSE](LICENSE) para detalhes.

## 🆘 Suporte

Para suporte ou dúvidas:
- Abra uma issue no GitHub
- Verifique os logs da aplicação para diagnóstico
- Certifique-se de que os PDFs estão no formato esperado da Receita Federal

## 🔄 Versão

**v2.0** - Análise completa com dashboard interativo
- Interface redesenhada com abas
- Dashboard com estatísticas visuais
- Filtros avançados e busca em tempo real
- Exportação aprimorada
- Suporte a mais tipos de parcelamento

---

**Desenvolvido com ❤️ para automatizar processos fiscais**