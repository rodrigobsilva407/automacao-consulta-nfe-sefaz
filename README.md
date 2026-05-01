# 🚀 Automação de Consulta de NF-e (SEFAZ-CE)

Projeto desenvolvido para automatizar a consulta de notas fiscais eletrônicas diretamente na API do portal SITRAM (SEFAZ-CE), com geração de relatório Excel profissional.

## 💡 Problema resolvido
Processos manuais de consulta de NF-e são lentos, sujeitos a erro e pouco escaláveis.

## ⚙️ Solução
Aplicação em Python com interface gráfica (PyQt6) que:

- Consulta múltiplas NF-es em paralelo (multithreading)
- Integra com APIs da SEFAZ
- Trata dados automaticamente
- Gera relatório Excel estruturado e formatado
- Apresenta logs, progresso e resumo da execução

## 🧠 Tecnologias utilizadas
- Python
- Pandas
- Requests
- PyQt6
- OpenPyXL
- Multithreading

## 📊 Principais funcionalidades
- Consulta em lote por chave de acesso
- Extração de:
  - Dados da nota
  - Lançamentos
  - Itens
- Geração de relatório com:
  - Abas organizadas
  - KPIs automáticos
  - Formatação profissional

## 🚀 Diferenciais
- Processamento paralelo (ganho de performance)
- Interface amigável (não técnico consegue usar)
- Tratamento robusto de erros (timeout, falha API, etc.)
- Automação completa do fluxo fiscal

## 📸 Demonstração
(Coloque prints aqui da interface)

## 📂 Como usar
1. Inserir arquivo com chaves
2. Executar aplicação
3. Gerar relatório automaticamente

---

Desenvolvido com foco em automação fiscal e análise de dados.
