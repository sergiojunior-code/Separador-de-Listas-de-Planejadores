# 📄 Separador de Listas de Planejadores

Aplicação desktop em C# (Windows Forms) para **processamento automático de pedidos**, leitura de dados de Excel e separação de páginas de PDF por ordem de compra, organizando arquivos em diretórios específicos e gerando um relatório final.

---

## 🚀 Funcionalidades

* 📥 Leitura de arquivo Excel (.xlsx / .xls)
* 📄 Leitura e análise de PDF
* 🔍 Identificação de:

  * Ordem de compra
  * Ambiente
  * Planejadores
* ✂️ Extração de páginas específicas do PDF
* 📁 Organização automática em pastas de destino
* 📊 Geração de um novo Excel com:

  * Ordem de compra
  * Subplanejadores
  * Planejadores encontrados
* ⚠️ Relatório de pedidos não encontrados

---

## 🧠 Como funciona

1. O usuário carrega um arquivo Excel contendo os pedidos
2. O sistema valida os dados (formato esperado: `XXXX - Nome`)
3. O usuário carrega um PDF com múltiplas ordens
4. O sistema:

   * Percorre cada página do PDF
   * Identifica pedidos e planejadores
   * Extrai páginas correspondentes
   * Salva PDFs separados por pedido
5. Ao final:

   * Gera um Excel consolidado
   * Mostra resumo no painel

---

## 📂 Estrutura esperada

O sistema busca automaticamente diretórios no caminho:

```
J:\Pedidos [ANO]
```

Exemplo:

```
J:\Pedidos 2026
J:\Pedidos 2025
```

E dentro:

```
[Contrato]\[Ambiente]
```

---

## 🧾 Formato do Excel de entrada

* Coluna utilizada: **Coluna D**
* Formato obrigatório:

```
12345 - Nome do Ambiente
```

---

## 🧱 Tecnologias utilizadas

* C# (.NET Windows Forms)
* ClosedXML (manipulação de Excel)
* iText7 (leitura e manipulação de PDF)

---

## 📌 Classes principais

### `Armazenar`

Responsável por guardar:

* Ordem de compra
* Subplanejadores
* Planejadores

---

## ⚠️ Validações e Regras

* Excel deve conter dados válidos na coluna D
* Cada item deve conter `" - "` no texto
* PDF deve conter as ordens correspondentes
* Caso não encontre a pasta:

  * Uma nova será criada no diretório do PDF

---

## 📊 Saída gerada

### PDFs

Arquivos separados por ordem:

```
[OrdemCompra] - Lista de Planejadores.pdf
```

### Excel final

```
Pedidos e Planejadores.xlsx
```

Contendo:

* Ordem de Compra
* Sub Planejadores
* Planejadores

---

## ❗ Possíveis erros

* Arquivo Excel vazio ou inválido
* PDF sem correspondência
* Estrutura de pastas não encontrada
* Problemas de leitura de célula

---

## ✅ Fluxo de uso

1. Clique em **Carregar Excel**
2. Clique em **Carregar PDF**
3. Clique em **Iniciar**
4. Aguarde o processamento
5. Verifique:

   * PDFs gerados
   * Excel final
   * Log no painel

---

## 📝 Observações

* O sistema detecta automaticamente o **ano do pedido** baseado no PDF
* Subplanejadores são reduzidos para siglas
* Evita duplicação de planejadores

---

## 👨‍💻 Autor

Projeto desenvolvido para automação de processos internos de separação de pedidos e planejadores.

---

## 📌 Melhorias futuras

* Interface mais moderna
* Barra de progresso
* Log em arquivo externo
* Suporte a múltiplos PDFs
* Configuração dinâmica de diretórios

---
