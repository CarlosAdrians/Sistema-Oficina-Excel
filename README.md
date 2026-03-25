# 🛠️ Sistema de Ordem de Serviço para Oficina (Excel + VBA)

## 📌 Sobre o projeto

Este projeto é um sistema completo de gerenciamento de ordem de serviço desenvolvido em **Microsoft Excel com VBA**, voltado para uma oficina mecânica com foco em direção hidráulica.

A aplicação substitui o uso de talões manuais por uma solução digital estruturada, permitindo melhor controle de clientes, veículos, serviços e pagamentos.

---

## 🎯 Objetivo

Automatizar e organizar o fluxo da oficina, permitindo:

* Cadastro estruturado de clientes
* Controle de veículos vinculados
* Gerenciamento de ordens de serviço
* Registro detalhado de itens de serviço
* Controle financeiro (pagamentos e adiantamentos)

---

## 🧠 Tecnologias utilizadas

* Microsoft Excel (.xlsm)
* VBA (Visual Basic for Applications)

---

## 🏗️ Arquitetura do Sistema

O sistema foi estruturado seguindo princípios de organização semelhantes a aplicações reais:

* Separação por módulos (Clientes, Veículos, Utilitários)
* Funções reutilizáveis
* Validações centralizadas
* Uso de tabelas como banco de dados (ListObject)

---

## 🗄️ Estrutura do Banco de Dados

### 👤 Clientes (`tbClientes`)

```text
ID_CLIENTE | NOME | TIPO_PESSOA | CPF | CNPJ | TELEFONE | ENDERECO | DATA_CADASTRO
```

---

### 🚗 Veículos (`tbVeiculos`)

```text
ID_VEICULO | ID_CLIENTE | MARCA | MODELO | ANO | COR | PLACA
```

---

### 📋 Ordens de Serviço (`tbOs`)

```text
ID_OS | DATA_ABERTURA | DATA_FECHAMENTO | ID_CLIENTE | ID_VEICULO
OBSERVACOES | VALOR_TOTAL | STATUS
```

---

### 🧾 Itens da OS (`tbItensOs`)

```text
ID_ITEM | ID_OS | TIPO | QTD | DESCRICAO | VALOR_UNIT | TOTAL
```

---

### 💰 Pagamentos (`tbPagamentos`)

```text
ID_PAGAMENTO | ID_OS | TOTAL | ADIANTAMENTO | RESTANTE
DATA_PAGAMENTO | FORMA | STATUS_PAGAMENTO
```

---

## 🔗 Relacionamentos

* Cliente → possui vários veículos
* Cliente → possui várias ordens de serviço
* Veículo → vinculado a um cliente
* OS → possui vários itens
* OS → possui pagamentos

---

## ⚙️ Funcionalidades

### 👤 Gestão de Clientes

* Cadastro com validação completa
* CPF ou CNPJ (nunca ambos)
* Verificação de duplicidade
* Sanitização de dados (remoção de máscara)

---

### 🚗 Gestão de Veículos

* Cadastro vinculado ao cliente
* Controle por ID (chave estrangeira)
* Informações detalhadas do veículo

---

### 📋 Ordem de Serviço

* Abertura e fechamento de OS
* Vinculação com cliente e veículo
* Controle de status

---

### 🧾 Itens de Serviço

* Registro de peças e serviços
* Quantidade, valor unitário e total

---

### 💰 Pagamentos

* Controle de valores totais
* Registro de adiantamento
* Cálculo de valor restante
* Status do pagamento

---

## 🔒 Validações implementadas

* Limpeza automática:

  * CPF
  * CNPJ
  * Telefone

* Validação de:

  * Nome (obrigatório e tamanho mínimo)
  * CPF (11 dígitos)
  * CNPJ (14 dígitos)
  * Telefone (10 ou 11 dígitos)
  * Endereço obrigatório

* Regra de negócio:

  * Cliente deve possuir **CPF ou CNPJ (exclusivamente um)**

---

## 📁 Estrutura do Projeto

```bash
📁 Sistema-OS-Oficina
│
├── 📁 vba
│   ├── modClientes.bas
│   ├── modVeiculos.bas
│   ├── modUtilitarios.bas
│
├── Sistema_Oficina.xlsm
└── README.md
```

---

## 🚀 Como utilizar

1. Abrir o arquivo:

```bash
Sistema_Oficina.xlsm
```

2. Habilitar macros no Excel

3. Utilizar os módulos e funcionalidades implementadas

---

## 💡 Aprendizados

Este projeto envolveu:

* Modelagem de dados no Excel
* Simulação de banco relacional
* Organização modular em VBA
* Validação e sanitização de dados
* Estruturação de CRUD
* Versionamento com Git/GitHub

---

## 📈 Próximas melhorias

* Interface com UserForms
* Impressão de ordem de serviço
* Relatórios automáticos
* Dashboard gerencial
* Validação avançada de CPF/CNPJ
* Controle de estoque de peças

---

## 👨‍💻 Autor

Carlos Adrians
Estudante de Sistemas de Informação

---

## 📌 Observação

Projeto desenvolvido com base em uma necessidade real da oficina familiar, com foco em aprendizado prático e aplicação de conceitos de desenvolvimento de sistemas.
