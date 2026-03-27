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

* Separação por módulos (Clientes, Veículos, OS, Itens, Pagamentos e Utilitários)
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

* Busca por ID e documento
* Sanitização de CPF/CNPJ (remoção de caracteres)
* Validação de CPF
* Verificação de duplicidade

---

### 🚗 Gestão de Veículos

* Busca por ID
* Busca por placa (normalizada: sem hífen e em maiúsculo)
* Vinculação com cliente

---

### 📋 Ordem de Serviço

* Criação de OS com validação de cliente e veículo
* Busca por ID
* Controle de status
* Bloqueio de operações inválidas

---

### 🧾 Itens de Serviço

* Adição de itens à OS
* Validação da existência da OS
* Bloqueio de edição em OS finalizada ou cancelada

---

### 💰 Pagamentos

* Registro de pagamentos vinculados à OS
* Consulta por ID e por OS

---

## 🔒 Validações implementadas

* Limpeza automática:

  * CPF
  * CNPJ
  * Telefone
  * Placa

* Validação de:

  * CPF (estrutura e dígitos verificadores)
  * Existência de cliente e veículo antes de criar OS

* Regras de negócio:

  * Não é possível criar OS sem cliente ou veículo válido
  * Não é possível adicionar itens em OS finalizada ou cancelada
  * IDs são gerados automaticamente com base nas tabelas

---

## 📁 Estrutura do Projeto

```bash
📁 Sistema-OS-Oficina
│
├── 📁 vba
│   ├── modClientes.bas
│   ├── modVeiculos.bas
│   ├── modOS.bas
│   ├── modItensOS.bas
│   ├── modPagamentos.bas
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

3. Utilizar as funcionalidades implementadas

---

## 💡 Aprendizados

Este projeto envolveu:

* Modelagem de dados no Excel
* Simulação de banco relacional
* Organização modular em VBA
* Validação e sanitização de dados
* Implementação de regras de negócio
* Estruturação de CRUD
* Versionamento com Git/GitHub

---

## 📈 Próximas melhorias

* Interface com UserForms
* Impressão de ordem de serviço
* Relatórios automáticos
* Dashboard gerencial

---

## 👨‍💻 Autor

Carlos Adrians  
Estudante de Sistemas de Informação

---

## 📌 Observação

Projeto desenvolvido com base em uma necessidade real da oficina familiar, com foco em aprendizado prático e aplicação de conceitos de desenvolvimento de sistemas.
