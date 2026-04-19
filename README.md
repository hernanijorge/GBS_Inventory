# GBS Inventory Manager

Sistema de controle de estoque de equipamentos da **Global Business Solution, Boston**.

Desenvolvido em **VB.NET + Oracle Database** seguindo arquitetura **MVC**.

---

## 📋 Requisitos

- **Visual Studio 2022** (Community ou superior) com workload de *Desenvolvimento .NET Framework*
- **.NET Framework 4.8**
- **Oracle Database** (XE 18c/21c ou superior, ou instância Oracle corporativa)
- Acesso de rede ao banco Oracle
- Microsoft Access Database Engine 2016 (ou superior) para importação de planilhas `.xlsx`

---

## 🏗️ Estrutura do projeto

```
GBS_Inventory/
├── GBS_Inventory.sln                    ← abrir este arquivo no Visual Studio
├── Database/                            ← scripts SQL de instalação
│   ├── 01_DDL_Tabelas.sql
│   ├── 02_DDL_Sequences.sql
│   ├── 03_PACK_EQUIPAMENTO_SPEC.sql
│   ├── 04_PACK_EQUIPAMENTO_BODY.sql
│   ├── 05_PACK_UPGRADE.sql
│   └── 06_PACK_REMESSA.sql
└── GBS_Inventory/                       ← projeto VB.NET
    ├── GBS_Inventory.vbproj
    ├── App.config                       ← ajustar connection string aqui
    ├── packages.config
    ├── My Project/
    ├── Oracle/
    │   └── OracleHelper.vb              ← utilitário de acesso ao Oracle
    ├── Models/                          ← camada M do MVC
    │   ├── Equipamento.vb
    │   ├── Upgrade.vb
    │   └── Remessa.vb
    ├── classes/                         ← acesso a dados (gravação + leitura)
    │   ├── clsGravacaoEquipamento.vb
    │   ├── clsGravacaoUpgrade.vb
    │   ├── clsGravacaoRemessa.vb
    │   ├── clsLeituraEquipamento.vb
    │   ├── clsLeituraUpgrade.vb
    │   ├── clsLeituraRemessa.vb
    │   └── clsImportacaoExcel.vb
    ├── Controllers/                     ← camada C do MVC
    │   ├── EquipamentoController.vb
    │   ├── UpgradeController.vb
    │   ├── RemessaController.vb
    │   └── ImportacaoController.vb
    ├── Utils/
    │   └── TemaEscuro.vb                ← paleta de cores e estilização
    └── Views/                           ← camada V do MVC (WinForms)
        ├── frmPrincipal.vb + Designer
        ├── frmScanner.vb + Designer
        ├── frmImportacao.vb + Designer
        ├── frmUpgrade.vb + Designer
        └── frmRemessa.vb + Designer
```

---

## 🚀 Instalação passo a passo

### 1. Preparar o banco Oracle

Crie um usuário/schema dedicado ao sistema:

```sql
-- Conectado como SYSTEM
CREATE USER gbs_user IDENTIFIED BY gbs_pass;
GRANT CONNECT, RESOURCE TO gbs_user;
GRANT CREATE TABLE, CREATE SEQUENCE, CREATE PROCEDURE, CREATE VIEW TO gbs_user;
ALTER USER gbs_user QUOTA UNLIMITED ON USERS;
```

### 2. Executar os scripts SQL

Conectado como `gbs_user`, execute na ordem:

1. `Database/01_DDL_Tabelas.sql`
2. `Database/02_DDL_Sequences.sql`
3. `Database/03_PACK_EQUIPAMENTO_SPEC.sql`
4. `Database/04_PACK_EQUIPAMENTO_BODY.sql`
5. `Database/05_PACK_UPGRADE.sql`
6. `Database/06_PACK_REMESSA.sql`

### 3. Abrir o projeto no Visual Studio

1. Abra `GBS_Inventory.sln` no Visual Studio 2022.
2. **Clique com botão direito no projeto → Gerenciar Pacotes NuGet**.
3. Instale `Oracle.ManagedDataAccess` (versão 21.x ou superior).
4. Compile o projeto (`Ctrl+Shift+B`).

### 4. Configurar a connection string

Abra `App.config` e ajuste:

```xml
<connectionStrings>
  <add name="OracleDB"
       connectionString="Data Source=localhost:1521/XE;User ID=gbs_user;Password=gbs_pass"
       providerName="Oracle.ManagedDataAccess.Client"/>
</connectionStrings>
```

### 5. Executar

Pressione **F5** ou clique em **Iniciar**.

---

## ✨ Funcionalidades

### Dashboard
- Total de unidades, em estoque, em boa condição
- Upgrades nos últimos 30 dias
- Remessas ativas

### Estoque
- Listagem completa com filtros
- Pesquisa por UID, serial, modelo, CPU
- Edição e exclusão

### Scanner UID
- Campo grande para leitura de código de barras ou digitação
- Exibe card completo com todas as specs do equipamento
- Botão rápido para registrar upgrade

### Importação de Planilha
- Lê todas as abas do arquivo `.xlsx`
- UPSERT automático (insere novos, atualiza existentes)
- Progress bar em tempo real
- Log detalhado de erros

### Upgrades
- Registra RAM, SSD, HDD, bateria, tela, teclado, cover, GPU
- Origem: remessa extra, compra nova, transferido, garantia
- Custo em USD, técnico, observações
- Atualiza automaticamente os campos RAM_GB e STORAGE_GB do equipamento

### Remessas
- Suporte a FedEx, UPS, USPS, DHL
- Outbound (envio) e Inbound (recebimento)
- Tracking number com URL direta de rastreamento
- Ao marcar como DELIVERED, os equipamentos da remessa são automaticamente marcados como SOLD

---

## 🎨 Arquitetura MVC

```
┌────────────────────────────────────────────────────────────┐
│  VIEW (WinForms)                                           │
│  frmPrincipal · frmScanner · frmImportacao · frmUpgrade   │
└──────────────────────┬─────────────────────────────────────┘
                       │ (eventos)
┌──────────────────────▼─────────────────────────────────────┐
│  CONTROLLER (lógica de negócio + transações)               │
│  EquipamentoController · UpgradeController · ...          │
└──────────────────────┬─────────────────────────────────────┘
                       │ (instancia)
┌──────────────────────▼─────────────────────────────────────┐
│  MODEL (entidades)         │  CLASSES DE DADOS             │
│  Equipamento · Upgrade ... │  clsGravacao* + clsLeitura*  │
└─────────────────────────┬──┴──────────────────────────────┘
                          │ (OracleHelper)
┌─────────────────────────▼──────────────────────────────────┐
│  ORACLE DATABASE                                           │
│  PACK_EQUIPAMENTO · PACK_UPGRADE · PACK_REMESSA           │
│  TBL_EQUIPAMENTO · TBL_EQUIPAMENTO_UPGRADE · TBL_REMESSA  │
└────────────────────────────────────────────────────────────┘
```

---

## 🎨 Paleta de cores

| Elemento           | Cor         | Hex       |
| ------------------ | ----------- | --------- |
| Fundo              | Quase preto | `#080c10` |
| Surface            | Cinza azul  | `#131920` |
| Surface (inputs)   | Mais claro  | `#1c242e` |
| Borda              | Sutil       | `#26303c` |
| Accent (destaque)  | Ciano/teal  | `#00e5a0` |
| Texto              | Quase branco| `#e8edf2` |
| Texto mutado       | Cinza       | `#8c9baa` |
| Verde (status OK)  |             | `#2ecc71` |
| Vermelho (erro)    |             | `#e74c3c` |
| Amarelo (destaque) |             | `#ffce54` |

---

## 📞 Suporte

Em caso de dúvidas ou problemas, entre em contato com a equipe de TI da GBS.

**Versão:** 1.0.0
**Data:** 2026
