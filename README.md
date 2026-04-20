# GBS Inventory Manager

<p align="center">
  <img src="https://img.shields.io/badge/VB.NET-.NET%20Framework%204.8-512BD4?logo=.net&logoColor=white" alt=".NET">
  <img src="https://img.shields.io/badge/Oracle-XE%2011g%2B-F80000?logo=oracle&logoColor=white" alt="Oracle">
  <img src="https://img.shields.io/badge/IDE-Visual%20Studio%202022-5C2D91?logo=visualstudio&logoColor=white" alt="VS 2022">
  <img src="https://img.shields.io/badge/Arquitetura-MVC-success" alt="MVC">
  <img src="https://img.shields.io/badge/Status-beta-yellow" alt="Beta">
</p>

Sistema de controle de estoque de equipamentos refurbished (laptops, desktops, tablets) desenvolvido para a **Global Business Solution**, em Boston/MA.

Construído em **VB.NET + Oracle** seguindo arquitetura **MVC em pastas**, packages PL/SQL para a camada de negócio e interface WinForms em tema escuro.

---

## 📸 Interface

A aplicação possui 5 abas principais:

| Aba | Função |
|-----|--------|
| **Dashboard** | 5 cards com totais (estoque, condição boa, upgrades 30d, remessas ativas) |
| **Estoque** | Grid de equipamentos com busca por UID/serial/modelo + atalhos para Scanner e Importação |
| **Remessas** | Controle de envios FedEx/UPS/USPS/DHL com URL de rastreamento automática |
| **Upgrades** | Histórico de upgrades (RAM, SSD, bateria, etc) com origem da peça |
| **Importação** | Carrega planilha Excel da GBS (25 abas) e faz upsert automático |

O formulário `Scanner UID` abre em janela separada pelo botão na aba Estoque, com campo grande para leitura de código de barras e card de detalhes em tempo real.

---

## 🎯 O que este sistema faz

- ✅ Cadastro, edição e exclusão de equipamentos (laptops/desktops/tablets)
- ✅ **Scanner de Internal UID** — campo de alta performance para leitura de código de barras
- ✅ **Importação em massa** da planilha GBS (`.xlsx` com 25 abas) via UPSERT
- ✅ Registro de **upgrades de componentes** (RAM, SSD, HDD, bateria, tela, teclado, cover, GPU) com origem (compra nova, extra de remessa, transferência, garantia)
- ✅ Atualização **automática** dos campos RAM_GB/STORAGE_GB do equipamento ao registrar upgrade (lógica no package Oracle)
- ✅ Controle de **remessas** (FedEx, UPS, USPS, DHL) com tracking number e URL direta de rastreamento
- ✅ Cascata automática: ao marcar remessa como `DELIVERED`, os equipamentos dela mudam para `SOLD`
- ✅ Dashboard com 5 indicadores em tempo real

## 🚧 Roadmap (não implementado nesta versão)

Estes itens estão **planejados** mas ainda não fazem parte do código. A tabela `TBL_WHATSAPP_LOG` existe no schema para preparar o caminho.

- [ ] Bot WhatsApp para consulta remota do estoque (Twilio Sandbox)
- [ ] Integração eBay (criar/atualizar listings)
- [ ] Consulta de tracking em tempo real via APIs FedEx/UPS/USPS
- [ ] Versão mobile PWA
- [ ] Migração do driver para `Oracle.ManagedDataAccess` confirmada (já funciona)

---

## 📋 Requisitos

| Componente | Versão | Observação |
|------------|--------|------------|
| Visual Studio | 2022 (Community ou superior) | Workload *Desenvolvimento .NET desktop* |
| .NET Framework | 4.8 | — |
| Oracle Database | XE 11g ou superior | Também funciona em 18c/19c/21c |
| Oracle.ManagedDataAccess | 21.13.0+ | Instalado via NuGet, não requer Oracle Client |
| Microsoft Access Database Engine | 2016 x86 | Apenas para importação de planilhas .xlsx |

### Sistema operacional
Windows 10 / 11 (64 bits). A aplicação roda em AnyCPU.

---

## 🚀 Instalação rápida

### 1. Clonar o repositório

```bash
git clone https://github.com/SEU_USUARIO/gbs-inventory.git
cd gbs-inventory
```

### 2. Preparar o banco Oracle

Como `SYSTEM`, crie o usuário da aplicação:

```sql
CREATE USER gbs_user IDENTIFIED BY "gbs_pass"
    DEFAULT TABLESPACE USERS
    TEMPORARY TABLESPACE TEMP
    QUOTA UNLIMITED ON USERS;

GRANT CREATE SESSION, CREATE TABLE, CREATE VIEW,
      CREATE SEQUENCE, CREATE PROCEDURE TO gbs_user;
```

### 3. Executar os scripts (como `gbs_user`)

Na ordem:

1. `Database/01_DDL_Tabelas.sql`
2. `Database/02_DDL_Sequences.sql`
3. `Database/03_PACK_EQUIPAMENTO_SPEC.sql`
4. `Database/04_PACK_EQUIPAMENTO_BODY.sql`
5. `Database/05_PACK_UPGRADE.sql`
6. `Database/06_PACK_REMESSA.sql`

Validação:

```sql
SELECT object_type, COUNT(*)
FROM user_objects
WHERE object_type IN ('TABLE','SEQUENCE','PACKAGE','PACKAGE BODY')
AND status = 'VALID'
GROUP BY object_type;
```

Esperado: 6 TABLE, 6 SEQUENCE, 3 PACKAGE, 3 PACKAGE BODY — **todos VALID**.

### 4. Configurar a connection string

Edite `GBS_Inventory/App.config`:

```xml
<connectionStrings>
  <add name="OracleDB"
       connectionString="User Id=gbs_user;Password=gbs_pass;Data Source=localhost:1521/XE;"
       providerName="Oracle.ManagedDataAccess.Client"/>
</connectionStrings>
```

> **Importante para Oracle XE 11g**: use `Data Source=localhost:1521/XE`. Não use apenas `Data Source=XE` porque o `Oracle.ManagedDataAccess` não lê `tnsnames.ora` por padrão.

### 5. Abrir no Visual Studio

1. Abra `GBS_Inventory.sln`
2. Botão direito na solution → **Restore NuGet Packages**
3. **Build → Build Solution** (`Ctrl+Shift+B`)
4. **F5** para rodar

Documentação detalhada em [docs/INSTALL.md](docs/INSTALL.md).

---

## 🏗️ Arquitetura

```
┌─────────────────────────────────────────────────────────┐
│ VIEW (WinForms)                                         │
│ frmPrincipal · frmScanner · frmImportacao · frmUpgrade │
│ frmRemessa                                              │
└───────────────────────┬─────────────────────────────────┘
                        │ (eventos)
┌───────────────────────▼─────────────────────────────────┐
│ CONTROLLER (lógica + controle de transações)            │
│ EquipamentoController · UpgradeController               │
│ RemessaController · ImportacaoController                │
└───────────────────────┬─────────────────────────────────┘
                        │ (instancia)
┌───────────────────────▼─────────────────────────────────┐
│ MODEL                 │ CLASSES DE ACESSO A DADOS       │
│ Equipamento · Upgrade │ clsGravacao* + clsLeitura*      │
│ Remessa · ItemRemessa │ clsImportacaoExcel              │
└──────────────────────┬┴─────────────────────────────────┘
                       │ (OracleHelper)
┌──────────────────────▼──────────────────────────────────┐
│ ORACLE DATABASE                                         │
│ PACK_EQUIPAMENTO · PACK_UPGRADE · PACK_REMESSA          │
│ 6 tabelas · 6 sequences                                 │
└─────────────────────────────────────────────────────────┘
```

Detalhes em [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).

---

## 📁 Estrutura de diretórios

```
GBS_Inventory/
├── GBS_Inventory.sln
├── README.md                          ← este arquivo
├── LICENSE
├── CHANGELOG.md
├── CONTRIBUTING.md
├── .gitignore
├── Database/                          ← scripts de instalação Oracle
│   ├── 01_DDL_Tabelas.sql
│   ├── 02_DDL_Sequences.sql
│   ├── 03_PACK_EQUIPAMENTO_SPEC.sql
│   ├── 04_PACK_EQUIPAMENTO_BODY.sql
│   ├── 05_PACK_UPGRADE.sql
│   └── 06_PACK_REMESSA.sql
├── docs/
│   ├── INSTALL.md                     ← guia de instalação passo a passo
│   ├── ARCHITECTURE.md                ← arquitetura MVC explicada
│   ├── DATABASE.md                    ← dicionário de dados completo
│   ├── PACKAGES.md                    ← referência de procedures
│   ├── UI.md                          ← guia de uso das telas
│   ├── TESTING.md                     ← roteiro de testes end-to-end
│   └── TROUBLESHOOTING.md             ← erros comuns e soluções
└── GBS_Inventory/
    ├── GBS_Inventory.vbproj
    ├── App.config                     ← connection string
    ├── packages.config
    ├── My Project/
    ├── Oracle/
    │   └── OracleHelper.vb
    ├── Models/                        ← M do MVC
    │   ├── Equipamento.vb
    │   ├── Upgrade.vb
    │   └── Remessa.vb
    ├── classes/                       ← acesso a dados
    │   ├── clsGravacaoEquipamento.vb
    │   ├── clsGravacaoUpgrade.vb
    │   ├── clsGravacaoRemessa.vb
    │   ├── clsLeituraEquipamento.vb
    │   ├── clsLeituraUpgrade.vb
    │   ├── clsLeituraRemessa.vb
    │   └── clsImportacaoExcel.vb
    ├── Controllers/                   ← C do MVC
    │   ├── EquipamentoController.vb
    │   ├── UpgradeController.vb
    │   ├── RemessaController.vb
    │   └── ImportacaoController.vb
    ├── Utils/
    │   └── TemaEscuro.vb
    └── Views/                         ← V do MVC
        ├── frmPrincipal.vb + Designer
        ├── frmScanner.vb + Designer
        ├── frmImportacao.vb + Designer
        ├── frmUpgrade.vb + Designer
        └── frmRemessa.vb + Designer
```

---

## 🗄️ Modelo de dados

| Tabela | Registros esperados | Chave natural |
|--------|---------------------|---------------|
| `TBL_EQUIPAMENTO` | ~1000/ano | `INTERNAL_UID` (UK) + `SERIAL_NUMBER` (UK) |
| `TBL_EQUIPAMENTO_UPGRADE` | ~500/ano | FK → equipamento |
| `TBL_REMESSA` | ~200/ano | `REMESSA_REF` (UK) |
| `TBL_REMESSA_ITEM` | 1-N por remessa | FK → remessa + equipamento |
| `TBL_IMPORTACAO_LOG` | ~50/ano | Auditoria |
| `TBL_WHATSAPP_LOG` | (reservada) | — |

Dicionário completo em [docs/DATABASE.md](docs/DATABASE.md).

---

## 🎨 Paleta de cores (tema escuro)

| Elemento | Cor | Hex |
|----------|-----|-----|
| Fundo | Quase preto | `#080c10` |
| Surface | Cinza-azul escuro | `#131920` |
| Surface inputs | Mais claro | `#1c242e` |
| Borda sutil | Cinza-azul | `#26303c` |
| Accent (destaque) | Ciano/teal | `#00e5a0` |
| Texto | Quase branco | `#e8edf2` |
| Texto mutado | Cinza | `#8c9baa` |
| Verde status OK | Verde | `#2ecc71` |
| Vermelho erro | Vermelho | `#e74c3c` |

---

## 🧪 Testes

Roteiro completo de testes end-to-end em [docs/TESTING.md](docs/TESTING.md), incluindo:

1. Validação PL/SQL (insert, upgrade, remessa, cascata DELIVERED → SOLD)
2. Testes funcionais de cada tela
3. Teste de importação com a planilha real da GBS
4. Checklist de validação antes do deploy

---

## 🛠️ Troubleshooting

Erros comuns e soluções em [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md):

- `ORA-12154`, `ORA-01017`, `ORA-00942`, `PLS-00357`, `PLS-00201`
- `Microsoft.ACE.OLEDB.12.0 provider is not registered`
- Bug do Visual Studio Designer com `Me.Width`
- `Oracle.ManagedDataAccess` não restaurado

---

## 🤝 Contribuindo

Veja [CONTRIBUTING.md](CONTRIBUTING.md).

Padrões adotados:

- Classes de acesso a dados seguem estrutura em `#Region`: Atributos / Construtor / Transação / Métodos Transacionais
- Procedures Oracle com prefixo `PROC_`, functions com `FUNC_`
- Parâmetros sempre com prefixo `V_` (valor) / `P_` (ponteiro, não usado)
- Nomes de tabela em português (`TBL_`), prefixo em UPPER_CASE
- Commits em português, imperativos (`adiciona tela X`, `corrige bug Y`)

---

## 📄 Licença

Este projeto é propriedade da **Global Business Solution**. Uso interno.

Veja [LICENSE](LICENSE).

---

## 👤 Autor

**Hernani** — desenvolvedor principal.

Para dúvidas técnicas, abra uma [issue](../../issues).

---

## 📝 Changelog

Veja [CHANGELOG.md](CHANGELOG.md) para histórico de versões.
