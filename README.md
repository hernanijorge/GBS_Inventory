# GBS Inventory Manager

Sistema WinForms em VB.NET para controle de equipamentos, upgrades, remessas, clientes e invoices, com persistencia em Oracle.

## Stack

- Visual Studio 2022
- .NET Framework 4.8
- VB.NET
- Oracle Database XE ou compativel
- Oracle.ManagedDataAccess 21.13.0

## Estrutura

```text
GBS_Inventory/
|-- Database/
|   |-- 01_DDL_Tabelas.sql
|   |-- 02_DDL_Sequences.sql
|   |-- 03_PACK_EQUIPAMENTO_SPEC.sql
|   |-- 04_PACK_EQUIPAMENTO_BODY.sql
|   |-- 05_PACK_UPGRADE.sql
|   |-- 06_PACK_REMESSA.sql
|   |-- 07_DDL_CADASTROS_INVOICE.sql
|   |-- 08_PACK_CLIENTE.sql
|   `-- 09_PACK_INVOICE.sql
|-- GBS_Inventory/
|   |-- Controllers/
|   |-- Models/
|   |-- Oracle/
|   |-- Utils/
|   |-- Views/
|   |-- classes/
|   |-- App.config
|   `-- GBS_Inventory.vbproj
`-- GBS_Inventory.sln
```

## Modulos

- Dashboard
- Registration
- Inventory
- Shipments
- Upgrades
- Import
- Clientes
- Invoice

## Banco de dados

Os scripts da pasta `Database` refletem a estrutura atual usada pela aplicacao.

Ordem de execucao recomendada:

1. `01_DDL_Tabelas.sql`
2. `02_DDL_Sequences.sql`
3. `03_PACK_EQUIPAMENTO_SPEC.sql`
4. `04_PACK_EQUIPAMENTO_BODY.sql`
5. `05_PACK_UPGRADE.sql`
6. `06_PACK_REMESSA.sql`
7. `07_DDL_CADASTROS_INVOICE.sql`
8. `08_PACK_CLIENTE.sql`
9. `09_PACK_INVOICE.sql`

Packages esperados pela aplicacao:

- `PACK_EQUIPAMENTO`
- `PACK_UPGRADE`
- `PACK_REMESSA`
- `PACK_CLIENTE`
- `PACK_INVOICE`

## Connection string

Arquivo:

- `GBS_Inventory/App.config`

Exemplo de configuracao local:

```xml
<connectionStrings>
  <add name="OracleDB"
       connectionString="User Id=gbs_owner;Password=gbs_pass;Data Source=localhost:1521/XE;"
       providerName="Oracle.ManagedDataAccess.Client"/>
</connectionStrings>
```

## Build

1. Abrir `GBS_Inventory.sln` no Visual Studio 2022
2. Restaurar os pacotes NuGet
3. Compilar a solution
4. Executar a aplicacao

Atalho de compilacao:

```text
Ctrl + Shift + B
```

Atalho para executar:

```text
F5
```

## Observacoes

- A aplicacao usa `BindByName = True` no acesso Oracle. Os nomes dos parametros no VB.NET precisam coincidir com os nomes publicados nos packages.
- A pasta `Database` foi consolidada para manter apenas os scripts principais. Scripts de correcao intermediarios foram removidos.
- O projeto usa `Oracle.ManagedDataAccess`, sem dependencia de Oracle Client para a biblioteca .NET.

## Arquivos importantes

- `GBS_Inventory/Oracle/OracleHelper.vb`
- `GBS_Inventory/classes/clsGravacaoEquipamento.vb`
- `GBS_Inventory/classes/clsGravacaoUpgrade.vb`
- `GBS_Inventory/classes/clsGravacaoRemessa.vb`
- `GBS_Inventory/classes/clsLeituraEquipamento.vb`
- `GBS_Inventory/classes/clsLeituraUpgrade.vb`
- `GBS_Inventory/classes/clsLeituraRemessa.vb`

## Status

Repositorio com build local validado no Visual Studio e integracao ajustada para as assinaturas atuais dos packages Oracle.
