# Plano de Teste A/B - Clientes e Invoice (GBS_Inventory)

## Objetivo
Validar duas variantes visuais e funcionais nas telas `Clientes` e `Invoice`, garantindo:
- consistencia com o padrao Helius;
- ausencia de erro ao abrir/carregar dados;
- estabilidade das operacoes principais (CRUD cliente, gerar invoice, enviar email).

## Variantes
- **Variante A (baseline):** tema original (`UiVariantClientesInvoice = A`)
- **Variante B (novo padrao):** tema Helius (`UiVariantClientesInvoice = B`)

Arquivo de controle:
- `C:\Users\herna\Desktop\GBS\Cloude\GBS Inventory\GBS_Inventory\GBS_Inventory\App.config`
- chave: `UiVariantClientesInvoice`

## Escopo funcional do teste
1. Abertura da tela `Clientes`.
2. Abertura da tela `Invoice`.
3. Pesquisa/listagem de clientes.
4. Inclusao/alteracao/exclusao (ou inativacao) de cliente.
5. Carregamento de itens por remessa na invoice.
6. Geracao de invoice (registro + Word + PDF).
7. Envio de email com anexos (PDF/Word/logo).

## Criticidade (go/no-go)
- **Bloqueante:** erro nao tratado ao abrir `Clientes` ou `Invoice`.
- **Alta:** falha em salvar cliente ou gerar invoice.
- **Media:** inconsistencia visual grave (controles ilegiveis, colisoes).

## Massa de dados recomendada
- 1 cliente ativo existente.
- 1 cliente novo para criar durante o teste.
- 1 remessa ativa com pelo menos 3 itens.
- 1 email de homologacao para disparo (SMTP valido).

## Roteiro de execucao A/B
1. Configurar `UiVariantClientesInvoice = A`.
2. Reiniciar sistema e executar todos os cenarios abaixo.
3. Registrar resultados (tempo, erro, evidencias).
4. Configurar `UiVariantClientesInvoice = B`.
5. Repetir os mesmos cenarios.
6. Comparar resultados A vs B.

## Cenarios de teste
### C01 - Abrir tela de Clientes
- Passos: abrir `Clientes` pelo menu principal.
- Esperado: tela abre sem excecao Oracle/SQL.
- Evidencia: screenshot + horario.

### C02 - Listar/pesquisar clientes
- Passos: pesquisar por nome, email e documento.
- Esperado: grid responde sem travar e sem erro.
- Evidencia: total retornado e tempo de resposta.

### C03 - Criar cliente
- Passos: preencher campos obrigatorios e salvar.
- Esperado: ID gerado + registro visivel no grid.

### C04 - Alterar cliente
- Passos: editar cliente existente e salvar.
- Esperado: dados persistidos e refletidos na recarga.

### C05 - Excluir/Inativar cliente
- Passos: excluir cliente sem invoice e cliente com invoice.
- Esperado:
- sem invoice: remocao;
- com invoice: inativacao (`ATIVO = 'N'`).

### C06 - Abrir tela de Invoice
- Passos: abrir `Invoice` pelo menu principal.
- Esperado: sem erro ao carregar combo de cliente/remessa.

### C07 - Carregar itens da remessa
- Passos: selecionar remessa e clicar `Carregar Itens`.
- Esperado: grid preenchido; subtotal e total atualizados.

### C08 - Editar item e recalcular
- Passos: alterar `QTY` e `UNIT_PRICE_USD`.
- Esperado: `LINE_TOTAL_USD` atualizado e total recalculado.

### C09 - Gerar invoice (Word/PDF)
- Passos: preencher dados, gerar invoice.
- Esperado:
- registro no Oracle criado;
- arquivo `.doc` e `.pdf` criados no output;
- numero invoice exibido.

### C10 - Enviar invoice por email
- Passos: clicar `Enviar`.
- Esperado: email enviado com anexos (pdf, doc e logo quando existir) e status `SENT`.

## Indicadores para comparacao A vs B
- Taxa de sucesso por cenario (meta: 100% nos bloqueantes).
- Quantidade de excecoes por sessao (meta: 0).
- Tempo medio para completar C01-C10.
- Feedback visual do usuario (legibilidade, consistencia, conforto).

## Criterio de aprovacao da variante B
- Nenhum erro bloqueante.
- Mesmo resultado funcional da variante A (ou melhor).
- Melhor percepcao visual/consistencia pelo usuario.

## Plano de rollback
Se houver falha relevante na variante B:
1. voltar `UiVariantClientesInvoice` para `A`;
2. reiniciar aplicacao;
3. manter operacao com baseline;
4. registrar causa raiz e abrir correcao incremental.
