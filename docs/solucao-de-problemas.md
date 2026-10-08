# Solução de problemas

Problemas que já apareceram durante a montagem do bot e como foram resolvidos.

## O bot não responde nada

Execute **`diagnosticar`** e olhe o resultado de `getWebhookInfo`:

| O que aparece | Causa | Solução |
|---|---|---|
| `"url": ""` e `pending_update_count` > 0 | Webhook nunca registrado | Confira `WEB_APP_URL` e execute `registrarWebhook` |
| `last_error_message` com `302` | O Google respondeu com redirecionamento e o Telegram tratou como falha | Veja a próxima seção |
| `last_error_message` com `401`/`403` | Implantação não está pública | Implantação com acesso **"Qualquer pessoa"** e nova versão |

Também vale abrir **Execuções** (ícone ☰▸ na barra lateral do Apps Script) para ver se o `doPost` está rodando e se deu erro.

## A mesma mensagem é processada várias vezes / as seguintes nunca chegam

Sintoma: em **Execuções** aparecem vários `doPost` seguidos, cada um em menos de 1 s, com intervalos crescentes (1 s, 3 s, 5 s, 10 s…), e as mensagens novas não são respondidas.

Causa: o Telegram reenvia a mesma mensagem porque recebeu um redirecionamento (302) do Google, e fica preso nela.

Solução (já aplicada no código): responder com `HtmlService` em vez de `ContentService`:

```js
function ok_() { return HtmlService.createHtmlOutput('ok'); }
```

Depois: publicar **Nova versão** e executar `registrarWebhook` (que também limpa a fila travada).

Plano B, se ainda falhar: `ativarModoPolling` (o script busca as mensagens a cada minuto; resposta em até 1 min).

## "⚠️ O limite gratuito do Gemini foi atingido" logo na primeira mensagem

Não era limite: o modelo estava sobrecarregado (erro 503) ou não estava disponível na camada gratuita.

1. Execute **`testarGemini`** — ele testa vários modelos com a sua chave.
2. Coloque em `MODELO_GEMINI` o que respondeu `✅ FUNCIONA` mais rápido.

O código atual já tenta os `MODELOS_RESERVA` sozinho e mostra o erro real na mensagem.

Referência (08/10/2026): `gemini-3.1-flash-lite` respondeu em < 1 s; `gemini-3.5-flash` funcionou mas levou ~18 s; os modelos 2.5 foram desativados para usuários novos.

## O bot não mostra eventos de uma agenda

O bot só lê: a agenda principal, a "Trabalho (Outlook)" e as listadas em `AGENDAS_EXTRAS`.

1. Execute **`listarAgendas`** — mostra o nome exato de cada agenda e quantos eventos tem nos próximos 30 dias.
2. Copie o nome exato para `AGENDAS_EXTRAS`.
3. Publique **Nova versão**.
4. No Telegram, mande `/limpar` antes de perguntar de novo (a memória da conversa pode repetir a resposta antiga).

## Mudei o código e nada mudou

Faltou publicar: **Implantar → Gerenciar implantações → ✏️ → Versão: "Nova versão" → Implantar.** Salvar não atualiza o bot.

## Reunião do Outlook não apareceu

- Espere até ~30 min (15 min do script + atualização do link pela Microsoft) ou mande `/sincronizar`.
- Execute `diagnosticar` para confirmar que o Outlook está configurado.
- Abra o link ICS no navegador: se não baixar um arquivo começando com `BEGIN:VCALENDAR`, a publicação foi desativada no Outlook.

## "Este app não foi verificado" ao autorizar

Normal: o script é seu e não passou pela verificação do Google. **Avançado → Acessar (não seguro) → Permitir.**
