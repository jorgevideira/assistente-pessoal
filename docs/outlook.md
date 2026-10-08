# Agenda do Outlook (trabalho)

O bot espelha a agenda do Outlook / Microsoft 365 numa agenda separada do Google chamada **"Trabalho (Outlook)"**, em **mão única**:

| No Outlook | No Google (até ~15 min depois) |
|---|---|
| Criou um compromisso | É criado |
| Alterou horário, título ou local | É atualizado |
| Removeu ou cancelou | É apagado |
| Moveu/excluiu só uma ocorrência de uma série | Só aquela ocorrência muda |

Nada é enviado de volta ao Outlook. Pelo bot, os eventos do trabalho são **somente consulta** e aparecem com 💼.

## 1. Publicar a agenda no Outlook

1. Abra o **Outlook na web** (outlook.office.com).
2. ⚙️ **Configurações → Calendário → Calendários compartilhados**.
3. Em **Publicar um calendário**, escolha o calendário e o nível de detalhe:
   - **"Pode ver títulos, locais e mais detalhes"** — o bot sabe do que é cada reunião;
   - **"Pode ver apenas quando estou ocupado"** — só horários (mais privado).
4. **Publicar** e copie o link **ICS**.

Se a opção não aparecer, o administrador da empresa bloqueou a publicação; só o TI libera.

> 🔒 Quem tiver o link vê sua agenda. Não publique o link em lugar nenhum nem faça commit dele.
> Confira a política da empresa: no plano gratuito do Gemini, o conteúdo pode ser usado pelo Google para treinar modelos. Publicar com "apenas quando estou ocupado" evita expor assuntos.

## 2. Ligar a sincronização

1. Cole o link em `OUTLOOK_ICS_URL`, no topo do código, e salve.
2. Publique uma **Nova versão** (Implantar → Gerenciar implantações → ✏️).
3. Execute **`configurarOutlook`** e autorize (ele precisa criar uma agenda).

O registro mostra quantos compromissos foram trazidos. A partir daí roda sozinho a cada 15 minutos. Para forçar na hora, mande `/sincronizar` no Telegram.

## Como funciona por dentro

1. Baixa o ICS (`UrlFetchApp`).
2. Interpreta os eventos, incluindo:
   - fusos do Windows (`E. South America Standard Time`, `Eastern Standard Time`...) e IANA, convertidos para Brasília;
   - repetições (`RRULE`): diária, semanal, mensal (dia fixo, "2ª terça", "último dia útil"), anual, com `INTERVAL`, `COUNT` e `UNTIL`;
   - exclusões (`EXDATE`), ocorrências alteradas (`RECURRENCE-ID`) e eventos cancelados;
   - eventos de dia inteiro.
3. Compara com o que já está na agenda "Trabalho (Outlook)" usando duas *tags* em cada evento: uma chave (evento + horário) e um *hash* do conteúdo.
4. Cria o que falta, apaga o que sumiu e recria o que mudou.

Janela sincronizada: do dia anterior até **60 dias à frente** (`SYNC_DIAS`).

## Cuidados

- **Não edite** eventos dentro da agenda "Trabalho (Outlook)" no Google: a próxima sincronização desfaz.
- Eventos criados à mão nessa agenda (sem as tags do script) não são tocados.
- Tempo até aparecer: os 15 min do script **mais** o tempo que a Microsoft leva para atualizar o link publicado (normalmente minutos, às vezes mais).
- Se o link parar de funcionar, o bot avisa no Telegram (no máximo uma vez a cada 6 h) e não apaga nada.

## Desligar

Execute **`desativarOutlook`**. A agenda "Trabalho (Outlook)" continua no Google; apague-a manualmente se quiser.
