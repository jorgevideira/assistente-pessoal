# Configurações

Tudo fica no topo de `src/Codigo.gs`. Depois de mudar qualquer coisa: salve e publique uma **Nova versão**.

## Valores obrigatórios

| Constante | O que é |
|---|---|
| `TELEGRAM_TOKEN` | Token do bot, dado pelo @BotFather |
| `GEMINI_API_KEY` | Chave da API do Gemini (AI Studio) |
| `WEB_APP_URL` | URL da implantação do App da Web (termina em `/exec`) |
| `OUTLOOK_ICS_URL` | *(opcional)* Link ICS publicado pelo Outlook |

## Ajustes

| Constante | Padrão | Para que serve |
|---|---|---|
| `MODELO_GEMINI` | `'gemini-3.1-flash-lite'` | Modelo de IA principal. Rode `testarGemini` para ver quais funcionam. |
| `MODELOS_RESERVA` | `['gemini-flash-lite-latest', 'gemini-3.5-flash']` | Usados automaticamente se o principal estiver fora do ar. |
| `HORA_RESUMO` | `7` | Hora do resumo diário. Depois de mudar, rode `criarGatilhos`. |
| `LEMBRETES_MINUTOS` | `[15]` | Avisos antes de cada compromisso. Ex.: `[60, 15]` = 1 h antes e 15 min antes. Máximo 360. |
| `AGENDAS_EXTRAS` | `['Videira Sertãozinho']` | Outras agendas do Google para consultar (somente leitura), pelo nome exato. Rode `listarAgendas` para ver os nomes. |
| `NOME_AGENDA_TRABALHO` | `'Trabalho (Outlook)'` | Nome da agenda criada para o Outlook. |
| `SYNC_DIAS` | `60` | Quantos dias à frente sincronizar do Outlook. |
| `CATEGORIAS` | Alimentação, Mercado, Transporte… | Categorias de gasto que a IA pode usar. |

## Como os lembretes funcionam

O gatilho `verificarLembretes` roda a cada 5 minutos e olha todas as agendas (pessoal, extras e trabalho). Por isso um aviso de 15 min chega entre 10 e 15 minutos antes. Eventos de dia inteiro não geram lembrete.

## Onde ficam os dados

| Dado | Onde |
|---|---|
| Gastos | Planilha "Assistente - Gastos" no seu Google Drive (pode editar à mão) |
| Compromissos pessoais | Sua agenda principal do Google |
| Compromissos do trabalho | Agenda "Trabalho (Outlook)" no Google |
| Segredos, dono, IDs | Propriedades do script (⚙️ Configurações do projeto → Propriedades do script) |
| Memória da conversa | Cache do script, por 30 minutos (`/limpar` apaga) |
