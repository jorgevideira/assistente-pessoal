# Instalação do zero

Tempo: ~30 minutos. Você precisa de uma conta Google e do Telegram.

## 1. Criar o bot no Telegram

1. No Telegram, procure **@BotFather** e mande `/newbot`.
2. Escolha um nome e um usuário (precisa terminar em `bot`).
3. Guarde o **token** que ele devolve (formato `123456789:ABC...`).

## 2. Pegar a chave do Gemini

1. Acesse **https://aistudio.google.com** com sua conta Google.
2. **Get API key → Create API key**.
3. Guarde a chave.

## 3. Criar o projeto no Apps Script

1. Acesse **https://script.google.com** → **Novo projeto** e dê um nome (ex.: "Assistente").
2. Apague o código padrão e cole todo o conteúdo de [`src/Codigo.gs`](../src/Codigo.gs).
3. No topo, preencha `TELEGRAM_TOKEN` e `GEMINI_API_KEY`.
4. Em ⚙️ **Configurações do projeto**, defina o fuso horário como **(GMT-03:00) São Paulo**.
   *Opcional:* marque "Mostrar arquivo de manifesto appsscript.json" e cole o conteúdo de [`src/appsscript.json`](../src/appsscript.json).
5. Salve 💾.

## 4. Executar `configurar`

1. Na lista de funções (ao lado de "Depuração"), escolha **`configurar`** → **Executar**.
2. Autorize o acesso. Como o script é seu, vai aparecer "app não verificado": **Avançado → Acessar (não seguro) → Permitir**.
3. No registro aparece o link da planilha **"Assistente - Gastos"**, criada no seu Drive.

## 5. Publicar como App da Web

1. **Implantar → Nova implantação → ⚙️ → App da Web**.
2. **Executar como:** Eu. **Quem pode acessar:** **Qualquer pessoa** (não "Qualquer pessoa com Conta do Google").
3. **Implantar** e copie a URL (termina em `/exec`).

## 6. Conectar ao Telegram

1. Cole a URL em `WEB_APP_URL`, no topo do código, e salve.
2. Execute **`registrarWebhook`**. O registro deve mostrar `"ok": true`.

## 7. Ativar resumo diário e lembretes

Execute **`criarGatilhos`**.

## 8. Testar

1. Abra o seu bot no Telegram e mande `/start`. Você vira o **dono**.
2. Experimente: "gastei 45 no mercado", "marca dentista amanhã às 10h", um áudio, a foto de um cupom.

Resposta normal: **3 a 10 segundos**. Se não vier, veja [solução de problemas](solucao-de-problemas.md).

## 9. (Opcional) Outlook do trabalho

Veja [outlook.md](outlook.md).

---

## Atualizando o código depois

Sempre que alterar o código:

1. Salve 💾.
2. **Implantar → Gerenciar implantações → ✏️ → Versão: "Nova versão" → Implantar.**

> ⚠️ Sem esse passo o bot continua rodando a versão antiga — salvar não basta. A URL continua a mesma.

Ao colar uma versão nova inteira, recoloque seus valores no topo (`TELEGRAM_TOKEN`, `GEMINI_API_KEY`, `WEB_APP_URL`, `OUTLOOK_ICS_URL`). Não é preciso rodar `configurar` nem `registrarWebhook` de novo: os segredos e o dono ficam salvos nas Propriedades do script.

## Funções que você executa pelo editor

| Função | Quando usar |
|---|---|
| `configurar` | Primeira instalação (salva segredos, cria a planilha) |
| `registrarWebhook` | Depois de implantar, ou para destravar mensagens presas |
| `criarGatilhos` | Liga resumo diário e lembretes |
| `configurarOutlook` | Liga a sincronização com o Outlook |
| `desativarOutlook` | Desliga a sincronização |
| `ativarModoPolling` | Plano B se o webhook não funcionar (resposta em até 1 min) |
| `diagnosticar` | Mostra o estado do webhook, do dono e do Outlook |
| `testarGemini` | Testa quais modelos de IA respondem com a sua chave |
| `listarAgendas` | Mostra as agendas visíveis e quantos eventos cada uma tem |
