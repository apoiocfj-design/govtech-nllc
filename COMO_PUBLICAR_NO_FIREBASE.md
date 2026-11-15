# Guia de Publicação no Firebase & Google Cloud
**Sistema GovTech NLLC (Lei 14.133/2021 & IN 65/2021)**

Este guia explica como hospedar o aplicativo na sua conta Google (**`ricardoluizrodriguesgomes@gmail.com`**) para obter um link oficial de alta velocidade (ex: `https://govtech-nllc.web.app`) acessível por qualquer pessoa no Brasil.

---

## 1. A sua conta gratuita atende ou precisa contratar algum serviço?

### Resposta Direta:
**A sua conta atende 100% de forma GRATUITA (R$ 0,00 de custo real). Não há necessidade de contratar plano pago empresarial.**

No entanto, há uma regra de funcionamento do Google que você precisa saber:

1. **O que é 100% Gratuito:**
   - O Google concede um **Nível Gratuito Permanente (Always Free Tier)** todos os meses:
     - **Firebase Hosting:** 10 GB de armazenamento e transferência gratuita.
     - **Google Cloud Run:** Até **2.000.000 (2 milhões) de requisições por mês grátis**, 360.000 GB-segundos de memória e 180.000 vCPU-segundos grátis.
   - Como o nosso sistema para a capitania/órgão realiza centenas ou alguns milhares de consultas por mês, o seu consumo fica **muito abaixo** da cota de isenção, resultando em fatura de **R$ 0,00**.

2. **A exigência do Plano "Blaze" no Firebase:**
   - O plano padrão inicial do Firebase (*Spark*) só aceita arquivos estáticos puros (HTML/JS).
   - Como o nosso aplicativo tem um **backend Python** que faz requisições externas para as APIs do Governo (`Compras.gov.br` e `PNCP`), o Firebase exige a ativação do plano **Blaze (Pay-as-you-go)** no projeto.
   - **O que significa o Plano Blaze?** Você cadastra um cartão de crédito no Google Cloud para comprovação de identidade e proteção contra bots/mineradores, **mas nada é cobrado** desde que você fique dentro do limite gratuito de 2 milhões de requisições.
   - **Trava de Segurança:** Você pode configurar no painel do Google Cloud um **Orçamento/Alerta de Gastos com teto de R$ 0,01** para garantir que o Google te alerte imediatamente se algo tentar ser cobrado.

---

## 2. Passo a Passo para Publicar (10 a 15 minutos)

### Passo 1: Criar o Projeto no Console do Firebase
1. Acesse: **[https://console.firebase.google.com/](https://console.firebase.google.com/)** logado com sua conta `ricardoluizrodriguesgomes@gmail.com`.
2. Clique em **"Adicionar projeto"** e dê um nome (ex: `govtech-compras-marinha`).
3. Pode desativar o Google Analytics se preferir simplicidade e clique em **Criar projeto**.
4. No menu inferior esquerdo, mude o plano de *Spark* para **Blaze** (cadastrando sua conta de faturamento do Google Cloud).

---

### Passo 2: Instalar as ferramentas no seu computador (se ainda não tiver)
Abra o terminal (PowerShell) e instale:
```powershell
# 1. Instalar o Google Cloud SDK (gcloud)
winget install Google.CloudSDK

# 2. Instalar o Firebase Tools (necessita Node.js)
npm install -g firebase-tools
```

---

### Passo 3: Fazer Login e Publicar o Backend (Cloud Run)
No terminal, dentro da pasta do projeto (`C:\Users\CFJ-SECOM\.gemini\antigravity\scratch\govtech-nllc`):

```powershell
# 1. Conectar na sua conta Google
gcloud auth login ricardoluizrodriguesgomes@gmail.com

# 2. Definir o projeto criado
gcloud config set project NOME_DO_SEU_PROJETO

# 3. Publicar o contêiner do nosso app na região de São Paulo (menor latência no Brasil)
gcloud run deploy govtech-nllc --source . --region southamerica-east1 --allow-unauthenticated
```
*O Google irá construir a imagem Docker automaticamente na nuvem (Cloud Build) e devolverá a URL do serviço.*

---

### Passo 4: Conectar o Firebase Hosting ao Cloud Run
No mesmo terminal:
```powershell
# 1. Login no Firebase
firebase login

# 2. Vincular o projeto
firebase use NOME_DO_SEU_PROJETO

# 3. Publicar o Hosting
firebase deploy --only hosting
```

Pronto! O Firebase exibirá o seu link oficial:
👉 **`https://NOME_DO_SEU_PROJETO.web.app`**  
*(Qualquer pessoa que acessar esse link no celular ou computador usará o sistema completo 24h por dia).*

---

## 3. Alternativa 100% Gratuita Sem Cartão de Crédito (Render.com)

Se você **não quiser cadastrar cartão de crédito** no Google Cloud/Firebase neste momento, há uma alternativa excelente e igualmente profissional:

- **[Render.com](https://render.com/)**:
  - Aceita contêineres Docker e Python gratuitamente sem pedir cartão de crédito.
  - Fornece um link seguro com HTTPS automático (ex: `https://govtech-nllc.onrender.com`).
  - O mesmo `Dockerfile` que já deixamos pronto na pasta funciona diretamente no Render!
