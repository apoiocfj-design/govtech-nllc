# Guia Definitivo: Publicação no Render.com (100% Gratuito e Sem Cartão)
**Sistema GovTech NLLC (Lei 14.133/2021 & IN 65/2021)**

O **[Render.com](https://render.com/)** é a melhor opção para colocar o sistema no ar para qualquer pessoa no Brasil:
- **100% Gratuito** (plano Free para Web Services).
- **Sem pedir cartão de crédito**.
- Fornece link público oficial com cadeado de segurança HTTPS (ex: `https://govtech-nllc.onrender.com`).
- Deixamos tudo pré-configurado com o arquivo `render.yaml` e `Dockerfile`.

---

## O que já fizemos para você automaticamente:
1. **Instalamos o Git** na sua máquina.
2. **Inicializamos o repositório Git local** e fizemos o primeiro *commit* de todos os arquivos.
3. **Comprimimos o banco de dados oficial CNBS** de 208 MB para 38 MB (`cnbs_catalog.db.gz`), contornando o limite de 100 MB do GitHub sem perder nenhum dado.
4. **Configuramos o `render.yaml`** (Blueprint) e o `Dockerfile` com porta dinâmica, verificação de integridade e descompressão instantânea.
5. **Criamos o script executável:** [`PUBLICAR_NO_RENDER.bat`](file:///C:/Users/CFJ-SECOM/.gemini/antigravity/scratch/govtech-nllc/PUBLICAR_NO_RENDER.bat).

---

## O único passo que você precisa fazer (Leva 2 minutos):

### Opção Recomendada (Automática com o script):
Dê um duplo clique no arquivo:  
👉 **[`PUBLICAR_NO_RENDER.bat`](file:///C:/Users/CFJ-SECOM/.gemini/antigravity/scratch/govtech-nllc/PUBLICAR_NO_RENDER.bat)**

Ele fará o seguinte:
1. Abrirá a página do GitHub para criar um repositório com o nome `govtech-nllc`. Você só clica no botão verde **"Create repository"**.
2. Você cola o link HTTPS do seu repositório na tela preta (ex: `https://github.com/seu-usuario/govtech-nllc.git`) e aperta **ENTER**.
3. O script sobe todo o código automaticamente para o GitHub.
4. O script abre a página do **Render.com**. Você faz login com seu e-mail Google (`ricardoluizrodriguesgomes@gmail.com`), seleciona o repositório e clica em **"Create Web Service"**.

---

## Resultado Final:
Em cerca de 2 a 3 minutos, o Render construirá a imagem e gerará o link público:
👉 **`https://govtech-nllc.onrender.com`**

Qualquer colega, auditor, pregoeiro ou comprador em qualquer lugar do Brasil poderá acessar por esse link no celular ou no computador 24h por dia!
