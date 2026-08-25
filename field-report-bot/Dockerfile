FROM node:20-slim

# System Chromium + the shared libraries whatsapp-web.js/puppeteer need to run it,
# instead of letting puppeteer download its own copy (smaller image, faster build).
RUN apt-get update && apt-get install -y --no-install-recommends \
    chromium \
    fonts-liberation \
    fonts-noto \
    ca-certificates \
    dumb-init \
    && rm -rf /var/lib/apt/lists/*

ENV PUPPETEER_SKIP_CHROMIUM_DOWNLOAD=true \
    PUPPETEER_EXECUTABLE_PATH=/usr/bin/chromium \
    NODE_ENV=production

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci --omit=dev

COPY . .

# WhatsApp login session and generated reports must survive restarts/redeploys.
VOLUME ["/app/.wwebjs_auth", "/app/data"]

EXPOSE 3211

# dumb-init reaps the headless Chromium child process cleanly on stop/restart.
ENTRYPOINT ["dumb-init", "--"]
CMD ["node", "src/index.js"]
