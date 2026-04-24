FROM node:22-slim

# Install LibreOffice (for DOCX → PDF) and poppler-utils (for PDF → PNG)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    poppler-utils \
    fonts-noto-cjk \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package*.json ./
RUN npm ci --omit=dev

COPY src/ ./src/

# GCP credentials are mounted at runtime via GOOGLE_APPLICATION_CREDENTIALS
ENV NODE_ENV=production
ENV PORT=3456
ENV HOST=0.0.0.0

EXPOSE 3456

CMD ["node", "src/http-server.js"]
