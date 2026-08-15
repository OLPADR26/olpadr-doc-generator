FROM node:20-slim

RUN apt-get update && apt-get install -y pandoc libreoffice --no-install-recommends && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY package.json .
RUN npm install
COPY . .

EXPOSE 3000
CMD ["node", "index.js"]
