FROM node:22-bookworm-slim

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci --include=dev

COPY . .
RUN mkdir -p /app/data && chmod 777 /app/data
RUN npm run build

ENV NODE_ENV=production
EXPOSE 43123

CMD ["npm", "start"]
