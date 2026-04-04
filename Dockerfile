FROM node:20-alpine

WORKDIR /app

COPY package.json ./
RUN npm install --omit=dev

COPY backend ./backend
COPY public ./public
COPY data ./data

ENV PORT=3000
ENV DATA_FILE=/app/data/store.json

EXPOSE 3000

CMD ["npm", "start"]
