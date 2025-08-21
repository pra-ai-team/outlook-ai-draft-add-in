FROM node:20-alpine

WORKDIR /app

# Install server dependencies
COPY server/package.json ./server/package.json
RUN cd server && npm install --only=production

# Copy source
COPY server ./server
COPY web ./web
COPY AI_setting ./AI_setting

ENV NODE_ENV=production
ENV PORT=3000

EXPOSE 3000

CMD ["node", "server/index.js"]


