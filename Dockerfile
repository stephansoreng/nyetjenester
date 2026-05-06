# Stage 1: Build
FROM --platform=linux/amd64 node:22-alpine AS builder

WORKDIR /app

COPY package*.json ./
RUN npm ci

COPY . .

# Skip Excel import (requests.json is already committed); just run Vite
RUN npx vite build

# Stage 2: Serve
FROM --platform=linux/amd64 nginx:1.27-alpine

COPY --from=builder /app/dist /usr/share/nginx/html
COPY nginx.conf /etc/nginx/conf.d/default.conf

EXPOSE 80

CMD ["nginx", "-g", "daemon off;"]
