# Hosting:
GitHub Pages:
- [Hello World!](https://akshitsingla.github.io/experiments-outlook_addin/)
- [Message Playgroung](https://akshitsingla.github.io/experiments-outlook_addin/message-read.html)

# Local Deployment

## Run pre-installed `python` web-server on MacOS
```sh
python3 -m http.server 3000
```

## Tunnel through ngrok
```sh
# CREATE DOCKER NETWORK
docker network create experiments_network

# RUN WEB SERVER
docker run \
  --network experiments_network \
  --name experiments_container_nginx \
  -v ./:/usr/share/nginx/html \
  -p 8080:80 \
  nginx:alpine

# TUNNEL THROUGH NGINX
export NGROK_AUTH_TOKEN="_YOUR_TOKEN_HERE_"
export NGROK_DOMAIN="_YOUR_DOMAIN_HERE_"
docker run  -it \
  -e NGROK_AUTHTOKEN=$NGROK_AUTH_TOKEN \
  --network experiments_network \
  --name experiments_container_ngrok \
  ngrok/ngrok:alpine \
  http --domain=$NGROK_DOMAIN experiments_container_nginx:80
```
**Note**
1. Please update the ngrok command with your `NGROK_AUTH_TOKEN`.
2. Please update your own `NGROK_DOMAIN` in above command as well as in manifest file (for local debugging).

# Credits
1. https://www.youtube.com/watch?v=ZWw-fJ7eldU&t=1000s 