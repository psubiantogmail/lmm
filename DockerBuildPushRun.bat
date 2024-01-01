docker build --file .\Docker\app_server\Dockerfile -t ghcr.io/psubiantogmail/lmm-ghcr:latest .
docker push ghcr.io/psubiantogmail/lmm-ghcr:latest
REM docker run -d -p 80:80 ghcr.io/psubiantogmail/lmm-ghcr:latest 
