FROM python:3.11-slim

WORKDIR /app

COPY server.py /app/server.py

RUN openssl req -x509 -newkey rsa:2048 -nodes \
    -keyout /app/server.key \
    -out /app/server.crt \
    -days 3650 \
    -subj "/CN=hostname-lab"

EXPOSE 5000 5002

CMD ["python", "/app/server.py"]
