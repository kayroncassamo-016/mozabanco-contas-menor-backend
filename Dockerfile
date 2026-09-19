FROM python:3.12-slim

WORKDIR /app

# Instala as dependências primeiro (aproveita cache do Docker entre deploys)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copia o resto do código da aplicação
COPY . .

# Railway injeta a porta via variável de ambiente $PORT
ENV PORT=8000
EXPOSE 8000

# Timeout generoso no keep-alive para aguentar uploads grandes;
# --timeout-keep-alive em segundos
CMD uvicorn main:app --host 0.0.0.0 --port ${PORT} --timeout-keep-alive 300
