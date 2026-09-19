# Deploy do FastAPI (EACT backend) no Railway

## 1. Copiar os ficheiros
Copie `Dockerfile`, `requirements.txt` e `.dockerignore` desta pasta para
dentro de `ops_backend/` (a raiz do seu projeto FastAPI, ao lado de `main.py`).

## 2. Ajustar o CORS no main.py
Troque isto:

```python
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
```

Por isto (permite localhost em dev e o domínio da Vercel em produção via variável de ambiente):

```python
import os

allowed_origins = ["http://localhost:3000"]
prod_origin = os.getenv("FRONTEND_ORIGIN")
if prod_origin:
    allowed_origins.append(prod_origin)

app.add_middleware(
    CORSMiddleware,
    allow_origins=allowed_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
```

## 3. Deploy no Railway
1. Vá a [railway.app](https://railway.app) e crie conta (login com GitHub)
2. **New Project → Deploy from GitHub repo** → selecione o repositório
   - Se o `ops_backend` e o `ops_frontend` estiverem no mesmo repo, no Railway
     defina **Root Directory** = `ops_backend` nas Settings do serviço
3. O Railway deteta o `Dockerfile` automaticamente e faz o build
4. Em **Settings → Networking**, clique em **Generate Domain** para obter um
   URL público (algo como `ops-backend-production.up.railway.app`)
5. Em **Variables**, adicione:
   - `FRONTEND_ORIGIN` = `https://o-seu-dominio.vercel.app`

## 4. Apontar o Next.js para o backend novo
No `route.ts` do Next.js, troque a URL fixa `http://127.0.0.1:8000/api/eact`
por uma variável de ambiente:

```ts
const backendUrl = process.env.EACT_BACKEND_URL ?? "http://127.0.0.1:8000/api/eact";

const response = await fetch(backendUrl, {
  method: "POST",
  headers: { "content-type": contentType },
  body: req.body,
  duplex: "half",
} as RequestInit & { duplex: "half" });
```

E na Vercel (Project Settings → Environment Variables):
- `EACT_BACKEND_URL` = `https://ops-backend-production.up.railway.app/api/eact`

Localmente, no `.env.local`, não precisa definir nada — continua a usar o
fallback `127.0.0.1:8000` para desenvolvimento.

## 5. Testar
```bash
curl -X POST "https://ops-backend-production.up.railway.app/api/eact" -F "file=@E:\EACT_0731.xlsx"
```
Sem ngrok, sem túnel, sem limite de tamanho de corpo — o Railway aceita
uploads grandes diretamente.

## Nota importante
O `requirements.txt` inclui `python-multipart`, que o FastAPI **exige**
internamente para processar `UploadFile` / `multipart/form-data`. Confirme
que já o tem instalado no seu `.venv` local também
(`pip install python-multipart`) — se não tiver, isso por si só já poderia
estar a causar falhas silenciosas mesmo antes de chegar ao ngrok.
