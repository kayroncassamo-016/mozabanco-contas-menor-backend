

# from fastapi.middleware.cors import CORSMiddleware
# from datetime import datetime, date, timedelta
# from pathlib import Path
# from tempfile import NamedTemporaryFile
# from zipfile import ZipFile
# from concurrent.futures import ThreadPoolExecutor
# import asyncio
# import os
# import re
# import uuid

# from fastapi import FastAPI, UploadFile, File, Form, HTTPException

# from openpyxl import load_workbook

# app = FastAPI()

# # Em produção, defina FRONTEND_ORIGIN no Railway com o domínio
# # do seu frontend na Vercel (ex: https://o-seu-app.vercel.app)
# allowed_origins = ["http://localhost:3000", "https://mozabanco-relatorios-clientes-e-contas.vercel.app"]
# prod_origin = os.getenv("FRONTEND_ORIGIN")
# if prod_origin:
#     allowed_origins.append(prod_origin)

# app.add_middleware(
#     CORSMiddleware,
#     allow_origins=allowed_origins,
#     allow_credentials=True,
#     allow_methods=["*"],
#     allow_headers=["*"],
# )

# # ============================================================
# # ESTADO EM MEMÓRIA
# # ============================================================
# #
# # uploads: acompanha uploads em pedaços ainda em curso
# #   uploads[uploadId] = {"path": "...", "received": 3, "total": 29}
# #
# # jobs: acompanha o processamento depois do upload terminar
# #   jobs[job_id] = {"status": "processing"}
# #   jobs[job_id] = {"status": "done", "result": {...}}
# #   jobs[job_id] = {"status": "error", "error": "..."}
# #
# # NOTA: vive em memória do processo — se o serviço reiniciar a
# # meio de um upload/processamento, esse job perde-se. Aceitável
# # para o volume de uso atual (uso interno, um utilizador de cada
# # vez). Se precisar de robustez entre reinícios, trocar por Redis
# # ou uma tabela na base de dados.
# uploads: dict[str, dict] = {}
# jobs: dict[str, dict] = {}

# executor = ThreadPoolExecutor(max_workers=2)


# def normalize_text(value) -> str:
#     if value is None:
#         return ""

#     import unicodedata

#     return (
#         unicodedata.normalize("NFD", str(value))
#         .encode("ascii", "ignore")
#         .decode("ascii")
#         .lower()
#         .strip()
#     )


# def norm_val(value) -> str:
#     if value is None:
#         return ""

#     return (
#         str(value)
#         .replace("\u00a0", " ")
#         .strip()
#         .upper()
#     )


# # ============================================================
# # DT_ACT
# # ============================================================

# def parse_dt_act(value, ref_for_century: datetime | date):
#     """
#     DT_ACT vem como texto 'AA.MM.DD'.

#     Exemplo:
#         24.05.17 -> 17/05/2024

#     Se 20XX ficar no futuro relativamente à referência,
#     assume-se 19XX.
#     """

#     if not isinstance(value, str):
#         return None

#     value = value.strip()

#     match = re.fullmatch(r"(\d{2})\.(\d{2})\.(\d{2})", value)

#     if not match:
#         return None

#     yy, mm, dd = match.groups()

#     year = 2000 + int(yy)

#     try:
#         parsed = datetime(year, int(mm), int(dd))
#     except ValueError:
#         return None

#     if parsed > _as_datetime(ref_for_century):
#         year = 1900 + int(yy)

#         try:
#             parsed = datetime(year, int(mm), int(dd))
#         except ValueError:
#             return None

#     return parsed


# def _as_datetime(value) -> datetime:
#     if isinstance(value, datetime):
#         return value

#     if isinstance(value, date):
#         return datetime(value.year, value.month, value.day)

#     return datetime.now()


# # ============================================================
# # CREATED DATE
# # ============================================================

# def get_created_date(file_path: str):
#     """
#     Lê docProps/core.xml diretamente do XLSX (que é um ZIP).
#     """

#     try:
#         with ZipFile(file_path, "r") as archive:
#             try:
#                 xml_bytes = archive.read("docProps/core.xml")
#             except KeyError:
#                 return None

#         xml = xml_bytes.decode("utf-8", errors="ignore")

#         match = re.search(
#             r"<dcterms:created[^>]*>([^<]+)</dcterms:created>",
#             xml,
#         )

#         if not match:
#             return None

#         value = match.group(1)
#         value = value.replace("Z", "+00:00")

#         try:
#             parsed = datetime.fromisoformat(value)

#             if parsed.tzinfo is not None:
#                 parsed = parsed.astimezone().replace(tzinfo=None)

#             return parsed

#         except ValueError:
#             return None

#     except Exception:
#         return None


# # ============================================================
# # EACT PROCESSING (inalterado)
# # ============================================================

# def process_eact(file_path: str):
#     created_date = get_created_date(file_path)

#     workbook = load_workbook(
#         filename=file_path,
#         read_only=True,
#         data_only=True,
#     )

#     try:
#         total_bruto = 0

#         vistos_enc = set()
#         encerradas = 0

#         vistos_solta = set()
#         entidades_solta_count = 0

#         entidades_eact = {}

#         max_dt_act_raw = None
#         max_dt_act_parsed_for_now = None

#         consecutive_empty = 0

#         worksheet = None

#         for ws in workbook.worksheets:
#             sheet_name = normalize_text(ws.title)

#             if "export" in sheet_name:
#                 worksheet = ws
#                 break

#         if worksheet is None:
#             raise ValueError("Não foi encontrada a folha EXPORT no ficheiro EACT.")

#         rows = worksheet.iter_rows(values_only=True)

#         try:
#             header_row = next(rows)
#         except StopIteration:
#             raise ValueError("A folha EXPORT está vazia.")

#         header = [normalize_text(value) for value in header_row]

#         def find_column(name: str) -> int:
#             target = normalize_text(name)

#             try:
#                 return header.index(target)
#             except ValueError:
#                 return -1

#         c_sit = find_column("DSC_SIT")
#         c_contrato = find_column("COD_CONTRATO")
#         c_ent_solta = find_column("ENTIDADE_SOLTA")
#         c_ent_assoc = find_column("ENTIDADE_ASSOCIADA")
#         c_info_act = find_column("INFO_ACT")
#         c_dt_act = find_column("DT_ACT")
#         c_emp_part = find_column("DSC_EMP_PART")
#         c_doc_valido = find_column("DOCUMENTO_VALIDO")

#         required_columns = {
#             "DSC_SIT": c_sit,
#             "COD_CONTRATO": c_contrato,
#             "ENTIDADE_SOLTA": c_ent_solta,
#             "ENTIDADE_ASSOCIADA": c_ent_assoc,
#             "INFO_ACT": c_info_act,
#             "DT_ACT": c_dt_act,
#             "DSC_EMP_PART": c_emp_part,
#             "DOCUMENTO_VALIDO": c_doc_valido,
#         }

#         missing = [name for name, index in required_columns.items() if index == -1]

#         if missing:
#             raise ValueError("Colunas obrigatórias em falta: " + ", ".join(missing))

#         for row in rows:

#             if all(value is None for value in row):
#                 consecutive_empty += 1

#                 if consecutive_empty > 2000:
#                     break

#                 continue

#             consecutive_empty = 0
#             total_bruto += 1

#             def get_value(index):
#                 if index >= len(row):
#                     return None
#                 return row[index]

#             sit = norm_val(get_value(c_sit))
#             ent_solta = norm_val(get_value(c_ent_solta))

#             if sit == "ENCERRADA":
#                 key = get_value(c_contrato)
#                 if key not in vistos_enc:
#                     vistos_enc.add(key)
#                     encerradas += 1

#             if ent_solta == "SIM":
#                 key = get_value(c_ent_assoc)
#                 if key not in vistos_solta:
#                     vistos_solta.add(key)
#                     entidades_solta_count += 1

#             if sit != "ENCERRADA" and ent_solta == "NAO":

#                 ent = get_value(c_ent_assoc)

#                 if ent not in entidades_eact:
#                     entidades_eact[ent] = {
#                         "infoAct": get_value(c_info_act),
#                         "dtActRaw": get_value(c_dt_act),
#                         "empPart": get_value(c_emp_part),
#                         "docValido": get_value(c_doc_valido),
#                     }

#                 provisional_ref = (
#                     max_dt_act_parsed_for_now
#                     if max_dt_act_parsed_for_now is not None
#                     else datetime.now()
#                 )

#                 dt_raw = get_value(c_dt_act)
#                 d_parsed = parse_dt_act(dt_raw, provisional_ref)

#                 if d_parsed is not None and (
#                     max_dt_act_parsed_for_now is None
#                     or d_parsed > max_dt_act_parsed_for_now
#                 ):
#                     max_dt_act_parsed_for_now = d_parsed
#                     if isinstance(dt_raw, str):
#                         max_dt_act_raw = dt_raw

#         total_entidades = len(entidades_eact)

#         final_ref = created_date or datetime.now()

#         max_dt_act = (
#             parse_dt_act(max_dt_act_raw, final_ref) if max_dt_act_raw else None
#         )

#         fiabilizadas = 0

#         if max_dt_act is not None:

#             same_day = False

#             if created_date is not None:
#                 same_day = (
#                     max_dt_act.year == created_date.year
#                     and max_dt_act.month == created_date.month
#                     and max_dt_act.day == created_date.day
#                 )

#             ref_date = max_dt_act

#             if same_day:
#                 ref_date = ref_date - timedelta(days=1)

#             inicio_janela = ref_date.replace(year=ref_date.year - 2)

#             for entity in entidades_eact.values():

#                 info_act = entity["infoAct"]
#                 dt_act_raw = entity["dtActRaw"]
#                 emp_part = entity["empPart"]
#                 doc_valido = entity["docValido"]

#                 if norm_val(info_act) != "SIM":
#                     continue

#                 dt = parse_dt_act(dt_act_raw, final_ref)

#                 if dt is None:
#                     continue

#                 if dt < inicio_janela or dt > ref_date:
#                     continue

#                 ep = norm_val(emp_part)

#                 if ep == "EMPRESA":
#                     fiabilizadas += 1

#                 elif ep == "PARTICULAR":
#                     dv = norm_val(doc_valido)
#                     if dv in {"SIM", "SEM DATA DE VALIDADE"}:
#                         fiabilizadas += 1

#         por_fiabilizar = total_entidades - fiabilizadas

#         return {
#             "totalBruto": total_bruto,
#             "encerradas": encerradas,
#             "entidadesSoltas": entidades_solta_count,
#             "totalEntidades": total_entidades,
#             "fiabilizadas": fiabilizadas,
#             "porFiabilizar": por_fiabilizar,
#         }

#     finally:
#         workbook.close()


# # ============================================================
# # PROCESSAMENTO EM BACKGROUND
# # ============================================================

# async def run_processing(job_id: str, temp_path: str):
#     loop = asyncio.get_event_loop()

#     try:
#         result = await loop.run_in_executor(executor, process_eact, temp_path)
#         jobs[job_id] = {"status": "done", "result": result}

#     except Exception as exc:
#         jobs[job_id] = {
#             "status": "error",
#             "error": f"{type(exc).__name__}: {exc}",
#         }

#     finally:
#         try:
#             Path(temp_path).unlink(missing_ok=True)
#         except Exception:
#             pass


# # ============================================================
# # ROUTES
# # ============================================================

# @app.get("/")
# async def root():
#     return {"status": "ok", "service": "EACT FastAPI backend"}


# @app.post("/api/eact/upload-chunk")
# async def upload_chunk(
#     file: UploadFile = File(...),
#     uploadId: str = Form(...),
#     chunkIndex: int = Form(...),
#     totalChunks: int = Form(...),
# ):
#     """
#     Recebe UM pedaço do ficheiro de cada vez (ex: 5MB). Cada
#     pedaço é um request pequeno e rápido — nunca fica perto de
#     nenhum timeout de proxy, seja qual for a velocidade de upload
#     do utilizador.

#     O frontend envia os pedaços em ordem, um a um (espera a
#     resposta de cada um antes de enviar o seguinte), por isso
#     basta concatenar por ordem de chegada.
#     """

#     if uploadId not in uploads:
#         temp = NamedTemporaryFile(suffix=".xlsx", prefix="eact-", delete=False)
#         temp_path = temp.name
#         temp.close()

#         uploads[uploadId] = {
#             "path": temp_path,
#             "received": 0,
#             "total": totalChunks,
#         }

#     info = uploads[uploadId]

#     chunk_bytes = await file.read()

#     with open(info["path"], "ab") as f:
#         f.write(chunk_bytes)

#     info["received"] += 1

#     is_last = info["received"] >= info["total"]

#     if is_last:
#         job_id = str(uuid.uuid4())
#         jobs[job_id] = {"status": "processing"}

#         temp_path = info["path"]
#         del uploads[uploadId]

#         asyncio.create_task(run_processing(job_id, temp_path))

#         return {"done": True, "job_id": job_id}

#     return {"done": False, "received": info["received"], "total": info["total"]}


# @app.get("/api/eact/status/{job_id}")
# async def eact_status(job_id: str):
#     job = jobs.get(job_id)

#     if job is None:
#         raise HTTPException(status_code=404, detail="Job não encontrado.")

#     return job



from fastapi.middleware.cors import CORSMiddleware
from datetime import datetime, date, timedelta
from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import ZipFile
from concurrent.futures import ThreadPoolExecutor
import asyncio
import os
import re
import uuid

from fastapi import FastAPI, UploadFile, File, Form, HTTPException

from openpyxl import load_workbook

app = FastAPI()

# Em produção, defina FRONTEND_ORIGIN no Railway com o domínio
# do seu frontend na Vercel (ex: https://o-seu-app.vercel.app)
allowed_origins = ["http://localhost:3000", "https://mozabanco-relatorios-clientes-e-contas.vercel.app"]
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

# ============================================================
# ESTADO EM MEMÓRIA
# ============================================================
#
# uploads: acompanha uploads em pedaços ainda em curso
#   uploads[uploadId] = {"path": "...", "received": 3, "total": 29}
#
# jobs: acompanha o processamento depois do upload terminar
#   jobs[job_id] = {"status": "processing"}
#   jobs[job_id] = {"status": "done", "result": {...}}
#   jobs[job_id] = {"status": "error", "error": "..."}
#
# NOTA: vive em memória do processo — se o serviço reiniciar a
# meio de um upload/processamento, esse job perde-se. Aceitável
# para o volume de uso atual (uso interno, um utilizador de cada
# vez). Se precisar de robustez entre reinícios, trocar por Redis
# ou uma tabela na base de dados.
uploads: dict[str, dict] = {}
jobs: dict[str, dict] = {}

executor = ThreadPoolExecutor(max_workers=2)


def normalize_text(value) -> str:
    if value is None:
        return ""

    import unicodedata

    normalized = unicodedata.normalize("NFD", str(value))
    normalized = "".join(
        char for char in normalized
        if unicodedata.category(char) != "Mn"
    )
    return normalized.lower().strip()


def norm_val(value) -> str:
    if value is None:
        return ""

    return (
        str(value)
        .replace("\u00a0", " ")
        
        .strip()
        .upper()
    )


# ============================================================
# DT_ACT
# ============================================================

def parse_dt_act(value, ref_for_century: datetime | date):
    """
    DT_ACT vem como texto 'AA.MM.DD'.

    Exemplo:
        24.05.17 -> 17/05/2024

    Se 20XX ficar no futuro relativamente à referência,
    assume-se 19XX.
    """

    if not isinstance(value, str):
        return None

    value = value.strip()

    match = re.fullmatch(r"(\d{2})\.(\d{2})\.(\d{2})", value)

    if not match:
        return None

    yy, mm, dd = match.groups()

    year = 2000 + int(yy)

    try:
        parsed = datetime(year, int(mm), int(dd))
    except ValueError:
        return None

    if parsed > _as_datetime(ref_for_century):
        year = 1900 + int(yy)

        try:
            parsed = datetime(year, int(mm), int(dd))
        except ValueError:
            return None

    return parsed


def _as_datetime(value) -> datetime:
    if isinstance(value, datetime):
        return value

    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)

    return datetime.now()


# ============================================================
# CREATED DATE
# ============================================================

def get_created_date(file_path: str):
    """
    Lê docProps/core.xml diretamente do XLSX (que é um ZIP).
    """

    try:
        with ZipFile(file_path, "r") as archive:
            try:
                xml_bytes = archive.read("docProps/core.xml")
            except KeyError:
                return None

        xml = xml_bytes.decode("utf-8", errors="ignore")

        match = re.search(
            r"<dcterms:created[^>]*>([^<]+)</dcterms:created>",
            xml,
        )

        if not match:
            return None

        value = match.group(1)
        value = value.replace("Z", "+00:00")

        try:
            parsed = datetime.fromisoformat(value)

            if parsed.tzinfo is not None:
                parsed = parsed.astimezone().replace(tzinfo=None)

            return parsed

        except ValueError:
            return None

    except Exception:
        return None


# ============================================================
# EACT PROCESSING (inalterado)
# ============================================================

def process_eact(file_path: str):
    """
    Processa o EACT com a mesma lógica do route.ts.

    Diferença importante em relação à implementação anterior:
    - Total de Entidades EACT: conjunto de ENTIDADE_ASSOCIADA apenas para
      registos que não estejam ENCERRADA e cuja ENTIDADE_SOLTA seja NAO.
    - Fiabilizada: por entidade, considera a DT_ACT MAIS RECENTE entre
      todas as contas que cumprem INFO_ACT + Empresa/Particular + documento.
      A janela de datas só é aplicada depois de encontrar a DT_ACT máxima global.
    """
    created_date = get_created_date(file_path)

    workbook = load_workbook(
        filename=file_path,
        read_only=True,
        data_only=True,
    )

    try:
        total_bruto = 0

        vistos_enc = set()
        encerradas = 0

        vistos_solta = set()
        entidades_solta_count = 0

        # Equivalente ao totalEntidadesSet do route.ts.
        total_entidades_set = set()

        # Entidades que possuem pelo menos uma conta qualificável.
        # Guardamos as DT_ACT qualificadas para, no fim, verificar se
        # EXISTE pelo menos uma dentro da janela de 2 anos.
        candidatas = {}

        max_dt_act_raw = None
        max_dt_act_parsed_for_now = None

        consecutive_empty = 0

        worksheet = None

        # O route.ts percorre as folhas e processa todas as que contenham
        # "export" no nome. Normalmente existe apenas uma.
        for ws in workbook.worksheets:
            sheet_name = normalize_text(ws.title)

            if "export" not in sheet_name:
                continue

            rows = ws.iter_rows(values_only=True)

            try:
                header_row = next(rows)
            except StopIteration:
                continue

            header = [normalize_text(value) for value in header_row]

            def find_column(name: str) -> int:
                target = normalize_text(name)

                try:
                    return header.index(target)
                except ValueError:
                    return -1

            c_sit = find_column("DSC_SIT")
            c_contrato = find_column("COD_CONTRATO")
            c_ent_solta = find_column("ENTIDADE_SOLTA")
            c_ent_assoc = find_column("ENTIDADE_ASSOCIADA")
            c_info_act = find_column("INFO_ACT")
            c_dt_act = find_column("DT_ACT")
            c_emp_part = find_column("DSC_EMP_PART")
            c_doc_valido = find_column("DOCUMENTO_VALIDO")

            required_columns = {
                "DSC_SIT": c_sit,
                "COD_CONTRATO": c_contrato,
                "ENTIDADE_SOLTA": c_ent_solta,
                "ENTIDADE_ASSOCIADA": c_ent_assoc,
                "INFO_ACT": c_info_act,
                "DT_ACT": c_dt_act,
                "DSC_EMP_PART": c_emp_part,
                "DOCUMENTO_VALIDO": c_doc_valido,
            }

            missing = [
                name
                for name, index in required_columns.items()
                if index == -1
            ]

            if missing:
                raise ValueError(
                    "Colunas obrigatórias em falta: " + ", ".join(missing)
                )

            for row in rows:
                # Equivalente ao:
                # vals.every((v) => v === undefined || v === null)
                if all(value is None for value in row):
                    consecutive_empty += 1

                    # O ficheiro pode declarar 1.048.576 linhas, mas ter
                    # dados apenas nas primeiras centenas de milhares.
                    if consecutive_empty > 2000:
                        break

                    continue

                consecutive_empty = 0
                total_bruto += 1

                def get_value(index):
                    if index < 0 or index >= len(row):
                        return None
                    return row[index]

                sit = norm_val(get_value(c_sit))
                ent_solta = norm_val(get_value(c_ent_solta))

                # --------------------------------------------------------
                # ENCERRADAS
                # --------------------------------------------------------
                if sit == "ENCERRADA":
                    key = get_value(c_contrato)

                    if key not in vistos_enc:
                        vistos_enc.add(key)
                        encerradas += 1

                # --------------------------------------------------------
                # ENTIDADES SOLTAS
                # --------------------------------------------------------
                if ent_solta == "SIM":
                    key = get_value(c_ent_assoc)

                    if key not in vistos_solta:
                        vistos_solta.add(key)
                        entidades_solta_count += 1

                # --------------------------------------------------------
                # DATA MÁXIMA GLOBAL
                # --------------------------------------------------------
                # O route.ts calcula a máxima global antes de excluir
                # ENCERRADA/entidade solta.
                provisional_ref = (
                    max_dt_act_parsed_for_now
                    if max_dt_act_parsed_for_now is not None
                    else datetime.now()
                )

                dt_raw = get_value(c_dt_act)
                d_parsed_global = parse_dt_act(
                    dt_raw,
                    provisional_ref,
                )

                if d_parsed_global is not None and (
                    max_dt_act_parsed_for_now is None
                    or d_parsed_global > max_dt_act_parsed_for_now
                ):
                    max_dt_act_parsed_for_now = d_parsed_global

                    if isinstance(dt_raw, str):
                        max_dt_act_raw = dt_raw

                # --------------------------------------------------------
                # EXCLUSÕES PARA TOTAL ENTIDADES / FIABILIZADAS
                # --------------------------------------------------------
                if sit == "ENCERRADA" or ent_solta != "NAO":
                    continue

                ent = get_value(c_ent_assoc)

                # Total Entidades EACT:
                # uma entidade conta uma única vez.
                total_entidades_set.add(ent)

                # --------------------------------------------------------
                # CANDIDATA A FIABILIZADA
                # --------------------------------------------------------
                info_act = norm_val(get_value(c_info_act))
                emp_part = norm_val(get_value(c_emp_part))
                doc_valido = norm_val(get_value(c_doc_valido))

                qualifica = (
                    info_act == "SIM"
                    and (
                        emp_part == "EMPRESA"
                        or (
                            emp_part == "PARTICULAR"
                            and doc_valido in {
                                "SIM",
                                "SEM DATA DE VALIDADE",
                            }
                        )
                    )
                )

                if not qualifica:
                    continue

                # A data só é considerada candidata se for válida.
                dt = parse_dt_act(dt_raw, provisional_ref)

                if dt is None:
                    continue

                # IMPORTANTE:
                # A entidade fica candidata se EXISTIR pelo menos uma
                # conta qualificada. Não usamos a DT_ACT mais recente
                # como representante da entidade, porque uma conta mais
                # recente fora da janela não pode invalidar outra conta
                # qualificada que esteja dentro da janela.
                candidatas.setdefault(ent, []).append(dt)

        # Se não houver nenhuma folha EXPORT, reproduz o erro esperado.
        if not any(
            "export" in normalize_text(ws.title)
            for ws in workbook.worksheets
        ):
            raise ValueError(
                "Não foi encontrada a folha EXPORT no ficheiro EACT."
            )

        total_entidades = len(total_entidades_set)

        # ------------------------------------------------------------
        # REFERÊNCIA FINAL / DATA MÁXIMA REAL
        # ------------------------------------------------------------
        final_ref = created_date or datetime.now()

        max_dt_act = (
            parse_dt_act(max_dt_act_raw, final_ref)
            if max_dt_act_raw
            else None
        )

        # ------------------------------------------------------------
        # FIABILIZADAS
        # ------------------------------------------------------------
        fiabilizadas = 0

        if max_dt_act is not None:
            same_day = (
                created_date is not None
                and max_dt_act.year == created_date.year
                and max_dt_act.month == created_date.month
                and max_dt_act.day == created_date.day
            )

            ref_date = max_dt_act

            # Se a DT_ACT máxima tiver a mesma data da criação do ficheiro,
            # a janela termina no dia anterior.
            if same_day:
                ref_date = ref_date - timedelta(days=1)

            inicio_janela = ref_date.replace(
                year=ref_date.year - 2
            )

            # Uma entidade é fiabilizada se EXISTIR pelo menos uma
            # conta qualificada cuja DT_ACT esteja dentro da janela.
            for datas in candidatas.values():
                if any(inicio_janela <= dt <= ref_date for dt in datas):
                    fiabilizadas += 1

        por_fiabilizar = total_entidades - fiabilizadas

        return {
            "totalBruto": total_bruto,
            "encerradas": encerradas,
            "entidadesSoltas": entidades_solta_count,
            "totalEntidades": total_entidades,
            "fiabilizadas": fiabilizadas,
            "porFiabilizar": por_fiabilizar,
        }

    finally:
        workbook.close()


# ============================================================
# PROCESSAMENTO EM BACKGROUND
# ============================================================

async def run_processing(job_id: str, temp_path: str):
    loop = asyncio.get_event_loop()

    try:
        result = await loop.run_in_executor(executor, process_eact, temp_path)
        jobs[job_id] = {"status": "done", "result": result}

    except Exception as exc:
        jobs[job_id] = {
            "status": "error",
            "error": f"{type(exc).__name__}: {exc}",
        }

    finally:
        try:
            Path(temp_path).unlink(missing_ok=True)
        except Exception:
            pass


# ============================================================
# ROUTES
# ============================================================

@app.get("/")
async def root():
    return {"status": "ok", "service": "EACT FastAPI backend"}


@app.post("/api/eact/upload-chunk")
async def upload_chunk(
    file: UploadFile = File(...),
    uploadId: str = Form(...),
    chunkIndex: int = Form(...),
    totalChunks: int = Form(...),
):
    """
    Recebe UM pedaço do ficheiro de cada vez (ex: 5MB). Cada
    pedaço é um request pequeno e rápido — nunca fica perto de
    nenhum timeout de proxy, seja qual for a velocidade de upload
    do utilizador.

    O frontend envia os pedaços em ordem, um a um (espera a
    resposta de cada um antes de enviar o seguinte), por isso
    basta concatenar por ordem de chegada.
    """

    if uploadId not in uploads:
        temp = NamedTemporaryFile(suffix=".xlsx", prefix="eact-", delete=False)
        temp_path = temp.name
        temp.close()

        uploads[uploadId] = {
            "path": temp_path,
            "received": 0,
            "total": totalChunks,
        }

    info = uploads[uploadId]

    chunk_bytes = await file.read()

    with open(info["path"], "ab") as f:
        f.write(chunk_bytes)

    info["received"] += 1

    is_last = info["received"] >= info["total"]

    if is_last:
        job_id = str(uuid.uuid4())
        jobs[job_id] = {"status": "processing"}

        temp_path = info["path"]
        del uploads[uploadId]

        asyncio.create_task(run_processing(job_id, temp_path))

        return {"done": True, "job_id": job_id}

    return {"done": False, "received": info["received"], "total": info["total"]}


@app.get("/api/eact/status/{job_id}")
async def eact_status(job_id: str):
    job = jobs.get(job_id)

    if job is None:
        raise HTTPException(status_code=404, detail="Job não encontrado.")

    return job