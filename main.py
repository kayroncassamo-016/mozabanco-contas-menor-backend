
#######################################################################################

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

#     normalized = unicodedata.normalize("NFD", str(value))
#     normalized = "".join(
#         char for char in normalized
#         if unicodedata.category(char) != "Mn"
#     )
#     return normalized.lower().strip()


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
#     """
#     Processa o EACT com a mesma lógica do route.ts.

#     Diferença importante em relação à implementação anterior:
#     - Total de Entidades EACT: conjunto de ENTIDADE_ASSOCIADA apenas para
#       registos que não estejam ENCERRADA e cuja ENTIDADE_SOLTA seja NAO.
#     - Fiabilizada: por entidade, considera a DT_ACT MAIS RECENTE entre
#       todas as contas que cumprem INFO_ACT + Empresa/Particular + documento.
#       A janela de datas só é aplicada depois de encontrar a DT_ACT máxima global.
#     """
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

#         # Equivalente ao totalEntidadesSet do route.ts.
#         total_entidades_set = set()

#         # Entidades que possuem pelo menos uma conta qualificável.
#         # Guardamos as DT_ACT qualificadas para, no fim, verificar se
#         # EXISTE pelo menos uma dentro da janela de 2 anos.
#         candidatas = {}

#         max_dt_act_raw = None
#         max_dt_act_parsed_for_now = None

#         consecutive_empty = 0

#         worksheet = None

#         # O route.ts percorre as folhas e processa todas as que contenham
#         # "export" no nome. Normalmente existe apenas uma.
#         for ws in workbook.worksheets:
#             sheet_name = normalize_text(ws.title)

#             if "export" not in sheet_name:
#                 continue

#             rows = ws.iter_rows(values_only=True)

#             try:
#                 header_row = next(rows)
#             except StopIteration:
#                 continue

#             header = [normalize_text(value) for value in header_row]

#             def find_column(name: str) -> int:
#                 target = normalize_text(name)

#                 try:
#                     return header.index(target)
#                 except ValueError:
#                     return -1

#             c_sit = find_column("DSC_SIT")
#             c_contrato = find_column("COD_CONTRATO")
#             c_ent_solta = find_column("ENTIDADE_SOLTA")
#             c_ent_assoc = find_column("ENTIDADE_ASSOCIADA")
#             c_info_act = find_column("INFO_ACT")
#             c_dt_act = find_column("DT_ACT")
#             c_emp_part = find_column("DSC_EMP_PART")
#             c_doc_valido = find_column("DOCUMENTO_VALIDO")

#             required_columns = {
#                 "DSC_SIT": c_sit,
#                 "COD_CONTRATO": c_contrato,
#                 "ENTIDADE_SOLTA": c_ent_solta,
#                 "ENTIDADE_ASSOCIADA": c_ent_assoc,
#                 "INFO_ACT": c_info_act,
#                 "DT_ACT": c_dt_act,
#                 "DSC_EMP_PART": c_emp_part,
#                 "DOCUMENTO_VALIDO": c_doc_valido,
#             }

#             missing = [
#                 name
#                 for name, index in required_columns.items()
#                 if index == -1
#             ]

#             if missing:
#                 raise ValueError(
#                     "Colunas obrigatórias em falta: " + ", ".join(missing)
#                 )

#             for row in rows:
#                 # Equivalente ao:
#                 # vals.every((v) => v === undefined || v === null)
#                 if all(value is None for value in row):
#                     consecutive_empty += 1

#                     # O ficheiro pode declarar 1.048.576 linhas, mas ter
#                     # dados apenas nas primeiras centenas de milhares.
#                     if consecutive_empty > 2000:
#                         break

#                     continue

#                 consecutive_empty = 0
#                 total_bruto += 1

#                 def get_value(index):
#                     if index < 0 or index >= len(row):
#                         return None
#                     return row[index]

#                 sit = norm_val(get_value(c_sit))
#                 ent_solta = norm_val(get_value(c_ent_solta))

#                 # --------------------------------------------------------
#                 # ENCERRADAS
#                 # --------------------------------------------------------
#                 if sit == "ENCERRADA":
#                     key = get_value(c_contrato)

#                     if key not in vistos_enc:
#                         vistos_enc.add(key)
#                         encerradas += 1

#                 # --------------------------------------------------------
#                 # ENTIDADES SOLTAS
#                 # --------------------------------------------------------
#                 if ent_solta == "SIM":
#                     key = get_value(c_ent_assoc)

#                     if key not in vistos_solta:
#                         vistos_solta.add(key)
#                         entidades_solta_count += 1

#                 # --------------------------------------------------------
#                 # DATA MÁXIMA GLOBAL
#                 # --------------------------------------------------------
#                 # O route.ts calcula a máxima global antes de excluir
#                 # ENCERRADA/entidade solta.
#                 provisional_ref = (
#                     max_dt_act_parsed_for_now
#                     if max_dt_act_parsed_for_now is not None
#                     else datetime.now()
#                 )

#                 dt_raw = get_value(c_dt_act)
#                 d_parsed_global = parse_dt_act(
#                     dt_raw,
#                     provisional_ref,
#                 )

#                 if d_parsed_global is not None and (
#                     max_dt_act_parsed_for_now is None
#                     or d_parsed_global > max_dt_act_parsed_for_now
#                 ):
#                     max_dt_act_parsed_for_now = d_parsed_global

#                     if isinstance(dt_raw, str):
#                         max_dt_act_raw = dt_raw

#                 # --------------------------------------------------------
#                 # EXCLUSÕES PARA TOTAL ENTIDADES / FIABILIZADAS
#                 # --------------------------------------------------------
#                 if sit == "ENCERRADA" or ent_solta != "NAO":
#                     continue

#                 ent = get_value(c_ent_assoc)

#                 # Total Entidades EACT:
#                 # uma entidade conta uma única vez.
#                 total_entidades_set.add(ent)

#                 # --------------------------------------------------------
#                 # CANDIDATA A FIABILIZADA
#                 # --------------------------------------------------------
#                 info_act = norm_val(get_value(c_info_act))
#                 emp_part = norm_val(get_value(c_emp_part))
#                 doc_valido = norm_val(get_value(c_doc_valido))

#                 qualifica = (
#                     info_act == "SIM"
#                     and (
#                         emp_part == "EMPRESA"
#                         or (
#                             emp_part == "PARTICULAR"
#                             and doc_valido in {
#                                 "SIM",
#                                 "SEM DATA DE VALIDADE",
#                             }
#                         )
#                     )
#                 )

#                 if not qualifica:
#                     continue

#                 # A data só é considerada candidata se for válida.
#                 dt = parse_dt_act(dt_raw, provisional_ref)

#                 if dt is None:
#                     continue

#                 # IMPORTANTE:
#                 # A entidade fica candidata se EXISTIR pelo menos uma
#                 # conta qualificada. Não usamos a DT_ACT mais recente
#                 # como representante da entidade, porque uma conta mais
#                 # recente fora da janela não pode invalidar outra conta
#                 # qualificada que esteja dentro da janela.
#                 candidatas.setdefault(ent, []).append(dt)

#         # Se não houver nenhuma folha EXPORT, reproduz o erro esperado.
#         if not any(
#             "export" in normalize_text(ws.title)
#             for ws in workbook.worksheets
#         ):
#             raise ValueError(
#                 "Não foi encontrada a folha EXPORT no ficheiro EACT."
#             )

#         total_entidades = len(total_entidades_set)

#         # ------------------------------------------------------------
#         # REFERÊNCIA FINAL / DATA MÁXIMA REAL
#         # ------------------------------------------------------------
#         final_ref = created_date or datetime.now()

#         max_dt_act = (
#             parse_dt_act(max_dt_act_raw, final_ref)
#             if max_dt_act_raw
#             else None
#         )

#         # ------------------------------------------------------------
#         # FIABILIZADAS
#         # ------------------------------------------------------------
#         fiabilizadas = 0

#         if max_dt_act is not None:
#             same_day = (
#                 created_date is not None
#                 and max_dt_act.year == created_date.year
#                 and max_dt_act.month == created_date.month
#                 and max_dt_act.day == created_date.day
#             )

#             ref_date = max_dt_act

#             # Se a DT_ACT máxima tiver a mesma data da criação do ficheiro,
#             # a janela termina no dia anterior.
#             if same_day:
#                 ref_date = ref_date - timedelta(days=1)

#             inicio_janela = ref_date.replace(
#                 year=ref_date.year - 2
#             )

#             # Uma entidade é fiabilizada se EXISTIR pelo menos uma
#             # conta qualificada cuja DT_ACT esteja dentro da janela.
#             for datas in candidatas.values():
#                 if any(inicio_janela <= dt <= ref_date for dt in datas):
#                     fiabilizadas += 1

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

##############################################################################


###########################################################################
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


# # ============================================================
# # CORS
# # ============================================================

# allowed_origins = [
#     "http://localhost:3000",
#     "https://mozabanco-relatorios-clientes-e-contas.vercel.app",
# ]

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

# uploads: dict[str, dict] = {}
# jobs: dict[str, dict] = {}

# executor = ThreadPoolExecutor(max_workers=2)


# # ============================================================
# # NORMALIZAÇÃO
# # ============================================================

# def normalize_text(value) -> str:
#     if value is None:
#         return ""

#     import unicodedata

#     normalized = unicodedata.normalize(
#         "NFD",
#         str(value),
#     )

#     normalized = "".join(
#         char
#         for char in normalized
#         if unicodedata.category(char) != "Mn"
#     )

#     return normalized.lower().strip()


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
# # DATAS
# # ============================================================

# def _as_datetime(value) -> datetime:
#     if isinstance(value, datetime):
#         return value

#     if isinstance(value, date):
#         return datetime(
#             value.year,
#             value.month,
#             value.day,
#         )

#     return datetime.now()


# def parse_dt_act(
#     value,
#     ref_for_century: datetime | date,
# ):
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

#     match = re.fullmatch(
#         r"(\d{2})\.(\d{2})\.(\d{2})",
#         value,
#     )

#     if not match:
#         return None

#     yy, mm, dd = match.groups()

#     year = 2000 + int(yy)

#     try:
#         parsed = datetime(
#             year,
#             int(mm),
#             int(dd),
#         )

#     except ValueError:
#         return None

#     if parsed > _as_datetime(ref_for_century):

#         year = 1900 + int(yy)

#         try:
#             parsed = datetime(
#                 year,
#                 int(mm),
#                 int(dd),
#             )

#         except ValueError:
#             return None

#     return parsed


# # ============================================================
# # CREATED DATE
# # ============================================================

# def get_created_date(file_path: str):
#     """
#     Lê docProps/core.xml diretamente do XLSX.
#     """

#     try:

#         with ZipFile(
#             file_path,
#             "r",
#         ) as archive:

#             try:
#                 xml_bytes = archive.read(
#                     "docProps/core.xml"
#                 )

#             except KeyError:
#                 return None

#         xml = xml_bytes.decode(
#             "utf-8",
#             errors="ignore",
#         )

#         match = re.search(
#             r"<dcterms:created[^>]*>([^<]+)</dcterms:created>",
#             xml,
#         )

#         if not match:
#             return None

#         value = match.group(1)

#         value = value.replace(
#             "Z",
#             "+00:00",
#         )

#         try:

#             parsed = datetime.fromisoformat(
#                 value
#             )

#             if parsed.tzinfo is not None:

#                 parsed = (
#                     parsed
#                     .astimezone()
#                     .replace(tzinfo=None)
#                 )

#             return parsed

#         except ValueError:
#             return None

#     except Exception:
#         return None


# # ============================================================
# # FUNÇÕES AUXILIARES DO PROCESSAMENTO
# # ============================================================

# def same_calendar_day(
#     a: datetime,
#     b: datetime,
# ) -> bool:

#     return (
#         a.year == b.year
#         and a.month == b.month
#         and a.day == b.day
#     )


# def subtract_two_years(
#     value: datetime,
# ) -> datetime:

#     try:

#         return value.replace(
#             year=value.year - 2
#         )

#     except ValueError:

#         # Segurança para 29/02.
#         return value.replace(
#             year=value.year - 2,
#             day=28,
#         )


# def get_value(
#     row,
#     index,
# ):

#     if index < 0 or index >= len(row):
#         return None

#     return row[index]


# def find_columns(header):

#     normalized = [
#         normalize_text(value)
#         for value in header
#     ]

#     def find(*names):

#         for name in names:

#             target = normalize_text(name)

#             if target in normalized:

#                 return normalized.index(
#                     target
#                 )

#         return -1

#     return {

#         # O procedimento manual usa COD_SIT.
#         # DSC_SIT fica como fallback.
#         "cod_sit": find("COD_SIT"),

#         "dsc_sit": find("DSC_SIT"),

#         "contrato":
#             find("COD_CONTRATO"),

#         "ent_solta":
#             find("ENTIDADE_SOLTA"),

#         "ent_assoc":
#             find("ENTIDADE_ASSOCIADA"),

#         "info_act":
#             find("INFO_ACT"),

#         "dt_act":
#             find("DT_ACT"),

#         "emp_part":
#             find("DSC_EMP_PART"),

#         "doc_valido":
#             find("DOCUMENTO_VALIDO"),
#     }


# def is_encerrada(
#     row,
#     columns,
# ) -> bool:

#     cod_sit = norm_val(
#         get_value(
#             row,
#             columns["cod_sit"],
#         )
#     )

#     dsc_sit = norm_val(
#         get_value(
#             row,
#             columns["dsc_sit"],
#         )
#     )

#     # Regra manual:
#     #
#     # COD_SIT = E
#     #
#     # DSC_SIT é apenas fallback para ficheiros
#     # que não tenham COD_SIT.

#     return (
#         cod_sit == "E"
#         or (
#             columns["cod_sit"] < 0
#             and dsc_sit == "ENCERRADA"
#         )
#     )


# def row_is_fiabilizada(
#     row,
#     columns,
#     inicio_janela: datetime,
#     ref_date: datetime,
# ) -> bool:

#     # ========================================================
#     # FILTROS BASE
#     # ========================================================

#     # COD_SIT != E
#     if is_encerrada(
#         row,
#         columns,
#     ):
#         return False

#     # ENTIDADE_SOLTA = NAO
#     if norm_val(
#         get_value(
#             row,
#             columns["ent_solta"],
#         )
#     ) != "NAO":

#         return False

#     # ========================================================
#     # INFO_ACT = SIM
#     # ========================================================

#     if norm_val(
#         get_value(
#             row,
#             columns["info_act"],
#         )
#     ) != "SIM":

#         return False

#     # ========================================================
#     # DT_ACT NOS ÚLTIMOS DOIS ANOS
#     # ========================================================

#     dt = parse_dt_act(
#         get_value(
#             row,
#             columns["dt_act"],
#         ),
#         ref_date,
#     )

#     if dt is None:
#         return False

#     # O limite inferior da janela é exclusivo.
#     # Ex.: com ref_date = 03/08/2026, 03/08/2024 não recebe X.
#     if dt <= inicio_janela:
#         return False

#     if dt > ref_date:
#         return False

#     # ========================================================
#     # EMPRESA
#     # ========================================================

#     emp_part = norm_val(
#         get_value(
#             row,
#             columns["emp_part"],
#         )
#     )

#     if emp_part == "EMPRESA":
#         return True

#     # ========================================================
#     # PARTICULAR
#     # ========================================================

#     if emp_part == "PARTICULAR":

#         doc_valido = norm_val(
#             get_value(
#                 row,
#                 columns["doc_valido"],
#             )
#         )

#         return doc_valido in {
#             "SIM",
#             "SEM DATA DE VALIDADE",
#         }

#     return False


# # ============================================================
# # EACT PROCESSING
# # ============================================================

# def process_eact(
#     file_path: str,
# ):
#     """
#     Processa o EACT em duas passagens.

#     PASSAGEM 1:
#         - totalBruto
#         - encerradas
#         - entidadesSoltas
#         - totalEntidades
#         - maior DT_ACT

#     PASSAGEM 2:
#         reproduz a coluna manual Fiabilizados.

#     Cada linha é analisada individualmente.

#     Se uma linha satisfizer os filtros,
#     ela receberia X no Excel.

#     No final:

#         COUNT(DISTINCT ENTIDADE_ASSOCIADA)

#     entre as linhas que receberiam X.
#     """

#     created_date = get_created_date(
#         file_path
#     )

#     # ========================================================
#     # PASSAGEM 1
#     # ========================================================

#     workbook = load_workbook(
#         filename=file_path,
#         read_only=True,
#         data_only=True,
#     )

#     try:

#         worksheet = None

#         for ws in workbook.worksheets:

#             if (
#                 "export"
#                 in normalize_text(
#                     ws.title
#                 )
#             ):

#                 worksheet = ws
#                 break

#         if worksheet is None:

#             raise ValueError(
#                 "Não foi encontrada a folha "
#                 "EXPORT no ficheiro EACT."
#             )

#         rows = worksheet.iter_rows(
#             values_only=True
#         )

#         try:

#             header_row = next(rows)

#         except StopIteration:

#             raise ValueError(
#                 "A folha EXPORT está vazia."
#             )

#         columns = find_columns(
#             header_row
#         )

#         # ====================================================
#         # COLUNAS OBRIGATÓRIAS
#         # ====================================================

#         required = {

#             "ENTIDADE_SOLTA":
#                 columns["ent_solta"],

#             "ENTIDADE_ASSOCIADA":
#                 columns["ent_assoc"],

#             "INFO_ACT":
#                 columns["info_act"],

#             "DT_ACT":
#                 columns["dt_act"],

#             "DSC_EMP_PART":
#                 columns["emp_part"],

#             "DOCUMENTO_VALIDO":
#                 columns["doc_valido"],
#         }

#         missing = [
#             name
#             for name, index
#             in required.items()
#             if index == -1
#         ]

#         if (
#             columns["cod_sit"] < 0
#             and columns["dsc_sit"] < 0
#         ):

#             missing.append(
#                 "COD_SIT/DSC_SIT"
#             )

#         if columns["contrato"] < 0:

#             missing.append(
#                 "COD_CONTRATO"
#             )

#         if missing:

#             raise ValueError(
#                 "Colunas obrigatórias "
#                 "em falta: "
#                 + ", ".join(
#                     dict.fromkeys(
#                         missing
#                     )
#                 )
#             )

#         # ====================================================
#         # CONTADORES
#         # ====================================================

#         total_bruto = 0

#         vistos_enc = set()
#         encerradas = 0

#         vistos_solta = set()
#         entidades_solta_count = 0

#         total_entidades_set = set()

#         max_dt_act = None

#         consecutive_empty = 0

#         # ====================================================
#         # LEITURA DAS LINHAS
#         # ====================================================

#         for row in rows:

#             if all(
#                 value is None
#                 for value in row
#             ):

#                 consecutive_empty += 1

#                 if (
#                     consecutive_empty
#                     > 2000
#                 ):

#                     break

#                 continue

#             consecutive_empty = 0

#             total_bruto += 1

#             # ==================================================
#             # ENCERRADAS
#             # ==================================================

#             if is_encerrada(
#                 row,
#                 columns,
#             ):

#                 key = get_value(
#                     row,
#                     columns["contrato"],
#                 )

#                 if key not in vistos_enc:

#                     vistos_enc.add(key)

#                     encerradas += 1

#             # ==================================================
#             # ENTIDADES SOLTAS
#             # ==================================================

#             ent_solta = norm_val(
#                 get_value(
#                     row,
#                     columns["ent_solta"],
#                 )
#             )

#             if ent_solta == "SIM":

#                 key = get_value(
#                     row,
#                     columns["ent_assoc"],
#                 )

#                 if key not in vistos_solta:

#                     vistos_solta.add(key)

#                     entidades_solta_count += 1

#             # ==================================================
#             # UNIVERSO
#             # ==================================================

#             if is_encerrada(
#                 row,
#                 columns,
#             ):

#                 continue

#             if ent_solta != "NAO":

#                 continue

#             total_entidades_set.add(
#                 get_value(
#                     row,
#                     columns["ent_assoc"],
#                 )
#             )

#             # ==================================================
#             # MAIOR DT_ACT
#             # ==================================================

#             dt = parse_dt_act(
#                 get_value(
#                     row,
#                     columns["dt_act"],
#                 ),
#                 created_date
#                 or datetime.now(),
#             )

#             if dt is not None:

#                 if (
#                     max_dt_act is None
#                     or dt > max_dt_act
#                 ):

#                     max_dt_act = dt

#         total_entidades = len(
#             total_entidades_set
#         )

#     finally:

#         workbook.close()

#     # ========================================================
#     # DATA DE REFERÊNCIA
#     # ========================================================

#     # Regra da data de referência:
#     #
#     # 1. Se a última DT_ACT for igual à data de criação
#     #    do ficheiro, usa-se o dia anterior à criação.
#     #
#     # 2. Se a última DT_ACT for anterior à criação,
#     #    usa-se a própria última DT_ACT.
#     #
#     # 3. Se a última DT_ACT for posterior à criação,
#     #    usa-se o dia anterior à criação.
#     #
#     # Assim, uma DT_ACT futura em relação ao ficheiro não
#     # desloca artificialmente a janela de fiabilização.

#     ref_date = None

#     if max_dt_act is not None:

#         if created_date is None:

#             # Sem data de criação, não há comparação possível.
#             # Mantém-se a última DT_ACT como referência.
#             ref_date = datetime(
#                 max_dt_act.year,
#                 max_dt_act.month,
#                 max_dt_act.day,
#             )

#         else:

#             max_dt_act_day = datetime(
#                 max_dt_act.year,
#                 max_dt_act.month,
#                 max_dt_act.day,
#             )

#             created_date_day = datetime(
#                 created_date.year,
#                 created_date.month,
#                 created_date.day,
#             )

#             if same_calendar_day(
#                 max_dt_act_day,
#                 created_date_day,
#             ):

#                 # Última DT_ACT = data de criação.
#                 ref_date = created_date_day - timedelta(
#                     days=1
#                 )

#             elif max_dt_act_day < created_date_day:

#                 # Última DT_ACT anterior à criação.
#                 ref_date = max_dt_act_day

#             else:

#                 # Última DT_ACT posterior à criação.
#                 # Usa-se o dia anterior à criação.
#                 ref_date = created_date_day - timedelta(
#                     days=1
#                 )

#     if ref_date is not None:

#         inicio_janela = (
#             subtract_two_years(
#                 ref_date
#             )
#         )

#     else:

#         inicio_janela = None

#     # ========================================================
#     # PASSAGEM 2
#     # ========================================================

#     fiabilizadas_set = set()

#     if (
#         ref_date is not None
#         and inicio_janela is not None
#     ):

#         workbook = load_workbook(
#             filename=file_path,
#             read_only=True,
#             data_only=True,
#         )

#         try:

#             worksheet = None

#             for ws in workbook.worksheets:

#                 if (
#                     "export"
#                     in normalize_text(
#                         ws.title
#                     )
#                 ):

#                     worksheet = ws
#                     break

#             if worksheet is None:

#                 raise ValueError(
#                     "Não foi encontrada a "
#                     "folha EXPORT."
#                 )

#             rows = worksheet.iter_rows(
#                 values_only=True
#             )

#             try:

#                 header_row = next(rows)

#             except StopIteration:

#                 raise ValueError(
#                     "A folha EXPORT está vazia."
#                 )

#             columns = find_columns(
#                 header_row
#             )

#             consecutive_empty = 0

#             for row in rows:

#                 if all(
#                     value is None
#                     for value in row
#                 ):

#                     consecutive_empty += 1

#                     if (
#                         consecutive_empty
#                         > 2000
#                     ):

#                         break

#                     continue

#                 consecutive_empty = 0

#                 # ==============================================
#                 # ESTA LINHA RECEBERIA X?
#                 # ==============================================

#                 if not row_is_fiabilizada(
#                     row,
#                     columns,
#                     inicio_janela,
#                     ref_date,
#                 ):

#                     continue

#                 # ==============================================
#                 # SIM -> X
#                 # ==============================================

#                 entidade = get_value(
#                     row,
#                     columns["ent_assoc"],
#                 )

#                 fiabilizadas_set.add(
#                     entidade
#                 )

#         finally:

#             workbook.close()

#     # ========================================================
#     # RESULTADO FINAL
#     # ========================================================

#     fiabilizadas = len(
#         fiabilizadas_set
#     )

#     por_fiabilizar = (
#         total_entidades
#         - fiabilizadas
#     )

#     return {
#         "totalBruto": total_bruto,

#         "encerradas":
#             encerradas,

#         "entidadesSoltas":
#             entidades_solta_count,

#         "totalEntidades":
#             total_entidades,

#         "fiabilizadas":
#             fiabilizadas,

#         "porFiabilizar":
#             por_fiabilizar,
#     }


# # ============================================================
# # PROCESSAMENTO EM BACKGROUND
# # ============================================================

# async def run_processing(
#     job_id: str,
#     temp_path: str,
# ):

#     loop = asyncio.get_event_loop()

#     try:

#         result = await loop.run_in_executor(
#             executor,
#             process_eact,
#             temp_path,
#         )

#         jobs[job_id] = {
#             "status": "done",
#             "result": result,
#         }

#     except Exception as exc:

#         jobs[job_id] = {
#             "status": "error",
#             "error": (
#                 f"{type(exc).__name__}: {exc}"
#             ),
#         }

#     finally:

#         try:

#             Path(
#                 temp_path
#             ).unlink(
#                 missing_ok=True
#             )

#         except Exception:
#             pass


# # ============================================================
# # ROUTES
# # ============================================================

# @app.get("/")
# async def root():

#     return {
#         "status": "ok",
#         "service": "EACT FastAPI backend",
#     }


# @app.post("/api/eact/upload-chunk")
# async def upload_chunk(
#     file: UploadFile = File(...),
#     uploadId: str = Form(...),
#     chunkIndex: int = Form(...),
#     totalChunks: int = Form(...),
# ):
#     """
#     Recebe UM pedaço do ficheiro de cada vez.

#     Cada pedaço é escrito no ficheiro temporário.
#     Quando o último pedaço chega, o processamento
#     é iniciado em background.
#     """

#     if uploadId not in uploads:

#         temp = NamedTemporaryFile(
#             suffix=".xlsx",
#             prefix="eact-",
#             delete=False,
#         )

#         temp_path = temp.name

#         temp.close()

#         uploads[uploadId] = {
#             "path": temp_path,
#             "received": 0,
#             "total": totalChunks,
#         }

#     info = uploads[uploadId]

#     chunk_bytes = await file.read()

#     with open(
#         info["path"],
#         "ab",
#     ) as f:

#         f.write(chunk_bytes)

#     info["received"] += 1

#     is_last = (
#         info["received"]
#         >= info["total"]
#     )

#     if is_last:

#         job_id = str(
#             uuid.uuid4()
#         )

#         jobs[job_id] = {
#             "status": "processing"
#         }

#         temp_path = info["path"]

#         del uploads[uploadId]

#         asyncio.create_task(
#             run_processing(
#                 job_id,
#                 temp_path,
#             )
#         )

#         return {
#             "done": True,
#             "job_id": job_id,
#         }

#     return {
#         "done": False,
#         "received": info["received"],
#         "total": info["total"],
#     }


# @app.get("/api/eact/status/{job_id}")
# async def eact_status(
#     job_id: str,
# ):

#     job = jobs.get(job_id)

#     if job is None:

#         raise HTTPException(
#             status_code=404,
#             detail="Job não encontrado.",
#         )

#     return job



##########################################################################
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
from typing import Any, Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException

from openpyxl import load_workbook


app = FastAPI()


# ============================================================
# CORS
# ============================================================

allowed_origins = [
    "http://localhost:3000",
    "https://mozabanco-relatorios-clientes-e-contas.vercel.app",
]

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

uploads: dict[str, dict] = {}
jobs: dict[str, dict] = {}

executor = ThreadPoolExecutor(max_workers=2)


# ============================================================
# NORMALIZAÇÃO
# ============================================================

def normalize_text(value) -> str:
    if value is None:
        return ""

    import unicodedata

    normalized = unicodedata.normalize(
        "NFD",
        str(value),
    )

    normalized = "".join(
        char
        for char in normalized
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
# DATAS
# ============================================================

def _as_datetime(value) -> datetime:
    if isinstance(value, datetime):
        return value

    if isinstance(value, date):
        return datetime(
            value.year,
            value.month,
            value.day,
        )

    return datetime.now()




def parse_dt_act(
    value,
    reference_date: datetime,
) -> Optional[datetime]:
    if value is None:
        return None

    if isinstance(value, datetime):
        return value.replace(tzinfo=None)

    if isinstance(value, date):
        return datetime(
            value.year,
            value.month,
            value.day,
        )

    text = str(value).strip()

    if not text:
        return None

    # Formato: YYYY-MM-DD HH:MM:SS
    # Exemplo: 2025-03-29 00:00:00
    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%Y.%m.%d",
    ):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass

    # Formato: YY-MM-DD
    # Exemplo: 25-03-29
    for fmt in (
        "%y-%m-%d",
        "%y.%m.%d",
    ):
        try:
            parsed = datetime.strptime(text, fmt)

            # Mantém a mesma lógica do formato YY.MM.DD:
            # inicialmente considera 20xx.
            if parsed > reference_date:
                parsed = parsed.replace(
                    year=parsed.year - 100
                )

            return parsed

        except ValueError:
            pass

    return None


# ============================================================
# CREATED DATE
# ============================================================

def get_created_date(file_path: str):
    """
    Lê docProps/core.xml diretamente do XLSX.
    """

    try:

        with ZipFile(
            file_path,
            "r",
        ) as archive:

            try:
                xml_bytes = archive.read(
                    "docProps/core.xml"
                )

            except KeyError:
                return None

        xml = xml_bytes.decode(
            "utf-8",
            errors="ignore",
        )

        match = re.search(
            r"<dcterms:created[^>]*>([^<]+)</dcterms:created>",
            xml,
        )

        if not match:
            return None

        value = match.group(1)

        value = value.replace(
            "Z",
            "+00:00",
        )

        try:

            parsed = datetime.fromisoformat(
                value
            )

            if parsed.tzinfo is not None:

                parsed = (
                    parsed
                    .astimezone()
                    .replace(tzinfo=None)
                )

            return parsed

        except ValueError:
            return None

    except Exception:
        return None


# ============================================================
# FUNÇÕES AUXILIARES DO PROCESSAMENTO
# ============================================================

def same_calendar_day(
    a: datetime,
    b: datetime,
) -> bool:

    return (
        a.year == b.year
        and a.month == b.month
        and a.day == b.day
    )


def subtract_two_years(
    value: datetime,
) -> datetime:

    try:

        return value.replace(
            year=value.year - 2
        )

    except ValueError:

        # Segurança para 29/02.
        return value.replace(
            year=value.year - 2,
            day=28,
        )


def get_value(
    row,
    index,
):

    if index < 0 or index >= len(row):
        return None

    return row[index]


def find_columns(header):

    normalized = [
        normalize_text(value)
        for value in header
    ]

    def find(*names):

        for name in names:

            target = normalize_text(name)

            if target in normalized:

                return normalized.index(
                    target
                )

        return -1

    return {

        # O procedimento manual usa COD_SIT.
        # DSC_SIT fica como fallback.
        "cod_sit": find("COD_SIT"),

        "dsc_sit": find("DSC_SIT"),

        "contrato":
            find("COD_CONTRATO"),

        "ent_solta":
            find("ENTIDADE_SOLTA"),

        "ent_assoc":
            find("ENTIDADE_ASSOCIADA"),

        "info_act":
            find("INFO_ACT"),

        "dt_act":
            find("DT_ACT"),

        "emp_part":
            find("DSC_EMP_PART"),

        "doc_valido":
            find("DOCUMENTO_VALIDO"),
    }


def is_encerrada(
    row,
    columns,
) -> bool:

    cod_sit = norm_val(
        get_value(
            row,
            columns["cod_sit"],
        )
    )

    dsc_sit = norm_val(
        get_value(
            row,
            columns["dsc_sit"],
        )
    )

    # Regra manual:
    #
    # COD_SIT = E
    #
    # DSC_SIT é apenas fallback para ficheiros
    # que não tenham COD_SIT.

    return (
        cod_sit == "E"
        or (
            columns["cod_sit"] < 0
            and dsc_sit == "ENCERRADA"
        )
    )


def row_is_fiabilizada(
    row,
    columns,
    inicio_janela: datetime,
    ref_date: datetime,
) -> bool:

    # ========================================================
    # FILTROS BASE
    # ========================================================

    # COD_SIT != E
    if is_encerrada(
        row,
        columns,
    ):
        return False

    # ENTIDADE_SOLTA = NAO
    if norm_val(
        get_value(
            row,
            columns["ent_solta"],
        )
    ) != "NAO":

        return False

    # ========================================================
    # INFO_ACT = SIM
    # ========================================================

    if norm_val(
        get_value(
            row,
            columns["info_act"],
        )
    ) != "SIM":

        return False

    # ========================================================
    # DT_ACT NOS ÚLTIMOS DOIS ANOS
    # ========================================================

    dt = parse_dt_act(
        get_value(
            row,
            columns["dt_act"],
        ),
        ref_date,
    )

    if dt is None:
        return False

    # O limite inferior da janela é exclusivo.
    # Ex.: com ref_date = 03/08/2026, 03/08/2024 não recebe X.
    if dt < inicio_janela:
        return False

    if dt > ref_date:
        return False

    # ========================================================
    # EMPRESA
    # ========================================================

    emp_part = norm_val(
        get_value(
            row,
            columns["emp_part"],
        )
    )

    if emp_part == "EMPRESA":
        return True

    # ========================================================
    # PARTICULAR
    # ========================================================

    if emp_part == "PARTICULAR":

        doc_valido = norm_val(
            get_value(
                row,
                columns["doc_valido"],
            )
        )

        return doc_valido in {
            "SIM",
            "SEM DATA DE VALIDADE",
        }

    return False


# ============================================================
# EACT PROCESSING
# ============================================================

def process_eact(
    file_path: str,
):
    """
    Processa o EACT em duas passagens.

    PASSAGEM 1:
        - totalBruto
        - encerradas
        - entidadesSoltas
        - totalEntidades
        - maior DT_ACT
          (APENAS entre as linhas com INFO_ACT = SIM,
           depois dos filtros COD_SIT != E e ENTIDADE_SOLTA = NAO)

    PASSAGEM 2:
        reproduz a coluna manual Fiabilizados.

    Cada linha é analisada individualmente.

    Se uma linha satisfizer os filtros,
    ela receberia X no Excel.

    No final:

        COUNT(DISTINCT ENTIDADE_ASSOCIADA)

    entre as linhas que receberiam X.
    """

    created_date = get_created_date(
        file_path
    )

    # ========================================================
    # PASSAGEM 1
    # ========================================================

    workbook = load_workbook(
        filename=file_path,
        read_only=True,
        data_only=True,
    )

    try:

        worksheet = None

        for ws in workbook.worksheets:

            if (
                "export"
                in normalize_text(
                    ws.title
                )
            ):

                worksheet = ws
                break

        if worksheet is None:

            raise ValueError(
                "Não foi encontrada a folha "
                "EXPORT no ficheiro EACT."
            )

        rows = worksheet.iter_rows(
            values_only=True
        )

        try:

            header_row = next(rows)

        except StopIteration:

            raise ValueError(
                "A folha EXPORT está vazia."
            )

        columns = find_columns(
            header_row
        )

        # ====================================================
        # COLUNAS OBRIGATÓRIAS
        # ====================================================

        required = {

            "ENTIDADE_SOLTA":
                columns["ent_solta"],

            "ENTIDADE_ASSOCIADA":
                columns["ent_assoc"],

            "INFO_ACT":
                columns["info_act"],

            "DT_ACT":
                columns["dt_act"],

            "DSC_EMP_PART":
                columns["emp_part"],

            "DOCUMENTO_VALIDO":
                columns["doc_valido"],
        }

        missing = [
            name
            for name, index
            in required.items()
            if index == -1
        ]

        if (
            columns["cod_sit"] < 0
            and columns["dsc_sit"] < 0
        ):

            missing.append(
                "COD_SIT/DSC_SIT"
            )

        if columns["contrato"] < 0:

            missing.append(
                "COD_CONTRATO"
            )

        if missing:

            raise ValueError(
                "Colunas obrigatórias "
                "em falta: "
                + ", ".join(
                    dict.fromkeys(
                        missing
                    )
                )
            )

        # ====================================================
        # CONTADORES
        # ====================================================

        total_bruto = 0

        vistos_enc = set()
        encerradas = 0

        vistos_solta = set()
        entidades_solta_count = 0

        total_entidades_set = set()

        # IMPORTANTE:
        # A maior DT_ACT usada para determinar a data de
        # referência é calculada SOMENTE depois de INFO_ACT
        # = SIM. Portanto, DT_ACT de uma linha INFO_ACT != SIM
        # nunca participa nesta seleção.
        max_dt_act = None

        consecutive_empty = 0

        # ====================================================
        # LEITURA DAS LINHAS
        # ====================================================

        for row in rows:

            if all(
                value is None
                for value in row
            ):

                consecutive_empty += 1

                if (
                    consecutive_empty
                    > 2000
                ):

                    break

                continue

            consecutive_empty = 0

            total_bruto += 1

            # ==================================================
            # ENCERRADAS
            # ==================================================

            if is_encerrada(
                row,
                columns,
            ):

                key = get_value(
                    row,
                    columns["contrato"],
                )

                if key not in vistos_enc:

                    vistos_enc.add(key)

                    encerradas += 1

            # ==================================================
            # ENTIDADES SOLTAS
            # ==================================================

            ent_solta = norm_val(
                get_value(
                    row,
                    columns["ent_solta"],
                )
            )

            if ent_solta == "SIM":

                key = get_value(
                    row,
                    columns["ent_assoc"],
                )

                if key not in vistos_solta:

                    vistos_solta.add(key)

                    entidades_solta_count += 1

            # ==================================================
            # UNIVERSO
            # ============================
            
            if is_encerrada(
                row,
                columns,
            ):

                continue

            if ent_solta != "NAO":

                continue

            total_entidades_set.add(
                get_value(
                    row,
                    columns["ent_assoc"],
                )
            )

            # ==================================================
            # INFO_ACT = SIM
            # ==================================================

            # Só depois de seleccionar INFO_ACT = SIM é que
            # a DT_ACT desta linha pode participar na seleção
            # da última DT_ACT / data de referência.
            if norm_val(
                get_value(
                    row,
                    columns["info_act"],
                )
            ) != "SIM":

                continue

            # ==================================================
            # MAIOR DT_ACT DO SUBCONJUNTO INFO_ACT = SIM
            # ==================================================

            dt = parse_dt_act(
                get_value(
                    row,
                    columns["dt_act"],
                ),
                created_date
                or datetime.now(),
            )

            if dt is not None:

                if (
                    max_dt_act is None
                    or dt > max_dt_act
                ):

                    max_dt_act = dt

        total_entidades = len(
            total_entidades_set
        )

    finally:

        workbook.close()

    # ========================================================
    # DATA DE REFERÊNCIA
    # ========================================================

    # Regra da data de referência:
    #
    # 1. Se a última DT_ACT for igual à data de criação
    #    do ficheiro, usa-se o dia anterior à criação.
    #
    # 2. Se a última DT_ACT for anterior à criação,
    #    usa-se a própria última DT_ACT.
    #
    # 3. Se a última DT_ACT for posterior à criação,
    #    usa-se o dia anterior à criação.
    #
    # Assim, uma DT_ACT futura em relação ao ficheiro não
    # desloca artificialmente a janela de fiabilização.

    ref_date = None

    if max_dt_act is not None:

        if created_date is None:

            # Sem data de criação, não há comparação possível.
            # Mantém-se a última DT_ACT como referência.
            ref_date = datetime(
                max_dt_act.year,
                max_dt_act.month,
                max_dt_act.day,
            )

        else:

            max_dt_act_day = datetime(
                max_dt_act.year,
                max_dt_act.month,
                max_dt_act.day,
            )

            created_date_day = datetime(
                created_date.year,
                created_date.month,
                created_date.day,
            )

            if same_calendar_day(
                max_dt_act_day,
                created_date_day,
            ):

                # Última DT_ACT = data de criação.
                ref_date = created_date_day - timedelta(
                    days=1
                )

            elif max_dt_act_day < created_date_day:

                # Última DT_ACT anterior à criação.
                ref_date = max_dt_act_day

            else:

                # Última DT_ACT posterior à criação.
                # Usa-se o dia anterior à criação.
                ref_date = created_date_day - timedelta(
                    days=1
                )

    if ref_date is not None:

        inicio_janela = (
            subtract_two_years(
                ref_date
            )
        )

    else:

        inicio_janela = None

    # ========================================================
    # PASSAGEM 2
    # ========================================================

    fiabilizadas_set = set()

    if (
        ref_date is not None
        and inicio_janela is not None
    ):

        workbook = load_workbook(
            filename=file_path,
            read_only=True,
            data_only=True,
        )

        try:

            worksheet = None

            for ws in workbook.worksheets:

                if (
                    "export"
                    in normalize_text(
                        ws.title
                    )
                ):

                    worksheet = ws
                    break

            if worksheet is None:

                raise ValueError(
                    "Não foi encontrada a "
                    "folha EXPORT."
                )

            rows = worksheet.iter_rows(
                values_only=True
            )

            try:

                header_row = next(rows)

            except StopIteration:

                raise ValueError(
                    "A folha EXPORT está vazia."
                )

            columns = find_columns(
                header_row
            )

            consecutive_empty = 0

            for row in rows:

                if all(
                    value is None
                    for value in row
                ):

                    consecutive_empty += 1

                    if (
                        consecutive_empty
                        > 2000
                    ):

                        break

                    continue

                consecutive_empty = 0

                # ==============================================
                # ESTA LINHA RECEBERIA X?
                # ==============================================

                if not row_is_fiabilizada(
                    row,
                    columns,
                    inicio_janela,
                    ref_date,
                ):

                    continue

                # ==============================================
                # SIM -> X
                # ==============================================

                entidade = get_value(
                    row,
                    columns["ent_assoc"],
                )

                fiabilizadas_set.add(
                    entidade
                )

        finally:

            workbook.close()

    # ========================================================
    # RESULTADO FINAL
    # ========================================================

    fiabilizadas = len(
        fiabilizadas_set
    )

    por_fiabilizar = (
        total_entidades
        - fiabilizadas
    )

    return {
        "totalBruto": total_bruto,

        "encerradas":
            encerradas,

        "entidadesSoltas":
            entidades_solta_count,

        "totalEntidades":
            total_entidades,

        "fiabilizadas":
            fiabilizadas,

        "porFiabilizar":
            por_fiabilizar,
    }


# ============================================================
# PROCESSAMENTO EM BACKGROUND
# ============================================================

async def run_processing(
    job_id: str,
    temp_path: str,
):

    loop = asyncio.get_event_loop()

    try:

        result = await loop.run_in_executor(
            executor,
            process_eact,
            temp_path,
        )

        jobs[job_id] = {
            "status": "done",
            "result": result,
        }

    except Exception as exc:

        jobs[job_id] = {
            "status": "error",
            "error": (
                f"{type(exc).__name__}: {exc}"
            ),
        }

    finally:

        try:

            Path(
                temp_path
            ).unlink(
                missing_ok=True
            )

        except Exception:
            pass


# ============================================================
# ROUTES
# ============================================================

@app.get("/")
async def root():

    return {
        "status": "ok",
        "service": "EACT FastAPI backend",
    }


@app.post("/api/eact/upload-chunk")
async def upload_chunk(
    file: UploadFile = File(...),
    uploadId: str = Form(...),
    chunkIndex: int = Form(...),
    totalChunks: int = Form(...),
):
    """
    Recebe UM pedaço do ficheiro de cada vez.

    Cada pedaço é escrito no ficheiro temporário.
    Quando o último pedaço chega, o processamento
    é iniciado em background.
    """

    if uploadId not in uploads:

        temp = NamedTemporaryFile(
            suffix=".xlsx",
            prefix="eact-",
            delete=False,
        )

        temp_path = temp.name

        temp.close()

        uploads[uploadId] = {
            "path": temp_path,
            "received": 0,
            "total": totalChunks,
        }

    info = uploads[uploadId]

    chunk_bytes = await file.read()

    with open(
        info["path"],
        "ab",
    ) as f:

        f.write(chunk_bytes)

    info["received"] += 1

    is_last = (
        info["received"]
        >= info["total"]
    )

    if is_last:

        job_id = str(
            uuid.uuid4()
        )

        jobs[job_id] = {
            "status": "processing"
        }

        temp_path = info["path"]

        del uploads[uploadId]

        asyncio.create_task(
            run_processing(
                job_id,
                temp_path,
            )
        )

        return {
            "done": True,
            "job_id": job_id,
        }

    return {
        "done": False,
        "received": info["received"],
        "total": info["total"],
    }


@app.get("/api/eact/status/{job_id}")
async def eact_status(
    job_id: str,
):

    job = jobs.get(job_id)

    if job is None:

        raise HTTPException(
            status_code=404,
            detail="Job não encontrado.",
        )



    return job
