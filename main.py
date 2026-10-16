

# ##########################################################################
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
# from typing import Any, Optional
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
#     reference_date: datetime,
# ) -> Optional[datetime]:
#     if value is None:
#         return None

#     if isinstance(value, datetime):
#         return value.replace(tzinfo=None)

#     if isinstance(value, date):
#         return datetime(
#             value.year,
#             value.month,
#             value.day,
#         )

#     text = str(value).strip()

#     if not text:
#         return None

#     # Formato: YYYY-MM-DD HH:MM:SS
#     # Exemplo: 2025-03-29 00:00:00
#     for fmt in (
#         "%Y-%m-%d %H:%M:%S",
#         "%Y-%m-%d",
#         "%Y.%m.%d",
#     ):
#         try:
#             return datetime.strptime(text, fmt)
#         except ValueError:
#             pass

#     # Formato: YY-MM-DD
#     # Exemplo: 25-03-29
#     for fmt in (
#         "%y-%m-%d",
#         "%y.%m.%d",
#     ):
#         try:
#             parsed = datetime.strptime(text, fmt)

#             # Mantém a mesma lógica do formato YY.MM.DD:
#             # inicialmente considera 20xx.
#             if parsed > reference_date:
#                 parsed = parsed.replace(
#                     year=parsed.year - 100
#                 )

#             return parsed

#         except ValueError:
#             pass

#     return None


# def corrigir_ano_futuro(
#     dt: datetime,
#     reference_date: datetime,
# ) -> datetime:
#     """
#     DT_ACT com ano superior ao atual (erro de digitação,
#     ex.: 2029 em vez de 2026) passa para o ano atual,
#     mantendo o dia e o mês.

#     O ano atual é o da data de referência (ref_date) do ficheiro.
#     Datas que fiquem depois de ref_date continuam a ser
#     descartadas pelo filtro da janela.
#     """

#     if dt.year <= reference_date.year:
#         return dt

#     try:
#         return dt.replace(
#             year=reference_date.year
#         )

#     except ValueError:

#         # 29/02 num ano não bissexto
#         return dt.replace(
#             year=reference_date.year,
#             day=28,
#         )


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

#         # Usadas apenas nos passos 1 a 10.4
#         # (exclusão por conta / por cliente).
#         "conta":
#             find("CONTA"),

#         "titularidade":
#             find("DSC_TIPO_TITULARIDADE"),

#         "sit_credito":
#             find("SITUACAO_CREDITO"),
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

#     # Ano superior ao atual (ex.: 2029) -> ano atual,
#     # mantendo dia e mês.
#     dt = corrigir_ano_futuro(
#         dt,
#         ref_date,
#     )

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
# # PASSOS 1 A 10.4 (EXCLUSÃO POR CONTA / POR CLIENTE E INDICADORES)
# # ============================================================

# def norm_codigo(value) -> str:
#     """
#     Normaliza códigos (COD_SIT, SITUACAO_CREDITO).

#     Garante que 9, 9.0 e "9" são tratados como "9".
#     """

#     if value is None:
#         return ""

#     if isinstance(value, float) and value.is_integer():
#         value = int(value)

#     return norm_val(value)


# def exclusao_por_conta(
#     cod_sit: str,
#     sit_credito: str,
# ) -> Optional[str]:
#     """
#     PASSO 1 - Exclusão por Conta (regra por linha).

#     Ordem:
#         COD_SIT = 9                       -> Sit 9
#         COD_SIT = 5                       -> Sit 5
#         COD_SIT = I                       -> Ina
#         em branco e SITUACAO_CREDITO 9/5  -> Cred
#     """

#     if cod_sit == "9":
#         return "Sit 9"

#     if cod_sit == "5":
#         return "Sit 5"

#     if cod_sit == "I":
#         return "Ina"

#     if sit_credito in ("9", "5"):
#         return "Cred"

#     return None


# def calcular_indicadores_entidades(
#     registos: list,
# ) -> dict:
#     """
#     Percorre os passos 1 a 10.4 sobre o universo
#     (COD_SIT != E e ENTIDADE_SOLTA = NAO), na ordem do ficheiro.

#     Cada registo:
#         (conta, e_titular_1, entidade, cod_sit, sit_credito, fiabilizada)

#     PASSO 1  Exclusão por Conta, por linha. Os fiabilizados (X)
#              NÃO são apagados nesta fase.
#     PASSO 2  Matriz de Titulares 01 (1.ª linha de cada conta, com os
#              X ainda lá) e PROCV: todos os titulares herdam o estado
#              do Titular 01 da conta.
#     PASSO 3  Entidades com pelo menos um contrato COD_SIT = N.
#     PASSO 4  Apaga o Ina das entidades com conta normal
#              (só o Ina; Sit 9, Sit 5 e Cred ficam).
#     PASSO 5  Apaga a Exclusão por Cliente das linhas fiabilizadas (X).
#     PASSO 6  Titulares 01, sem duplicados por entidade (1.ª linha).
#     PASSO 7  Restantes titulares, sem duplicados por entidade (1.ª linha).
#     PASSO 8  Junta 6 + 7 (Titulares 01 primeiro) e remove duplicados
#              por entidade: uma linha por entidade.
#     PASSO 9  Por fiabilizar = entidades sem X.
#     PASSO 10 Fiabilizadas, por fiabilizar prioritárias, Eact, excluídas.
#     """

#     n = len(registos)

#     # PASSO 1
#     estado = [
#         exclusao_por_conta(r[3], r[4])
#         for r in registos
#     ]

#     # PASSO 2
#     matriz_t1: dict = {}

#     for r, est in zip(registos, estado):

#         if r[1] and r[0] not in matriz_t1:
#             matriz_t1[r[0]] = est

#     cliente = [
#         matriz_t1.get(r[0])
#         for r in registos
#     ]

#     linhas_passo2 = sum(
#         1 for v in cliente if v is not None
#     )

#     # PASSO 3
#     entidades_normais = {
#         r[2]
#         for r in registos
#         if r[3] == "N"
#     }

#     # PASSO 4
#     for i, r in enumerate(registos):

#         if (
#             cliente[i] == "Ina"
#             and r[2] in entidades_normais
#         ):
#             cliente[i] = None

#     linhas_passo4 = sum(
#         1 for v in cliente if v is not None
#     )

#     # PASSO 5
#     for i, r in enumerate(registos):

#         if r[5]:
#             cliente[i] = None

#     linhas_excluidas = sum(
#         1 for v in cliente if v is not None
#     )

#     # PASSOS 6, 7 E 8
#     primeira_t1: dict = {}
#     primeira_outro: dict = {}

#     for i, r in enumerate(registos):

#         destino = (
#             primeira_t1
#             if r[1]
#             else primeira_outro
#         )

#         if r[2] not in destino:
#             destino[r[2]] = i

#     def contar(indices):

#         indices = list(indices)

#         return {
#             "entidades": len(indices),
#             "fiabilizadas": sum(
#                 1 for i in indices if registos[i][5]
#             ),
#             "excluidas": sum(
#                 1 for i in indices if cliente[i] is not None
#             ),
#         }

#     passo6 = contar(primeira_t1.values())
#     passo7 = contar(primeira_outro.values())

#     finais = dict(primeira_outro)
#     finais.update(primeira_t1)   # Titular 01 prevalece

#     passo8 = contar(finais.values())

#     # PASSOS 9 E 10
#     total_entidades = len(finais)

#     fiabilizadas = sum(
#         1 for i in finais.values() if registos[i][5]
#     )

#     por_fiabilizar = total_entidades - fiabilizadas

#     # 10.3 - Fiabilizado vazio e Exclusão por Cliente vazia
#     prioritarias = sum(
#         1
#         for i in finais.values()
#         if not registos[i][5] and cliente[i] is None
#     )

#     # 10.2
#     eact = fiabilizadas + prioritarias

#     # 10.4
#     excluidas = sum(
#         1 for i in finais.values() if cliente[i] is not None
#     )

#     por_valor: dict = {}

#     for i in finais.values():

#         if cliente[i] is not None:
#             por_valor[cliente[i]] = (
#                 por_valor.get(cliente[i], 0) + 1
#             )

#     verificacoes = {
#         # Total = fiabilizadas + por fiabilizar
#         "totalIgualFiabilizadasMaisPorFiabilizar":
#             total_entidades == fiabilizadas + por_fiabilizar,
#         # Por fiabilizar = prioritárias + excluídas
#         "porFiabilizarIgualPrioritariasMaisExcluidas":
#             por_fiabilizar == prioritarias + excluidas,
#         # Total = Eact (considerando as prioritárias) + excluídas
#         "totalIgualEactPrioritariasMaisExcluidas":
#             total_entidades == eact + excluidas,
#     }

#     if not all(verificacoes.values()):

#         raise RuntimeError(
#             "Os indicadores finais não são consistentes: "
#             + str(verificacoes)
#         )

#     return {
#         # Total de entidades (passo 8, sem duplicados).
#         # "Entidades Eact" sem mais nada é este total.
#         "totalEntidades": total_entidades,
#         "entidadesEact": total_entidades,

#         # 10.1
#         "fiabilizadas": fiabilizadas,

#         # 9
#         "porFiabilizar": por_fiabilizar,

#         # 10.3 - Por fiabilizar (considerando as prioritárias)
#         "porFiabilizarConsiderandoPrioritarias": prioritarias,

#         # 10.2 - Entidades Eact (considerando as prioritárias)
#         "entidadesEactConsiderandoPrioritarias": eact,

#         # 10.4
#         "excluidas": excluidas,
#         "detalhe": {
#             "linhasUniverso": n,
#             "linhasPasso2": linhas_passo2,
#             "linhasPasso4": linhas_passo4,
#             "linhasExcluidas": linhas_excluidas,
#             "entidadesNormais": len(entidades_normais),
#             "passo6": passo6,
#             "passo7": passo7,
#             "passo8": passo8,
#             "excluidasPorValor": por_valor,
#         },
#         "verificacoes": verificacoes,
#     }


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
#           (APENAS entre as linhas com INFO_ACT = SIM,
#            depois dos filtros COD_SIT != E e ENTIDADE_SOLTA = NAO)

#     PASSAGEM 2:
#         reproduz a coluna manual Fiabilizados.

#     Cada linha é analisada individualmente.

#     Se uma linha satisfizer os filtros,
#     ela receberia X no Excel.

#     No final:

#         COUNT(DISTINCT ENTIDADE_ASSOCIADA)

#     entre as linhas que receberiam X.

#     Na mesma passagem 2 são recolhidos os registos do universo
#     e, no fim, percorridos os passos 1 a 10.4
#     (calcular_indicadores_entidades), que determinam
#     fiabilizadas, por fiabilizar, por fiabilizar prioritárias,
#     Eact e excluídas. A passagem 1 (totalBruto, encerradas,
#     entidadesSoltas, totalEntidades) não foi alterada.
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

#         # IMPORTANTE:
#         # A maior DT_ACT usada para determinar a data de
#         # referência é calculada SOMENTE depois de INFO_ACT
#         # = SIM. Portanto, DT_ACT de uma linha INFO_ACT != SIM
#         # nunca participa nesta seleção.
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
#             # ============================
            
#             if is_encerrada(
#                 row,
#                 columns,
#             ):

#                 continue

#             if ent_solta != "NAO":

#                 continue

#             # ==================================================
#             # INFO_ACT = SIM
#             # ==================================================

#             # Só depois de seleccionar INFO_ACT = SIM é que
#             # a DT_ACT desta linha pode participar na seleção
#             # da última DT_ACT / data de referência.
#             if norm_val(
#                 get_value(
#                     row,
#                     columns["info_act"],
#                 )
#             ) != "SIM":

#                 continue

#             # ==================================================
#             # MAIOR DT_ACT DO SUBCONJUNTO INFO_ACT = SIM
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

#     # Registos do universo (COD_SIT != E e ENTIDADE_SOLTA = NAO),
#     # na ordem do ficheiro. Alimentam os passos 1 a 10.4.
#     registos = []

#     janela_valida = (
#         ref_date is not None
#         and inicio_janela is not None
#     )

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
#                 "Não foi encontrada a "
#                 "folha EXPORT."
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

#         # Os passos 1 a 10.4 precisam destas colunas.
#         # Os indicadores finais são obrigatórios no frontend,
#         # por isso a falta de qualquer uma é um erro.
#         colunas_passos = {
#             "COD_SIT": columns["cod_sit"],
#             "CONTA": columns["conta"],
#             "DSC_TIPO_TITULARIDADE": columns["titularidade"],
#             "SITUACAO_CREDITO": columns["sit_credito"],
#         }

#         em_falta = [
#             nome
#             for nome, indice in colunas_passos.items()
#             if indice < 0
#         ]

#         if em_falta:

#             raise ValueError(
#                 "Faltam colunas na folha EXPORT para "
#                 "calcular os indicadores: "
#                 + ", ".join(em_falta)
#             )

#         consecutive_empty = 0

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

#             # ==============================================
#             # UNIVERSO: COD_SIT != E e ENTIDADE_SOLTA = NAO
#             # ==============================================

#             if is_encerrada(
#                 row,
#                 columns,
#             ):

#                 continue

#             if norm_val(
#                 get_value(
#                     row,
#                     columns["ent_solta"],
#                 )
#             ) != "NAO":

#                 continue

#             entidade = get_value(
#                 row,
#                 columns["ent_assoc"],
#             )

#             # ==============================================
#             # ESTA LINHA RECEBERIA X?
#             # ==============================================

#             fiabilizada = (
#                 janela_valida
#                 and row_is_fiabilizada(
#                     row,
#                     columns,
#                     inicio_janela,
#                     ref_date,
#                 )
#             )

#             if fiabilizada:

#                 fiabilizadas_set.add(
#                     entidade
#                 )

#             registos.append(
#                 (
#                     get_value(
#                         row,
#                         columns["conta"],
#                     ),
#                     norm_val(
#                         get_value(
#                             row,
#                             columns["titularidade"],
#                         )
#                     ) == "TITULAR 01",
#                     entidade,
#                     norm_codigo(
#                         get_value(
#                             row,
#                             columns["cod_sit"],
#                         )
#                     ),
#                     norm_codigo(
#                         get_value(
#                             row,
#                             columns["sit_credito"],
#                         )
#                     ),
#                     fiabilizada,
#                 )
#             )

#     finally:

#         workbook.close()

#     # ========================================================
#     # RESULTADO FINAL
#     # ========================================================

#     # Passos 1 a 10.4: uma linha por entidade
#     # (Titular 01 prevalece), exclusão por cliente
#     # e indicadores finais.
#     indicadores = calcular_indicadores_entidades(
#         registos
#     )

#     resultado = {
#         # Pré-passo 1
#         "totalBruto": total_bruto,

#         "encerradas":
#             encerradas,

#         "entidadesSoltas":
#             entidades_solta_count,

#         # Total de entidades: resultado final do passo 8
#         "totalEntidades":
#             indicadores["totalEntidades"],

#         # "Entidades Eact" = total de entidades
#         "entidadesEact":
#             indicadores["entidadesEact"],

#         # 10.1 Fiabilizadas
#         "fiabilizadas":
#             indicadores["fiabilizadas"],

#         # 9 Por fiabilizar
#         "porFiabilizar":
#             indicadores["porFiabilizar"],

#         # 10.3 Por fiabilizar (considerando as prioritárias)
#         "porFiabilizarConsiderandoPrioritarias":
#             indicadores["porFiabilizarConsiderandoPrioritarias"],

#         # 10.2 Entidades Eact (considerando as prioritárias)
#         "entidadesEactConsiderandoPrioritarias":
#             indicadores["entidadesEactConsiderandoPrioritarias"],

#         # 10.4 Excluídas
#         "excluidas":
#             indicadores["excluidas"],

#         "detalhe":
#             indicadores["detalhe"],

#         "verificacoes":
#             indicadores["verificacoes"],
#     }

#     # Contagem antiga: entidades distintas com ALGUMA linha com X.
#     resultado["detalhe"][
#         "fiabilizadasAlgumaLinha"
#     ] = len(fiabilizadas_set)

#     return resultado


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
from functools import lru_cache
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


def corrigir_ano_futuro(
    dt: datetime,
    reference_date: datetime,
) -> datetime:
    """
    DT_ACT com ano superior ao atual (erro de digitação,
    ex.: 2029 em vez de 2026) passa para o ano atual,
    mantendo o dia e o mês.

    O ano atual é o da data de referência (ref_date) do ficheiro.
    Datas que fiquem depois de ref_date continuam a ser
    descartadas pelo filtro da janela.
    """

    if dt.year <= reference_date.year:
        return dt

    try:
        return dt.replace(
            year=reference_date.year
        )

    except ValueError:

        # 29/02 num ano não bissexto
        return dt.replace(
            year=reference_date.year,
            day=28,
        )


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

        # Usadas apenas nos passos 1 a 10.4
        # (exclusão por conta / por cliente).
        "conta":
            find("CONTA"),

        "titularidade":
            find("DSC_TIPO_TITULARIDADE"),

        "sit_credito":
            find("SITUACAO_CREDITO"),

        # Usadas apenas no passo 11 (entidades sem NUIT).
        "obs":
            find("OBS"),

        "dsc_seg":
            find("DSC_SEG"),

        "nuit":
            find("NUIT"),

        "nome":
            find("NOME"),
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

    # Ano superior ao atual (ex.: 2029) -> ano atual,
    # mantendo dia e mês.
    dt = corrigir_ano_futuro(
        dt,
        ref_date,
    )

    # O limite inferior da janela é exclusivo.
    # Ex.: com ref_date = 03/08/2026, 03/08/2024 não recebe X.
    if dt <= inicio_janela:
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
# PASSOS 1 A 10.4 (EXCLUSÃO POR CONTA / POR CLIENTE E INDICADORES)
# ============================================================

def norm_codigo(value) -> str:
    """
    Normaliza códigos (COD_SIT, SITUACAO_CREDITO).

    Garante que 9, 9.0 e "9" são tratados como "9".
    """

    if value is None:
        return ""

    if isinstance(value, float) and value.is_integer():
        value = int(value)

    return norm_val(value)


# ------------------------------------------------------------
# PASSO 11 - ENTIDADES SEM NUIT
# Os textos de OBS e DSC_SEG são comparados sem maiúsculas
# nem acentos (os valores reais do ficheiro trazem
# "Incorecto" com um só "r").
# ------------------------------------------------------------

OBS_BOM = "bom"

OBS_NUIT_NAO_PREENCHIDO = "nuit nao preenchido"

OBS_ERRO_CHECK_DIGIT = "erro de check digit"

OBS_DIGITOS_INCORRECTOS = (
    "contem numero incorecto de digitos",
    "contem numero incorrecto de digitos",
)

SEG_CONTA_INTERNA = "conta interna"

CAT_DIGITOS_INCORRECTOS = "Número incorrecto de dígitos"

CAT_NUIT_NAO_PREENCHIDO = "NUIT não preenchido"

CAT_CONTA_TECNICA = "Conta técnica (NUIT não preenchido)"

CAT_OUTROS = "Outros"


@lru_cache(maxsize=4096)
def _normalize_cached(text: str) -> str:
    return normalize_text(text)


def norm_texto(value) -> str:
    """normalize_text com cache (a coluna OBS repete muito)."""

    if value is None:
        return ""

    return _normalize_cached(str(value))


def valor_simples(value):
    """Valor seguro para JSON/Excel."""

    if value is None:
        return ""

    if isinstance(value, bool):
        return value

    if isinstance(value, float) and value.is_integer():
        return int(value)

    if isinstance(value, (int, float, str)):
        return value

    return str(value)


def categoria_sem_nuit(obs: str, seg: str) -> str:
    """
    Classifica uma entidade sem NUIT (OBS diferente de BOM).

    2 - OBS = Contém Número Incorecto de dígitos
    3 - OBS = Nuit não preenchido ou ERRO DE CHECK DIGIT,
        e DSC_SEG diferente de Conta Interna
    4 - OBS = Nuit não preenchido e DSC_SEG = Conta Interna
        (conta técnica)
    """

    if obs in OBS_DIGITOS_INCORRECTOS:
        return CAT_DIGITOS_INCORRECTOS

    if obs in (
        OBS_NUIT_NAO_PREENCHIDO,
        OBS_ERRO_CHECK_DIGIT,
    ):

        if seg != SEG_CONTA_INTERNA:
            return CAT_NUIT_NAO_PREENCHIDO

        if obs == OBS_NUIT_NAO_PREENCHIDO:
            return CAT_CONTA_TECNICA

    return CAT_OUTROS


def exclusao_por_conta(
    cod_sit: str,
    sit_credito: str,
) -> Optional[str]:
    """
    PASSO 1 - Exclusão por Conta (regra por linha).

    Ordem:
        COD_SIT = 9                       -> Sit 9
        COD_SIT = 5                       -> Sit 5
        COD_SIT = I                       -> Ina
        em branco e SITUACAO_CREDITO 9/5  -> Cred
    """

    if cod_sit == "9":
        return "Sit 9"

    if cod_sit == "5":
        return "Sit 5"

    if cod_sit == "I":
        return "Ina"

    if sit_credito in ("9", "5"):
        return "Cred"

    return None


def calcular_indicadores_entidades(
    registos: list,
) -> dict:
    """
    Percorre os passos 1 a 10.4 sobre o universo
    (COD_SIT != E e ENTIDADE_SOLTA = NAO), na ordem do ficheiro.

    Cada registo:
        (conta, e_titular_1, entidade, cod_sit, sit_credito, fiabilizada,
         obs_normalizada, dsc_seg_normalizada, extra)

    "extra" só existe quando OBS != BOM:
        (nome, nuit, obs, dsc_seg, titularidade)

    PASSO 1  Exclusão por Conta, por linha. Os fiabilizados (X)
             NÃO são apagados nesta fase.
    PASSO 2  Matriz de Titulares 01 (1.ª linha de cada conta, com os
             X ainda lá) e PROCV: todos os titulares herdam o estado
             do Titular 01 da conta.
    PASSO 3  Entidades com pelo menos um contrato COD_SIT = N.
    PASSO 4  Apaga o Ina das entidades com conta normal
             (só o Ina; Sit 9, Sit 5 e Cred ficam).
    PASSO 5  Apaga a Exclusão por Cliente das linhas fiabilizadas (X).
    PASSO 6  Titulares 01, sem duplicados por entidade (1.ª linha).
    PASSO 7  Restantes titulares, sem duplicados por entidade (1.ª linha).
    PASSO 8  Junta 6 + 7 (Titulares 01 primeiro) e remove duplicados
             por entidade: uma linha por entidade.
    PASSO 9  Por fiabilizar = entidades sem X.
    PASSO 10 Fiabilizadas, por fiabilizar prioritárias, Eact, excluídas.
    PASSO 11 Entidades sem NUIT (sobre a folha do passo 8):
             OBS diferente de BOM; número incorrecto de dígitos;
             NUIT não preenchido (sem Conta Interna);
             contas técnicas (Conta Interna).
    """

    n = len(registos)

    # PASSO 1
    estado = [
        exclusao_por_conta(r[3], r[4])
        for r in registos
    ]

    # PASSO 2
    matriz_t1: dict = {}

    for r, est in zip(registos, estado):

        if r[1] and r[0] not in matriz_t1:
            matriz_t1[r[0]] = est

    cliente = [
        matriz_t1.get(r[0])
        for r in registos
    ]

    linhas_passo2 = sum(
        1 for v in cliente if v is not None
    )

    # PASSO 3
    entidades_normais = {
        r[2]
        for r in registos
        if r[3] == "N"
    }

    # PASSO 4
    for i, r in enumerate(registos):

        if (
            cliente[i] == "Ina"
            and r[2] in entidades_normais
        ):
            cliente[i] = None

    linhas_passo4 = sum(
        1 for v in cliente if v is not None
    )

    # PASSO 5
    for i, r in enumerate(registos):

        if r[5]:
            cliente[i] = None

    linhas_excluidas = sum(
        1 for v in cliente if v is not None
    )

    # PASSOS 6, 7 E 8
    primeira_t1: dict = {}
    primeira_outro: dict = {}

    for i, r in enumerate(registos):

        destino = (
            primeira_t1
            if r[1]
            else primeira_outro
        )

        if r[2] not in destino:
            destino[r[2]] = i

    def contar(indices):

        indices = list(indices)

        return {
            "entidades": len(indices),
            "fiabilizadas": sum(
                1 for i in indices if registos[i][5]
            ),
            "excluidas": sum(
                1 for i in indices if cliente[i] is not None
            ),
        }

    passo6 = contar(primeira_t1.values())
    passo7 = contar(primeira_outro.values())

    finais = dict(primeira_outro)
    finais.update(primeira_t1)   # Titular 01 prevalece

    passo8 = contar(finais.values())

    # PASSOS 9 E 10
    total_entidades = len(finais)

    fiabilizadas = sum(
        1 for i in finais.values() if registos[i][5]
    )

    por_fiabilizar = total_entidades - fiabilizadas

    # 10.3 - Fiabilizado vazio e Exclusão por Cliente vazia
    prioritarias = sum(
        1
        for i in finais.values()
        if not registos[i][5] and cliente[i] is None
    )

    # 10.2
    eact = fiabilizadas + prioritarias

    # 10.4
    excluidas = sum(
        1 for i in finais.values() if cliente[i] is not None
    )

    # PASSO 11 - entidades sem NUIT (linha que ficou no passo 8)
    sem_nuit_lista = []

    n_incorrecto = 0
    n_nao_preenchido = 0
    n_tecnicas = 0
    n_outros = 0

    for i in sorted(finais.values()):

        r = registos[i]

        if r[6] == OBS_BOM:
            continue

        categoria = categoria_sem_nuit(r[6], r[7])

        if categoria == CAT_DIGITOS_INCORRECTOS:
            n_incorrecto += 1

        elif categoria == CAT_NUIT_NAO_PREENCHIDO:
            n_nao_preenchido += 1

        elif categoria == CAT_CONTA_TECNICA:
            n_tecnicas += 1

        else:
            n_outros += 1

        extra = r[8] or (None, None, None, None, None)

        sem_nuit_lista.append(
            {
                "entidade": valor_simples(r[2]),
                "nome": valor_simples(extra[0]),
                "nuit": valor_simples(extra[1]),
                "obs": valor_simples(extra[2]),
                "categoria": categoria,
                "segmento": valor_simples(extra[3]),
                "titularidade": valor_simples(extra[4]),
                "conta": valor_simples(r[0]),
                "fiabilizada": "X" if r[5] else "",
                "exclusaoPorCliente": cliente[i] or "",
            }
        )

    por_valor: dict = {}

    for i in finais.values():

        if cliente[i] is not None:
            por_valor[cliente[i]] = (
                por_valor.get(cliente[i], 0) + 1
            )

    verificacoes = {
        # Total = fiabilizadas + por fiabilizar
        "totalIgualFiabilizadasMaisPorFiabilizar":
            total_entidades == fiabilizadas + por_fiabilizar,
        # Por fiabilizar = prioritárias + excluídas
        "porFiabilizarIgualPrioritariasMaisExcluidas":
            por_fiabilizar == prioritarias + excluidas,
        # Total = Eact (considerando as prioritárias) + excluídas
        "totalIgualEactPrioritariasMaisExcluidas":
            total_entidades == eact + excluidas,
    }

    if not all(verificacoes.values()):

        raise RuntimeError(
            "Os indicadores finais não são consistentes: "
            + str(verificacoes)
        )

    return {
        # Total de entidades (passo 8, sem duplicados).
        # "Entidades Eact" sem mais nada é este total.
        "totalEntidades": total_entidades,
        "entidadesEact": total_entidades,

        # 10.1
        "fiabilizadas": fiabilizadas,

        # 9
        "porFiabilizar": por_fiabilizar,

        # 10.3 - Por fiabilizar (considerando as prioritárias)
        "porFiabilizarConsiderandoPrioritarias": prioritarias,

        # 10.2 - Entidades Eact (considerando as prioritárias)
        "entidadesEactConsiderandoPrioritarias": eact,

        # 10.4
        "excluidas": excluidas,

        # 11.1 - Entidades sem NUIT (OBS diferente de BOM)
        "entidadesSemNuit": len(sem_nuit_lista),

        # 11.2 - Com número incorrecto de dígitos
        "semNuitNumeroIncorrectoDigitos": n_incorrecto,

        # 11.3 - NUIT não preenchido (inclui ERRO DE CHECK DIGIT,
        #        sem Conta Interna)
        "semNuitNaoPreenchido": n_nao_preenchido,

        # 11.4 - Contas técnicas sem NUIT (Conta Interna)
        "contasTecnicasSemNuit": n_tecnicas,

        # Lista completa (para o Excel do dashboard)
        "entidadesSemNuitLista": sem_nuit_lista,
        "detalhe": {
            "linhasUniverso": n,
            "linhasPasso2": linhas_passo2,
            "linhasPasso4": linhas_passo4,
            "linhasExcluidas": linhas_excluidas,
            "entidadesNormais": len(entidades_normais),
            "passo6": passo6,
            "passo7": passo7,
            "passo8": passo8,
            "excluidasPorValor": por_valor,
            "semNuitOutros": n_outros,
        },
        "verificacoes": dict(
            verificacoes,
            # Informativa: 11.1 = 11.2 + 11.3 + 11.4 (sem "outros")
            semNuitSemOutros=(n_outros == 0),
        ),
    }


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

    Na mesma passagem 2 são recolhidos os registos do universo
    e, no fim, percorridos os passos 1 a 10.4
    (calcular_indicadores_entidades), que determinam
    fiabilizadas, por fiabilizar, por fiabilizar prioritárias,
    Eact e excluídas. A passagem 1 (totalBruto, encerradas,
    entidadesSoltas, totalEntidades) não foi alterada.
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

    # Registos do universo (COD_SIT != E e ENTIDADE_SOLTA = NAO),
    # na ordem do ficheiro. Alimentam os passos 1 a 10.4.
    registos = []

    janela_valida = (
        ref_date is not None
        and inicio_janela is not None
    )

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

        # Os passos 1 a 10.4 precisam destas colunas.
        # Os indicadores finais são obrigatórios no frontend,
        # por isso a falta de qualquer uma é um erro.
        colunas_passos = {
            "COD_SIT": columns["cod_sit"],
            "CONTA": columns["conta"],
            "DSC_TIPO_TITULARIDADE": columns["titularidade"],
            "SITUACAO_CREDITO": columns["sit_credito"],
            "OBS": columns["obs"],
            "DSC_SEG": columns["dsc_seg"],
        }

        em_falta = [
            nome
            for nome, indice in colunas_passos.items()
            if indice < 0
        ]

        if em_falta:

            raise ValueError(
                "Faltam colunas na folha EXPORT para "
                "calcular os indicadores: "
                + ", ".join(em_falta)
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
            # UNIVERSO: COD_SIT != E e ENTIDADE_SOLTA = NAO
            # ==============================================

            if is_encerrada(
                row,
                columns,
            ):

                continue

            if norm_val(
                get_value(
                    row,
                    columns["ent_solta"],
                )
            ) != "NAO":

                continue

            entidade = get_value(
                row,
                columns["ent_assoc"],
            )

            # ==============================================
            # ESTA LINHA RECEBERIA X?
            # ==============================================

            fiabilizada = (
                janela_valida
                and row_is_fiabilizada(
                    row,
                    columns,
                    inicio_janela,
                    ref_date,
                )
            )

            if fiabilizada:

                fiabilizadas_set.add(
                    entidade
                )

            # Passo 11: OBS e DSC_SEG (normalizados). Os dados
            # completos só se guardam se a entidade puder ficar
            # sem NUIT (OBS diferente de BOM).
            obs_norm = norm_texto(
                get_value(
                    row,
                    columns["obs"],
                )
            )

            extra_sem_nuit = None

            if obs_norm != OBS_BOM:

                extra_sem_nuit = (
                    get_value(
                        row,
                        columns["nome"],
                    ),
                    get_value(
                        row,
                        columns["nuit"],
                    ),
                    get_value(
                        row,
                        columns["obs"],
                    ),
                    get_value(
                        row,
                        columns["dsc_seg"],
                    ),
                    get_value(
                        row,
                        columns["titularidade"],
                    ),
                )

            registos.append(
                (
                    get_value(
                        row,
                        columns["conta"],
                    ),
                    norm_val(
                        get_value(
                            row,
                            columns["titularidade"],
                        )
                    ) == "TITULAR 01",
                    entidade,
                    norm_codigo(
                        get_value(
                            row,
                            columns["cod_sit"],
                        )
                    ),
                    norm_codigo(
                        get_value(
                            row,
                            columns["sit_credito"],
                        )
                    ),
                    fiabilizada,
                    obs_norm,
                    norm_texto(
                        get_value(
                            row,
                            columns["dsc_seg"],
                        )
                    ),
                    extra_sem_nuit,
                )
            )

    finally:

        workbook.close()

    # ========================================================
    # RESULTADO FINAL
    # ========================================================

    # Passos 1 a 10.4: uma linha por entidade
    # (Titular 01 prevalece), exclusão por cliente
    # e indicadores finais.
    indicadores = calcular_indicadores_entidades(
        registos
    )

    resultado = {
        # Pré-passo 1
        "totalBruto": total_bruto,

        "encerradas":
            encerradas,

        "entidadesSoltas":
            entidades_solta_count,

        # Total de entidades: resultado final do passo 8
        "totalEntidades":
            indicadores["totalEntidades"],

        # "Entidades Eact" = total de entidades
        "entidadesEact":
            indicadores["entidadesEact"],

        # 10.1 Fiabilizadas
        "fiabilizadas":
            indicadores["fiabilizadas"],

        # 9 Por fiabilizar
        "porFiabilizar":
            indicadores["porFiabilizar"],

        # 10.3 Por fiabilizar (considerando as prioritárias)
        "porFiabilizarConsiderandoPrioritarias":
            indicadores["porFiabilizarConsiderandoPrioritarias"],

        # 10.2 Entidades Eact (considerando as prioritárias)
        "entidadesEactConsiderandoPrioritarias":
            indicadores["entidadesEactConsiderandoPrioritarias"],

        # 10.4 Excluídas
        "excluidas":
            indicadores["excluidas"],

        # 11.1 Entidades sem NUIT
        "entidadesSemNuit":
            indicadores["entidadesSemNuit"],

        # 11.2 Com número incorrecto de dígitos
        "semNuitNumeroIncorrectoDigitos":
            indicadores["semNuitNumeroIncorrectoDigitos"],

        # 11.3 NUIT não preenchido (sem Conta Interna)
        "semNuitNaoPreenchido":
            indicadores["semNuitNaoPreenchido"],

        # 11.4 Contas técnicas sem NUIT
        "contasTecnicasSemNuit":
            indicadores["contasTecnicasSemNuit"],

        # Lista completa para o Excel do dashboard
        "entidadesSemNuitLista":
            indicadores["entidadesSemNuitLista"],

        "detalhe":
            indicadores["detalhe"],

        "verificacoes":
            indicadores["verificacoes"],
    }

    # Contagem antiga: entidades distintas com ALGUMA linha com X.
    resultado["detalhe"][
        "fiabilizadasAlgumaLinha"
    ] = len(fiabilizadas_set)

    return resultado


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