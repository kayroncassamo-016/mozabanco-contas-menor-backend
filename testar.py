import time
from main import process_eact

CAMINHO_FICHEIRO = r"E:\EACT_0731.xlsx"

print("A processar... (isto pode demorar uns minutos, é normal)")

inicio = time.time()
resultado = process_eact(CAMINHO_FICHEIRO)
fim = time.time()

print()
print(f"Tempo total: {fim - inicio:.1f} segundos")
print()
print(resultado)