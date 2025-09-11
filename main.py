import sys

sys.setrecursionlimit(sys.getrecursionlimit() * 5)  # Or a higher value if needed

from scripts.modelo_huawei import processar_mod_huawei

if __name__ == "__main__":
    processar_mod_huawei()



