import threading
import os

def bash(numero):
    os.system(f"python pedido{numero}.py")

def executarPedido(funcao,numero):
    t = threading.Thread(target=funcao, args=(numero,))
    t.start()

executarPedido(bash,0)
executarPedido(bash,1)
executarPedido(bash,2)
executarPedido(bash,3)
executarPedido(bash,4)