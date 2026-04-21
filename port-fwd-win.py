import socket
import threading
import sys

def forward(src, dst):
    try:
        while True:
            data = src.recv(65536)
            if not data:
                break
            dst.sendall(data)
    except:
        pass
    finally:
        src.close()
        dst.close()

def start_listener(listen_host, listen_port, target_host, target_port):
    def handle(client):
        try:
            backend = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            backend.connect((target_host, target_port))
            t1 = threading.Thread(target=forward, args=(client, backend), daemon=True)
            t2 = threading.Thread(target=forward, args=(backend, client), daemon=True)
            t1.start()
            t2.start()
        except Exception as e:
            print(f'Connection failed: {e}')
            client.close()

    server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    server.bind((listen_host, listen_port))
    server.listen(128)
    print(f'Forwarding {listen_host}:{listen_port} -> {target_host}:{target_port}')
    sys.stdout.flush()
    while True:
        client, addr = server.accept()
        threading.Thread(target=handle, args=(client,), daemon=True).start()

# 9080 -> 8080 (project: lc4j-agentic-tutorial)
t1 = threading.Thread(target=start_listener, args=('0.0.0.0', 9080, '127.0.0.1', 8080), daemon=True)
t1.start()

# 9081 -> 8081 (obsidian: scienc-project-repo)
t2 = threading.Thread(target=start_listener, args=('0.0.0.0', 9081, '127.0.0.1', 8081), daemon=True)
t2.start()

print('Port forwarding running. 9080->8080, 9081->8081')
sys.stdout.flush()

try:
    while True:
        import time
        time.sleep(3600)
except KeyboardInterrupt:
    print('Shutting down')
