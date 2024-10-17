import socket
import time

class Terminal:
    def __init__(self, ip, port):
        self.ip = ip
        self.port = port
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.sock.connect((self.ip, self.port))
        self.sock.settimeout(1)

    def send(self, message):
        self.sock.send(message.encode())

    def receive(self):
        try:
            data = self.sock.recv(1024)
            return data.decode()
        except socket.timeout:
            return None
        except socket.error:
            return None
    def close(self):
        self.sock.close()
if __name__ == "__main__":
    term = Terminal("127.0.0.1", 1234)
    term.send("Hello")
    print(term.receive())
    term.close()
    # print(term.receive())
    