import socket
import serial

# 配置串口
ser = serial.Serial('COM5', 9600)  # Arduino 所连接的端口

# 启动TCP服务器，等待教师端连接
def start_server():
    server_ip = ''
    server_port = 12345
    server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server_socket.bind((server_ip, server_port))
    server_socket.listen(5)

    print("教室端已启动，等待教师端连接...")

    while True:
        client_socket, client_address = server_socket.accept()
        print(f"来自教师端 {client_address} 的连接已建立")

        # 从教师端接收数据
        data = client_socket.recv(1024).decode()
        if data:
            print(f"接收到的数据: {data}")
            # 通过串口发送给 Arduino
            ser.write((data + '\n').encode())
        else:
            print("接收到的数据为空")

        client_socket.close()

if __name__ == "__main__":
    start_server()
