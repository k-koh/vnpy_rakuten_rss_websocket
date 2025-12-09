from websocket import create_connection
import json

ws = create_connection("ws://localhost:8765/ws/")
print("Connected to Excel Addin WebSocket")

while True:
    msg = ws.recv()
    data = json.loads(msg)
    print("Received:", data)