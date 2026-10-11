"""Servicio de notificaciones y entrega de archivos via Telegram Bot API."""
import os
import requests
from src.config import TELEGRAM_TOKEN, TELEGRAM_CHAT_ID


def send_telegram_reports(summary_text, file_paths, token=None, chat_id=None):
    """
    Envia el resumen HTML y los archivos generados al canal de Telegram configurado.
    """
    bot_token = token or TELEGRAM_TOKEN
    target_chat = chat_id or TELEGRAM_CHAT_ID

    print("--- INICIANDO ENVIO A TELEGRAM ---")
    url_msg = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    payload = {
        "chat_id": target_chat,
        "text": summary_text,
        "parse_mode": "HTML"
    }
    
    try:
        r = requests.post(url_msg, data=payload)
        res = r.json()
        if res.get("ok"):
            print("Mensaje de resumen enviado correctamente a Telegram.")
        else:
            print(f"Error al enviar resumen a Telegram: {res}")
            return False
    except Exception as e:
        print(f"Error de red al enviar mensaje a Telegram: {e}")
        return False

    url_doc = f"https://api.telegram.org/bot{bot_token}/sendDocument"
    for caption, path in file_paths.items():
        if os.path.exists(path):
            print(f"Enviando '{os.path.basename(path)}'...")
            try:
                with open(path, "rb") as f:
                    files = {"document": f}
                    data = {"chat_id": target_chat, "caption": caption}
                    r = requests.post(url_doc, data=data, files=files)
                    res = r.json()
                    if res.get("ok"):
                        print(f"Archivo '{os.path.basename(path)}' enviado con éxito.")
                    else:
                        print(f"Error de Telegram al enviar archivo '{os.path.basename(path)}': {res}")
            except Exception as e:
                print(f"Error al enviar archivo '{os.path.basename(path)}': {e}")
        else:
            print(f"Advertencia: El archivo '{path}' no existe y no se enviará.")
            
    print("--- ENVIO A TELEGRAM FINALIZADO ---")
    return True
