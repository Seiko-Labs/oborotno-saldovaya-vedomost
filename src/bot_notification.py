import os
import requests
from requests.adapters import HTTPAdapter


class TelegramNotifier:
    def __init__(self, chat_id: str, session: requests.Session, token: str = None, retries: int = 5):
        token: str = os.getenv('TOKEN') if not token else token
        chat_id: str = chat_id
        self.api_url = f'https://api.telegram.org/bot{token}/sendMessage'
        self.api_params = {'chat_id': chat_id, 'parse_mode': 'Markdown'}
        self.retries = retries
        self.session = session
        self.session.mount("http://", HTTPAdapter(max_retries=self.retries))

    def send_notification(self, message: str) -> requests.models.Response:
        return self.session.post(self.api_url, params=self.api_params, json={'text': message})
