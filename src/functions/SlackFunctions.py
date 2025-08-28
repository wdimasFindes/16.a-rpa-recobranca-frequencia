import requests
import json
import  os
from dotenv import load_dotenv

load_dotenv()

class SlackNotifier:
    def __init__(self, hook_url, channel, username, emoji="robot_face"):
        self.hook_url = hook_url
        self.channel = channel
        self.username = username
        self.emoji = emoji

    def parse_message(self, message):
        return {"channel": self.channel, "username": self.username, "text": message, "icon_emoji": ":" + self.emoji + ":"}

    def post_message(self, message):

        if self.hook_url is not None and self.hook_url != '':
            body = self.parse_message(message)
            headers = {'Content-Type': 'application/json'}
            response = requests.post(self.hook_url, headers=headers, data=json.dumps(body))
            response.raise_for_status()