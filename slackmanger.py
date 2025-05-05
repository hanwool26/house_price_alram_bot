import slack_sdk
from slack_sdk.errors import SlackApiError
import requests

class SlackManager:
    def __init__(self, channel_id, slack_token):
        self.channel_id = channel_id
        self.slack_token = slack_token
    
    def upload_file(self, file_path):
        file_name = file_path
        channel_id = self.channel_id
        token = self.slack_token

        client =slack_sdk.WebClient(token=token)

        print(file_name)

        try:
            result = client.files_upload_v2(channels=channel_id, file=file_name)
            print('Success send file via slack')
        except SlackApiError as e :
            print("Error uploading file: {}".format(e))
            
    def send_message(self, message):
        url = "https://slack.com/api/chat.postMessage"
        channel_id = self.channel_id
        token = self.slack_token
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }
        
        # 메시지 내용
        payload = {
            "channel": channel_id,  # 메시지를 보낼 채널 ID (사용자 DM일 경우 사용자 ID)
            "text": message,         # 보낼 메시지 내용
        }

        # POST 요청 보내기
        response = requests.post(url, headers=headers, json=payload)

        # 응답 확인
        if response.status_code == 200 and response.json().get('ok'):
            print(f"Message successfully sent to Slack.")
        else:
            print(f"Failed to send message to Slack. Error: {response.text}")

