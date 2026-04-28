import unittest

import wecom_delivery


class FakeClient:
    instances = []

    def __init__(self, bot_id, secret):
        self.bot_id = bot_id
        self.secret = secret
        self.connected = False
        self.sent = []
        self.disconnected = False
        FakeClient.instances.append(self)

    @property
    def is_connected(self):
        return self.connected

    async def connect(self):
        self.connected = True
        return self

    async def send_message(self, chat_id, body):
        self.sent.append((chat_id, body))
        return {"chatid": chat_id, "body": body}

    async def disconnect(self):
        self.connected = False
        self.disconnected = True


class WecomDeliveryTests(unittest.TestCase):
    def tearDown(self):
        FakeClient.instances.clear()

    def test_missing_smart_bot_fields(self):
        self.assertEqual(
            wecom_delivery.missing_smart_bot_fields("", "secret", ""),
            ["Bot ID", "Chat ID"],
        )

    def test_build_smart_bot_body(self):
        body = wecom_delivery.build_smart_bot_body("hello")
        self.assertEqual(body["msgtype"], "markdown")
        self.assertEqual(body["markdown"]["content"], "hello")

    def test_sender_reuses_client_with_same_credentials(self):
        sender = wecom_delivery.SmartBotSender(client_factory=FakeClient)
        try:
            sender.send_markdown("bot-1", "sec-1", "chat-1", "first")
            sender.send_markdown("bot-1", "sec-1", "chat-2", "second")
            self.assertEqual(len(FakeClient.instances), 1)
            self.assertEqual(
                FakeClient.instances[0].sent,
                [
                    ("chat-1", wecom_delivery.build_smart_bot_body("first")),
                    ("chat-2", wecom_delivery.build_smart_bot_body("second")),
                ],
            )
        finally:
            sender.close()

    def test_sender_reconnects_when_credentials_change(self):
        sender = wecom_delivery.SmartBotSender(client_factory=FakeClient)
        try:
            sender.send_markdown("bot-1", "sec-1", "chat-1", "first")
            first_client = FakeClient.instances[0]
            sender.send_markdown("bot-2", "sec-2", "chat-1", "second")
            self.assertEqual(len(FakeClient.instances), 2)
            self.assertTrue(first_client.disconnected)
        finally:
            sender.close()


if __name__ == "__main__":
    unittest.main()
