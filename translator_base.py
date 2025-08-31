from google.cloud import translate_v2 as translate

class TranslatorBase:
    def __init__(self, credential_json="translate-tool.json"):
        try:
            self.client = translate.Client.from_service_account_json(credential_json)
        except Exception as e:
            raise RuntimeError(f"Error init translate client: {e}")

    def translate_text(self, text: str, target_language: str = "vi") -> str:
        if not text.strip():
            return ""
        result = self.client.translate(text, target_language=target_language)
        return result["translatedText"]
