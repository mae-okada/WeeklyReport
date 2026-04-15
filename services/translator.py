class TranslatorService:
    def __init__(self):
        self.translator = self._setup_translator()

    def _setup_translator(self):
        try:
            from deep_translator import GoogleTranslator
            return GoogleTranslator(source='auto', target='ja')
        except Exception:
            return None

    def translate_project(self, text):
        text_lower = text.lower()

        if "forti" in text_lower:
            return "セキュリティ対策商品（更新）" if "renewal" in text_lower else "セキュリティ対策商品"

        if "acronis" in text_lower:
            return "アンチウイルスソフト（更新）" if "renewal" in text_lower else "アンチウイルスソフト"

        if "esign" in text_lower:
            return "電子署名ソフト"

        if text.lower().startswith("ndid - "):
            text = text[7:].strip()

        if self.translator:
            try:
                return self.translator.translate(text)
            except Exception:
                return text

        return text