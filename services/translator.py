class TranslatorService:
    """Service to translate project names using heuristics or an optional backend."""

    def __init__(self):
        """Initialize translation service and configure the backend translator."""
        self.translator = self._setup_translator()

    def _setup_translator(self):
        """Attempt to configure GoogleTranslator from deep_translator."""
        try:
            from deep_translator import GoogleTranslator
            return GoogleTranslator(source='auto', target='ja')
        except Exception:
            return None

    def translate_project(self, text):
        """Translate a project name using keyword rules and fallback logic."""
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