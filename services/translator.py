def setup_translator():
    try:
        from deep_translator import GoogleTranslator
        return GoogleTranslator(source='auto', target='ja')
    except:
        return None


def translate_project(text, translator):
    text_lower = text.lower()

    if "forti" in text_lower:
        return "セキュリティ対策商品（更新）" if "renewal" in text_lower else "セキュリティ対策商品"

    if "acronis" in text_lower:
        return "アンチウイルスソフト（更新）" if "renewal" in text_lower else "アンチウイルスソフト"

    if "esign" in text_lower:
        return "電子署名ソフト"

    if text.lower().startswith("ndid - "):
        text = text[7:].strip()

    if translator:
        try:
            return translator.translate(text)
        except:
            return text

    return text