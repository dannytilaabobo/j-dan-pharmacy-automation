import html
import re
import datetime

DASH_CHARS = ['–', '—', '‒', '―']  # en dash, em dash, figure dash, horizontal bar
QUOTE_MAP = {
    '‘': "'", '’': "'", '‚': "'", '‛': "'",
    '“': '"', '”': '"', '„': '"', '‟': '"',
}


def is_date_corrupted(value):
    return isinstance(value, (datetime.datetime, datetime.date))


def clean_text(value):
    """Normalize a product-name-like string: decode HTML entities, normalize
    dashes/quotes, collapse whitespace. Returns non-str, non-corrupted values
    (numbers, None) unchanged; date-corrupted values are returned as-is since
    they need manual review, not a guessed text reconstruction."""
    if value is None:
        return value
    if is_date_corrupted(value):
        return value
    if not isinstance(value, str):
        return value

    text = html.unescape(value)
    for dash in DASH_CHARS:
        text = text.replace(dash, '-')
    for smart, straight in QUOTE_MAP.items():
        text = text.replace(smart, straight)
    text = re.sub(r'\s+', ' ', text).strip()
    return text
