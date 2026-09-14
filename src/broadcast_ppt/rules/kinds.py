"""Slide kinds derived from the original decks' own structure."""

INTRO_MEDIA = "intro_media"
PRAYER_TITLE = "prayer_title"
CARD_TITLE = "card_title"           # generic order card: 예배로의 부름, 동행의 기도, 결단의 기도, 찬양 및 봉헌 ...
SCRIPTURE_TITLE = "scripture_title"
SERMON_TITLE = "sermon_title"
SCRIPTURE_BODY = "scripture_body"   # full-screen scripture text (reading or sermon quote)
MEDITATION = "meditation"
ANNOUNCEMENT_TITLE = "announcement_title"
ANNOUNCEMENT_BLOCK = "announcement_block"
CAPTION_SENTENCE = "caption_sentence"
HYMN_IMAGE = "hymn_image"
NAME_LIST = "name_list"
PICTURE_ONLY = "picture_only"
UNKNOWN = "unknown"

# kinds that are copied through unchanged (scaled) with a review flag
PASSTHROUGH = {HYMN_IMAGE, NAME_LIST, PICTURE_ONLY, UNKNOWN, INTRO_MEDIA}
