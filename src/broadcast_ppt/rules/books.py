"""Korean Bible book table: full names, common abbreviations, testament, order."""

from __future__ import annotations

# (full name, [abbreviations and spelling variants], testament)
_BOOKS: list[tuple[str, list[str], str]] = [
    ("창세기", ["창"], "old"), ("출애굽기", ["출"], "old"), ("레위기", ["레"], "old"), ("민수기", ["민"], "old"),
    ("신명기", ["신"], "old"), ("여호수아", ["수"], "old"), ("사사기", ["삿"], "old"), ("룻기", ["룻"], "old"),
    ("사무엘상", ["삼상"], "old"), ("사무엘하", ["삼하"], "old"), ("열왕기상", ["왕상"], "old"), ("열왕기하", ["왕하"], "old"),
    ("역대상", ["대상"], "old"), ("역대하", ["대하"], "old"), ("에스라", ["스"], "old"), ("느헤미야", ["느"], "old"),
    ("에스더", ["에"], "old"), ("욥기", ["욥"], "old"), ("시편", ["시"], "old"), ("잠언", ["잠"], "old"),
    ("전도서", ["전"], "old"), ("아가", ["아"], "old"), ("이사야", ["사"], "old"), ("예레미야", ["렘", "예레미아"], "old"),
    ("예레미야애가", ["애", "예레미아애가"], "old"), ("에스겔", ["겔"], "old"), ("다니엘", ["단"], "old"), ("호세아", ["호"], "old"),
    ("요엘", ["욜"], "old"), ("아모스", ["암"], "old"), ("오바댜", ["옵"], "old"), ("요나", ["욘"], "old"),
    ("미가", ["미"], "old"), ("나훔", ["나"], "old"), ("하박국", ["합"], "old"), ("스바냐", ["습"], "old"),
    ("학개", ["학"], "old"), ("스가랴", ["슥"], "old"), ("말라기", ["말"], "old"),
    ("마태복음", ["마", "마태"], "new"), ("마가복음", ["막", "마가"], "new"), ("누가복음", ["눅", "누가"], "new"),
    ("요한복음", ["요", "요한"], "new"), ("사도행전", ["행"], "new"), ("로마서", ["롬"], "new"),
    ("고린도전서", ["고전"], "new"), ("고린도후서", ["고후"], "new"), ("갈라디아서", ["갈"], "new"),
    ("에베소서", ["엡"], "new"), ("빌립보서", ["빌"], "new"), ("골로새서", ["골"], "new"),
    ("데살로니가전서", ["살전"], "new"), ("데살로니가후서", ["살후"], "new"), ("디모데전서", ["딤전"], "new"),
    ("디모데후서", ["딤후"], "new"), ("디도서", ["딛"], "new"), ("빌레몬서", ["몬"], "new"), ("히브리서", ["히"], "new"),
    ("야고보서", ["약"], "new"), ("베드로전서", ["벧전"], "new"), ("베드로후서", ["벧후"], "new"),
    ("요한일서", ["요일"], "new"), ("요한이서", ["요이"], "new"), ("요한삼서", ["요삼"], "new"),
    ("유다서", ["유"], "new"), ("요한계시록", ["계", "계시록"], "new"),
]

BOOKS: dict[str, tuple[str, str, int]] = {}  # any spelling -> (full name, testament, order)
for _order, (_full, _abbrs, _test) in enumerate(_BOOKS, start=1):
    BOOKS[_full] = (_full, _test, _order)
    for _a in _abbrs:
        BOOKS[_a] = (_full, _test, _order)

# Longest spelling first so "사무엘상" wins over "사", "요한일서" over "요".
BOOK_NAMES_BY_LENGTH: list[str] = sorted(BOOKS, key=len, reverse=True)
