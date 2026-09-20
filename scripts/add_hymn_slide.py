"""Add the hymn caption slide to the template deck (one-off; the slide is edited in PowerPoint afterwards).

    uv run python scripts/add_hymn_slide.py [template]

A rounded box with the lyric on top of it, on the master's green. The box is the shape of the chips the
source decks use (roundRect, thin grey outline) but filled solid black: on the chroma key anything
translucent lets the green through and the keyer takes the box with it. The lyric has no character
outline, only the drop shadow.
Widths here are only a starting point — `broadcast-ppt` sizes the box to each lyric at build time.
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from broadcast_ppt.builder.package import add_blank_slide  # noqa: E402
from broadcast_ppt.ooxml import A, P, REPO_ROOT  # noqa: E402

TEMPLATE = REPO_ROOT / "local/templates/24.상반기.pptx"
FONT = "나눔스퀘어 ExtraBold"
BOX = (996000, 5990467, 10200000, 630000)
LYRIC = (1096000, 5920747, 10000000, 769441)
LABEL = (1096000, 5520637, 10000000, 400110)
SHADOW = ('<a:effectLst><a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl">'
          '<a:srgbClr val="000000"><a:alpha val="43137"/></a:srgbClr></a:outerShdw></a:effectLst>')

SLIDE = f"""<p:spTree xmlns:p="{P}" xmlns:a="{A}">
  <p:sp>
    <p:nvSpPr><p:cNvPr id="2" name="가사 배경"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="{BOX[0]}" y="{BOX[1]}"/><a:ext cx="{BOX[2]}" cy="{BOX[3]}"/></a:xfrm>
      <a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom>
      <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      <a:ln w="3175" cap="flat"><a:solidFill><a:srgbClr val="808080"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>
    </p:spPr>
    <p:txBody><a:bodyPr anchor="ctr"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p/></p:txBody>
  </p:sp>
  <p:sp>
    <p:nvSpPr><p:cNvPr id="3" name="가사"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="{LYRIC[0]}" y="{LYRIC[1]}"/><a:ext cx="{LYRIC[2]}" cy="{LYRIC[3]}"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>
    </p:spPr>
    <p:txBody>
      <a:bodyPr wrap="none" rtlCol="0" anchor="ctr"><a:spAutoFit/></a:bodyPr><a:lstStyle/>
      <a:p><a:pPr algn="ctr"/>
        <a:r><a:rPr lang="ko-KR" altLang="en-US" sz="4400" b="1" dirty="0">
          <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>{SHADOW}
          <a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/></a:rPr><a:t>샘플 가사 한 줄입니다</a:t></a:r>
      </a:p>
    </p:txBody>
  </p:sp>
  <p:sp>
    <p:nvSpPr><p:cNvPr id="4" name="곡 제목"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="{LABEL[0]}" y="{LABEL[1]}"/><a:ext cx="{LABEL[2]}" cy="{LABEL[3]}"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>
    </p:spPr>
    <p:txBody>
      <a:bodyPr wrap="none" rtlCol="0" anchor="ctr"><a:spAutoFit/></a:bodyPr><a:lstStyle/>
      <a:p><a:pPr algn="ctr"/>
        <a:r><a:rPr lang="ko-KR" altLang="en-US" sz="2000" b="1" dirty="0">
          <a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill>{SHADOW}
          <a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/></a:rPr><a:t>108장 그 어린 주 예수</a:t></a:r>
      </a:p>
    </p:txBody>
  </p:sp>
</p:spTree>"""


def main(path: Path) -> int:
    prs = Presentation(str(path))
    slide = add_blank_slide(prs)
    tree = slide.shapes._spTree
    for sp in etree.fromstring(SLIDE.encode("utf-8")):
        tree.append(sp)
    backup = path.with_suffix(path.suffix + ".bak")
    if not backup.exists():
        shutil.copy2(path, backup)
    prs.save(str(path))
    print(f"템플릿 {len(prs.slides)}장째에 찬송 자막 슬라이드 추가 -> {path} (백업: {backup.name})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(Path(sys.argv[1]) if len(sys.argv) > 1 else TEMPLATE))
