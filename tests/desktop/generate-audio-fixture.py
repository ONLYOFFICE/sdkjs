#!/usr/bin/env python3
# SPDX-License-Identifier: AGPL-3.0-only
"""Generate an original two-slide click-audio PPTX. Requires python-pptx and ffmpeg."""
import io
from pathlib import Path
import subprocess
import tempfile

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.util import Inches


def main():
    presentation = Presentation()
    presentation.core_properties.author = ""
    presentation.core_properties.last_modified_by = ""
    presentation.core_properties.title = "Next triggers embedded audio"
    presentation.core_properties.comments = "Synthetic regression fixture; generated sine tones."
    image = Image.new("RGB", (96, 96), "white")
    ImageDraw.Draw(image).ellipse((12, 12, 84, 84), fill="#2463a0")
    poster = io.BytesIO()
    image.save(poster, format="PNG")
    with tempfile.TemporaryDirectory() as temporary:
        for frequency in (440, 880):
            audio = Path(temporary) / f"tone-{frequency}.mp3"
            subprocess.run([
                "ffmpeg", "-loglevel", "error", "-f", "lavfi", "-i",
                f"sine=frequency={frequency}:duration=2", "-ac", "1", "-q:a", "7", str(audio)
            ], check=True)
            slide = presentation.slides.add_slide(presentation.slide_layouts[6])
            text = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(2))
            text.text = f"Press Next once to hear a {frequency} Hz tone.\nDo not select the audio icon."
            poster.seek(0)
            shape = slide.shapes.add_movie(str(audio), Inches(4), Inches(4), Inches(1), Inches(1),
                                           poster_frame_image=poster, mime_type="audio/mpeg")
            # add_movie supplies the shared media picture and relationships.
            # Audio uses a:audioFile with the same linked media relationship.
            shape._element.find(".//" + qn("a:videoFile")).tag = qn("a:audioFile")
            slide._element.remove(slide._element.find(qn("p:timing")))
            slide._element.append(parse_xml(f'''<p:timing
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
                <p:childTnLst><p:seq concurrent="1" nextAc="seek">
                  <p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>
                    <p:par><p:cTn id="3" fill="hold">
                      <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
                      <p:childTnLst><p:par><p:cTn id="4" fill="hold">
                        <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                        <p:childTnLst><p:par><p:cTn id="5" presetID="1" presetClass="mediacall"
                            presetSubtype="0" fill="hold" nodeType="clickEffect">
                          <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                          <p:childTnLst><p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr>
                            <p:cTn id="6" dur="2000" fill="hold"/>
                            <p:tgtEl><p:spTgt spid="{shape.shape_id}"/></p:tgtEl>
                          </p:cBhvr></p:cmd></p:childTnLst>
                        </p:cTn></p:par></p:childTnLst>
                      </p:cTn></p:par></p:childTnLst>
                    </p:cTn></p:par>
                  </p:childTnLst></p:cTn>
                  <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
                  <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
                </p:seq>
                <p:audio><p:cMediaNode vol="100000">
                  <p:cTn id="7" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst></p:cTn>
                  <p:tgtEl><p:spTgt spid="{shape.shape_id}"/></p:tgtEl>
                </p:cMediaNode></p:audio></p:childTnLst>
              </p:cTn></p:par></p:tnLst>
            </p:timing>'''))
        output = Path(__file__).with_name("fixtures") / "next-click-audio.pptx"
        output.parent.mkdir(exist_ok=True)
        presentation.save(output)
        print(output)


if __name__ == "__main__":
    main()
