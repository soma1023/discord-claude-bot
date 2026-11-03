# -*- coding: utf-8 -*-
import base64
import gzip
import os
import tempfile
import unittest

from stream_highlight import prproj


def blob(text):
    # 実物と同じく、フォント名などの ASCII とテロップの文字がバイナリに混ざっている
    return base64.b64encode(b"\x18\x02\x00\x00NewRodinPro-EB\x00\x00" + text.encode("utf-8") + b"\x00\x01").decode()


def clip(oid, chain, start, end):
    return ('<VideoClipTrackItem ObjectID="%d" Version="8"><ClipTrackItem><ComponentOwner>'
            '<Components ObjectRef="%d"/></ComponentOwner><TrackItem><Start>%d</Start><End>%d</End>'
            '</TrackItem></ClipTrackItem></VideoClipTrackItem>' % (oid, chain, start, end))


def chain(oid, component):
    return ('<VideoComponentChain ObjectID="%d"><Components><Component Index="0" ObjectRef="%d"/>'
            '</Components></VideoComponentChain>' % (oid, component))


def component(oid, text_param, position_param):
    return ('<VideoFilterComponent ObjectID="%d"><Params><Param Index="0" ObjectRef="%d"/>'
            '<Param Index="1" ObjectRef="%d"/></Params></VideoFilterComponent>' % (oid, text_param, position_param))


def text_param(oid, value):
    return ('<ArbVideoComponentParam ObjectID="%d"><Name>ソーステキスト</Name>%s</ArbVideoComponentParam>'
            % (oid, value))


def position_param(oid, x, y):
    return ('<PointComponentParam ObjectID="%d"><Name>位置</Name>'
            '<StartKeyframe>-91445760000000000,%s:%s,0,0</StartKeyframe></PointComponentParam>' % (oid, x, y))


SAMPLE = "<Project>" + "".join([
    clip(1, 2, 0, 1000), chain(2, 3), component(3, 4, 5),
    text_param(4, '<StartKeyframeValue Encoding="base64" BinaryHash="aa">%s</StartKeyframeValue>'
               % blob("はい、ではまず\rれももさんの方から")),
    position_param(5, 0.125, 0.65),
    # 同じ値を持つ2枚目は、中身を書かずに BinaryHash だけで参照する
    clip(11, 12, 1000, 2000), chain(12, 13), component(13, 14, 15),
    text_param(14, '<StartKeyframeValue Encoding="base64" BinaryHash="aa"/>'),
    position_param(15, 0.125, 0.65),
    clip(21, 22, 2000, 3000), chain(22, 23), component(23, 24, 25),
    text_param(24, '<StartKeyframeValue Encoding="base64" BinaryHash="bb">%s</StartKeyframeValue>'
               % blob("よろしくお願いします")),
    position_param(25, 0.0375, 0.85),
]) + "</Project>"


class TestCaptions(unittest.TestCase):
    def test_lines_and_ranges(self):
        found = sorted(prproj.captions(SAMPLE), key=lambda c: c["start"])
        self.assertEqual([(c["start"], c["end"], c["lines"]) for c in found],
                         [(0, 1000, 2), (1000, 2000, 2), (2000, 3000, 1)])
        self.assertEqual(found[0]["text"], "はい、ではまず\rれももさんの方から")
        self.assertEqual(found[2]["position"], [0.0375, 0.85])

    def test_line_count(self):
        self.assertEqual(prproj.line_count("あ\r\nい\n\nう"), 3)
        self.assertEqual(prproj.line_count(""), 1)

    def test_reads_gzipped_file(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "a.prproj")
            with open(path, "wb") as fh:
                fh.write(gzip.compress(SAMPLE.encode("utf-8")))
            self.assertEqual(len(prproj.caption_lines(path)), 3)


if __name__ == "__main__":
    unittest.main()
