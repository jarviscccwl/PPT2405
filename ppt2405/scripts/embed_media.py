"""Embed local MP4 demonstrations in named picture slots; preserve native math.

Requires python-pptx and Pillow. Targets must be unanimated still-picture slots;
an existing movie is not a still picture, even though both use p:pic in OOXML.
The poster must be an uncropped frame of the video. Source provenance and video
dimensions are required in the media plan.
"""
import argparse
import hashlib
import json
from pathlib import Path
import posixpath
from zipfile import ZipFile
from lxml import etree
from PIL import Image
from pptx import Presentation


NS = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math'}


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def native_signature(path):
    """Compare actual math text/structure and native table cell text, not counts."""
    with ZipFile(path) as z:
        result = {}
        rels = etree.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
        targets = {r.get('Id'): r.get('Target') for r in rels}
        presentation = etree.fromstring(z.read('ppt/presentation.xml'))
        pns = 'http://schemas.openxmlformats.org/presentationml/2006/main'
        rns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        for index, slide in enumerate(presentation.findall('.//{%s}sldId' % pns)):
                target = targets[slide.get('{%s}id' % rns)]
                name = target.lstrip('/') if target.startswith('/') else posixpath.normpath('ppt/' + target)
                root = etree.fromstring(z.read(name))
                def structure(element):
                    return (element.tag, tuple(sorted(element.attrib.items())),
                            element.text or '', tuple(structure(x) for x in element))
                result[str(index)] = [structure(x) for x in root.findall('.//m:oMath', NS)]
                result[str(index) + ':tables'] = [structure(x) for x in root.findall('.//a:tbl', NS)]
        return result


def embed(input_path, plan_path, output_path):
    input_path, plan_path, output_path = map(Path, (input_path, plan_path, output_path))
    if input_path.resolve() == output_path.resolve():
        raise ValueError('Write a new output; do not overwrite the input deck.')
    plan = json.loads(plan_path.read_text(encoding='utf-8-sig'))
    if set(plan) != {'input_sha256', 'videos'}:
        raise ValueError('Plan requires exactly input_sha256 and videos.')
    if plan['input_sha256'] != digest(input_path):
        raise ValueError('Input deck SHA256 mismatch.')
    prs = Presentation(input_path)
    before = native_signature(input_path)
    report = {'input_sha256': digest(input_path), 'videos': []}
    occupied = set()
    for item in plan['videos']:
        required = {'slide', 'slot', 'path', 'poster', 'width_px', 'height_px', 'source'}
        if set(item) != required or not item['source']:
            raise ValueError(f'Media item requires exactly {sorted(required)} and a source.')
        slide_number = item['slide']
        if not isinstance(slide_number, int) or not 1 <= slide_number <= len(prs.slides):
            raise ValueError('Invalid one-based slide number.')
        key = (slide_number, item['slot'])
        if key in occupied:
            raise ValueError('Duplicate media target.')
        occupied.add(key)
        slide = prs.slides[slide_number - 1]
        targets = [s for s in slide.shapes if s.name == item['slot']]
        if len(targets) != 1 or targets[0]._element.tag.split('}')[-1] != 'pic':
            raise ValueError(f'{key}: expected exactly one named picture slot.')
        target = targets[0]
        # Removing a timed shape without updating its animation tree leaves
        # dangling spid references. In particular, python-pptx models both
        # pictures and movies as p:pic, so the tag check above is insufficient.
        if (target._element.find('.//a:videoFile', NS) is not None or
                target._element.find('.//a:audioFile', NS) is not None):
            raise ValueError(f'{key}: target already contains media; start from a still-picture deck.')
        if slide._element.xpath(f".//p:timing//*[@spid='{target.shape_id}']"):
            raise ValueError(f'{key}: target has animation/timing references; use an unanimated picture slot.')
        video, poster = [(plan_path.parent / item[k]).resolve() for k in ('path', 'poster')]
        if video.suffix.lower() != '.mp4' or not video.is_file():
            raise ValueError('Use a local MP4.')
        vw, vh = item['width_px'], item['height_px']
        if vw <= 0 or vh <= 0:
            raise ValueError('Video dimensions must be positive.')
        with Image.open(poster) as im:
            pw, ph = im.size
        if abs((pw / ph) / (vw / vh) - 1) > .01:
            raise ValueError('Poster/video aspect ratios differ: use a full video frame.')
        scale = min(target.width / vw, target.height / vh)
        w, h = round(vw * scale), round(vh * scale)
        x, y = target.left + (target.width - w) // 2, target.top + (target.height - h) // 2
        movie = slide.shapes.add_movie(str(video), x, y, w, h,
                                       poster_frame_image=str(poster), mime_type='video/mp4')
        movie.name = item['slot']
        # Keep the target's stacking position; the newly generated timing refers
        # to the movie's own shape id and remains intact after moving the element.
        tree = target._element.getparent()
        index = list(tree).index(target._element)
        tree.remove(movie._element)
        tree.insert(index, movie._element)
        tree.remove(target._element)
        report['videos'].append({'slide': slide_number, 'slot': item['slot'],
            'sha256': digest(video), 'poster_sha256': digest(poster),
            'source': item['source'], 'placed_emu': [x, y, w, h],
            'contain': True, 'playback': 'click', 'pdf': 'poster frame'})
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)
    if native_signature(output_path) != before:
        raise ValueError('Native formula or table content changed during media embedding.')
    report.update(status='passed', output_sha256=digest(output_path),
                  native_math_and_tables_preserved=True)
    return report


def main():
    p = argparse.ArgumentParser(description=__doc__)
    for name in ('input', 'plan', 'output', 'report'):
        p.add_argument('--' + name, required=True)
    p.add_argument('--overwrite', action='store_true')
    args = p.parse_args()
    if Path(args.output).exists() and not args.overwrite:
        p.error('Output exists; use --overwrite for a generated output.')
    result = embed(args.input, args.plan, args.output)
    Path(args.report).write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({'status': result['status'], 'videos': len(result['videos']),
                      'output': args.output}, ensure_ascii=False))


if __name__ == '__main__':
    main()
