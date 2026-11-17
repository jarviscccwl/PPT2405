#!/usr/bin/env python3
"""Inspect PDFs and extract complete figures from full-page renders.

No embedded-image extraction, automatic figure segmentation, or automatic visual
acceptance. Coordinates are local top-left PDF points on the rendered MediaBox.
"""
from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import json
import math
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys


class InputError(ValueError):
    pass


def sha256(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(chunk)
    return h.hexdigest()


def write_json(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('x', encoding='utf-8') as stream:
        json.dump(data, stream, ensure_ascii=False, indent=2, allow_nan=False)
        stream.write('\n')


def require_pdf(value, base=None):
    path = Path(value).expanduser()
    if base is not None and not path.is_absolute():
        path = base / path
    path = path.resolve()
    if not path.is_file():
        raise InputError(f'PDF does not exist: {path}')
    return path


def selected_pages(spec, count):
    if spec is None:
        return list(range(1, count + 1))
    pages = set()
    for part in str(spec).split(','):
        part = part.strip()
        if not re.fullmatch(r'\d+(?:-\d+)?', part):
            raise InputError('pages must be 1-based numbers/ranges, e.g. 1,3-5')
        ends = [int(x) for x in part.split('-')]
        first, last = ends[0], ends[-1]
        if not 1 <= first <= last <= count:
            raise InputError(f'page range {part} outside 1..{count}')
        pages.update(range(first, last + 1))
    return sorted(pages)


def checked_bbox(value, page, label='bbox_pt'):
    if not isinstance(value, list) or len(value) != 4:
        raise InputError(f'{label}: bbox must be [left, top, right, bottom]')
    if any(isinstance(x, bool) or not isinstance(x, (int, float)) or not math.isfinite(x) for x in value):
        raise InputError(f'{label}: bbox values must be finite numbers')
    left, top, right, bottom = map(float, value)
    if not (0 <= left < right <= page.width and 0 <= top < bottom <= page.height):
        raise InputError(f'{label}: bbox must have positive area and stay inside page [0,0,{page.width},{page.height}]')
    return [left, top, right, bottom]


def text_in_bbox(page, bbox):
    x0, top, _, _ = page.bbox
    l, t, r, b = bbox
    return page.crop((x0+l, top+t, x0+r, top+b)).extract_text(x_tolerance=1, y_tolerance=3) or ''


def inspect_page(page, number):
    # Extracted lines can merge two columns and body references can resemble
    # captions. These are explicitly candidates, never accepted figure regions.
    caption_pattern = re.compile(r'^\s*((?:Fig(?:ure)?\.?|Table|图|表)\s*(?:\d+|[IVXLC]+))', re.I)
    candidates = []
    x0, top, _, _ = page.bbox
    for line in page.extract_text_lines(layout=False, return_chars=False, x_tolerance=1):
        match = caption_pattern.match(line['text'])
        if match:
            candidates.append({'label_candidate':match.group(1), 'text':line['text'],
                               'bbox_pt':[line['x0']-x0,line['top']-top,line['x1']-x0,line['bottom']-top],
                               'candidate_only':True})
    return {'page':number,'width_pt':page.width,'height_pt':page.height,
            'rotation_degrees':page.rotation,'media_box':list(page.mediabox),
            'crop_box':list(page.cropbox) if page.cropbox else None,
            'counts':{'characters':len(page.chars),'words':len(page.extract_words()),
                      'images':len(page.images),'lines':len(page.lines),
                      'curves':len(page.curves),'rectangles':len(page.rects)},
            'caption_candidates':candidates}


def find_pdftoppm(explicit=None):
    # Explicit file/dir, environment, PATH, then current Python runtime siblings.
    # No machine-specific runtime installation or user directory is encoded.
    candidates = []
    for setting in [explicit, os.environ.get('PPT2405_POPPLER')]:
        if setting:
            candidate = Path(setting).expanduser()
            if candidate.is_dir():
                candidate /= 'pdftoppm.exe' if os.name == 'nt' else 'pdftoppm'
            if not candidate.is_file():
                raise InputError(f'pdftoppm not found at configured path: {candidate}')
            return str(candidate.resolve())
    found = shutil.which('pdftoppm')
    if found:
        return found
    for root in list(Path(sys.executable).resolve().parents)[:4]:
        candidates.extend([root/'native/poppler/Library/bin/pdftoppm.exe',
                           root/'poppler/Library/bin/pdftoppm.exe',
                           root/'bin/pdftoppm', root/'bin/pdftoppm.exe'])
    for candidate in candidates:
        if candidate.is_file():
            return str(candidate)
    raise InputError('pdftoppm unavailable. Use bundled dependencies and pass --pdftoppm FILE_OR_BIN_DIR, or set PPT2405_POPPLER/PATH. No packages are installed automatically.')


def output_directory(value):
    out = Path(value).expanduser().resolve()
    if out.exists():
        raise InputError(f'Output directory already exists; choose a new directory: {out}')
    return out


def render_page(pdf, number, dpi, prefix, page_width, page_height, executable):
    from PIL import Image
    command = [executable,'-f',str(number),'-l',str(number),'-r',str(dpi),'-png','-singlefile',str(pdf),str(prefix)]
    result = subprocess.run(command, capture_output=True, timeout=180)
    if result.returncode:
        raise InputError('pdftoppm failed: '+result.stderr.decode('utf-8',errors='replace').strip())
    image_path = prefix.with_suffix('.png')
    with Image.open(image_path) as opened:
        full = opened.convert('RGB')
    # Fail safely if page geometry / rotation interpretation differs from Poppler.
    if abs(full.width-page_width*dpi/72) > 2 or abs(full.height-page_height*dpi/72) > 2:
        raise InputError('Rendered page size disagrees with PDF MediaBox dimensions; do not use these coordinates without reviewing rotation/page boxes.')
    return full, image_path


def thumbnail(image, path, width=1600):
    review = image.copy()
    review.thumbnail((width, 2000))
    review.save(path)


def inspect_command(args, pdfplumber):
    pdf = require_pdf(args.pdf)
    if Path(args.output).exists():
        raise InputError(f'Output exists: {args.output}')
    with pdfplumber.open(pdf) as document:
        result = {'source_pdf':str(pdf),'source_sha256':sha256(pdf),'page_count':len(document.pages),
                  'coordinate_system':'PDF points, local top-left of rotated MediaBox; 1 point = 1/72 inch',
                  'caption_warning':'Candidates only; body references/merged columns require visual confirmation.',
                  'pages':[inspect_page(document.pages[n-1],n) for n in selected_pages(args.pages,len(document.pages))]}
    write_json(args.output,result)
    return {'inspection':str(Path(args.output).resolve()),'pages_inspected':len(result['pages'])}


def render_command(args, pdfplumber):
    pdf = require_pdf(args.pdf)
    out = output_directory(args.out_dir)
    executable = find_pdftoppm(args.pdftoppm)
    with pdfplumber.open(pdf) as document:
        pages = selected_pages(args.pages,len(document.pages))
        dimensions = [(n,document.pages[n-1].width,document.pages[n-1].height) for n in pages]
    out.mkdir(parents=True)
    records = []
    for number,width,height in dimensions:
        full,path = render_page(pdf,number,args.dpi,out/f'page-{number:04}',width,height,executable)
        preview=out/f'page-{number:04}-preview.png'
        thumbnail(full,preview)
        records.append({'page':number,'size_pt':[width,height],'size_px':list(full.size),
                        'full_page':path.name,'preview':preview.name})
    write_json(out/'render-manifest.json',{'source_pdf':str(pdf),'source_sha256':sha256(pdf),'dpi':args.dpi,'pages':records})
    return {'output_directory':str(out),'pages_rendered':len(records)}


def extract_command(args, pdfplumber):
    from PIL import ImageDraw
    plan_path = Path(args.plan).expanduser().resolve()
    plan = json.loads(plan_path.read_text(encoding='utf-8-sig'))
    if not isinstance(plan,dict) or not isinstance(plan.get('figures'),list) or not plan['figures']:
        raise InputError('Plan must contain a non-empty figures array')
    out = output_directory(args.out_dir)
    validated, ids, sources = [], set(), {}
    # Validate the complete plan BEFORE making output or invoking a renderer.
    for entry in plan['figures']:
        if not isinstance(entry,dict):
            raise InputError('Each figure must be an object')
        identifier = entry.get('id')
        reserved = {'con','prn','aux','nul'} | {f'{prefix}{n}' for prefix in ('com','lpt') for n in range(1,10)}
        if (not isinstance(identifier,str) or not re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9._-]{0,99}',identifier)
                or identifier.endswith('.') or identifier.split('.')[0].casefold() in reserved
                or identifier.casefold() in ids):
            raise InputError('Figure id must be unique ignoring case and use 1–100 filename-safe ASCII letters/digits/dot/dash/underscore, starting with a letter/digit; trailing dots and Windows device names are invalid')
        ids.add(identifier.casefold())
        source_value = entry.get('source_pdf',plan.get('source_pdf'))
        if not isinstance(source_value,str) or not source_value:
            raise InputError(f'{identifier}: source_pdf is required')
        pdf = require_pdf(source_value,plan_path.parent)
        number = entry.get('page')
        if isinstance(number,bool) or not isinstance(number,int):
            raise InputError(f'{identifier}: page must be a 1-based integer')
        panels = entry.get('expected_panels')
        if not isinstance(panels,list) or not panels or any(not isinstance(x,str) or not x.strip() for x in panels):
            raise InputError(f'{identifier}: non-empty expected_panels list is required (use one description for a single-panel figure)')
        figure_number = entry.get('figure_number')
        if not isinstance(figure_number,str) or not figure_number.strip():
            raise InputError(f'{identifier}: figure_number is required')
        with pdfplumber.open(pdf) as document:
            if not 1 <= number <= len(document.pages):
                raise InputError(f'{identifier}: page {number} outside 1..{len(document.pages)}')
            page = document.pages[number-1]
            bbox = checked_bbox(entry.get('bbox_pt'),page,f'{identifier} bbox_pt')
            caption_bbox = entry.get('caption_bbox_pt')
            if caption_bbox is not None:
                caption_bbox = checked_bbox(caption_bbox,page,f'{identifier} caption_bbox_pt')
            caption = text_in_bbox(page,caption_bbox) if caption_bbox else entry.get('caption_text','')
            if not isinstance(caption,str):
                raise InputError(f'{identifier}: caption_text must be text')
            info = {'width_pt':page.width,'height_pt':page.height,'rotation_degrees':page.rotation,
                    'page_counts':{'images':len(page.images),'lines':len(page.lines),'curves':len(page.curves),'rectangles':len(page.rects)}}
        if str(pdf) not in sources:
            sources[str(pdf)] = sha256(pdf)
        validated.append({'id':identifier,'pdf':pdf,'page':number,'figure_number':figure_number,
                          'printed_page':entry.get('printed_page'),'bbox_pt':bbox,'caption_bbox_pt':caption_bbox,
                          'caption_text':caption,'expected_panels':panels,**info})
    executable = find_pdftoppm(args.pdftoppm)
    out.mkdir(parents=True)
    (out/'pages').mkdir()
    (out/'figures').mkdir()
    cache,records = {},[]
    for item in validated:
        key=(str(item['pdf']),item['page'])
        if key not in cache:
            stem=f"source-{len(cache)+1:03}-page-{item['page']:04}"
            cache[key]=render_page(item['pdf'],item['page'],args.dpi,out/'pages'/stem,item['width_pt'],item['height_pt'],executable)
        full,full_path=cache[key]
        left,top,right,bottom=item['bbox_pt']
        sx,sy=full.width/item['width_pt'],full.height/item['height_pt']
        bbox_px=[math.floor(left*sx),math.floor(top*sy),math.ceil(right*sx),math.ceil(bottom*sy)]
        crop=full.crop(bbox_px)
        figure_dir=out/'figures'/item['id']
        figure_dir.mkdir()
        image_path=figure_dir/'figure.png'
        crop.save(image_path,dpi=(args.dpi,args.dpi))
        crop_preview=figure_dir/'crop-preview.png'
        thumbnail(crop,crop_preview)
        annotated=full.copy()
        ImageDraw.Draw(annotated).rectangle(bbox_px,outline=(220,35,45),width=max(2,round(args.dpi/60)))
        source_preview=figure_dir/'source-preview.png'
        thumbnail(annotated,source_preview)
        strips={'left':(0,0,min(2,crop.width),crop.height),'right':(max(0,crop.width-2),0,crop.width,crop.height),
                'top':(0,0,crop.width,min(2,crop.height)),'bottom':(0,max(0,crop.height-2),crop.width,crop.height)}
        edge_ink={name:sum(crop.crop(box).convert('L').histogram()[:235]) for name,box in strips.items()}
        cb=item['caption_bbox_pt']
        caption_included=bool(cb and left<=cb[0] and top<=cb[1] and right>=cb[2] and bottom>=cb[3])
        records.append({k:v for k,v in item.items() if k!='pdf'} | {
            'source_pdf':str(item['pdf']),'source_sha256':sources[str(item['pdf'])],
            'coordinate_system':'PDF points, local top-left of rotated MediaBox: [left,top,right,bottom]',
            'dpi':args.dpi,'bbox_px':bbox_px,'image_size_px':list(crop.size),
            'image':image_path.relative_to(out).as_posix(),'image_sha256':sha256(image_path),'full_page':full_path.relative_to(out).as_posix(),
            'source_preview':source_preview.relative_to(out).as_posix(),'crop_preview':crop_preview.relative_to(out).as_posix(),
            'caption_included_in_crop':caption_included,
            'edge_ink_pixels_in_2px_border':edge_ink,
            'edge_check_warning':'A blank border does not establish completeness; an entire omitted panel would evade this check.',
            'verified':False,'visual_review':{'status':'pending','reviewer':None,'reviewed_at':None,
                'checks':{'full_page_and_crop_viewed':False,'all_expected_panels':False,'axes_and_ticks':False,
                          'shared_and_independent_legends':False,'labels_and_annotations':False,
                          'figure_number_and_complete_caption_traceable':False,'no_clipped_content':False}}
        })
    manifest={'schema_version':1,'created_at':datetime.now(timezone.utc).isoformat(),
              'method':'Full-page Poppler rendering followed by plan-driven Pillow cropping; no embedded image extraction.',
              'scope':'Material preparation only; not end-to-end paper-PPT acceptance.',
              'verified':False,'figures':records}
    write_json(out/'manifest.json',manifest)
    return {'output_directory':str(out),'figures_extracted':len(records),'visual_review':'pending'}


def main(argv=None):
    parser=argparse.ArgumentParser(description=__doc__)
    commands=parser.add_subparsers(dest='command',required=True)
    inspect=commands.add_parser('inspect',help='Inventory page dimensions, objects and candidate caption lines')
    inspect.add_argument('pdf');inspect.add_argument('--pages',help='1-based list/ranges; omitted means every page')
    inspect.add_argument('--output',required=True)
    render=commands.add_parser('render',help='Render complete PDF pages (MediaBox)')
    render.add_argument('pdf');render.add_argument('--pages',required=True)
    extract=commands.add_parser('extract',help='Extract complete figure regions from a JSON plan; always pending visual review')
    extract.add_argument('--plan',required=True)
    for sub in [render,extract]:
        sub.add_argument('--out-dir',required=True,help='New output directory; existing directories are not overwritten')
        sub.add_argument('--dpi',type=int,default=300)
        sub.add_argument('--pdftoppm',help='Poppler executable or its bin directory; otherwise PPT2405_POPPLER/PATH/runtime discovery')
    args=parser.parse_args(argv)
    try:
        if hasattr(args,'dpi') and not 36 <= args.dpi <= 1200:
            raise InputError('dpi must be between 36 and 1200')
        import pdfplumber
        if args.command in ('render','extract'):
            import PIL
        result={'inspect':inspect_command,'render':render_command,'extract':extract_command}[args.command](args,pdfplumber)
        print(json.dumps(result,ensure_ascii=False))
        return 0
    except (InputError, OSError, ValueError, TypeError, subprocess.TimeoutExpired) as error:
        print(f'ERROR: {error}',file=sys.stderr)
        return 2
    except ImportError as error:
        print(f'ERROR: dependency missing ({error}). Locate the bundled Python/pdfplumber/Pillow runtime; do not silently install packages.',file=sys.stderr)
        return 2


if __name__=='__main__':
    if hasattr(sys.stdout,'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    raise SystemExit(main())
