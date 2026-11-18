#!/usr/bin/env python3
"""Compose ppt2405 from a hash-bound JSON plan. No COM, UI, or network is used."""
from __future__ import annotations
import argparse
import copy
import hashlib
import io
import json
import math
import os
import posixpath
import re
import sys
import uuid
from pathlib import Path
from urllib.parse import unquote
from xml.etree import ElementTree as E
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED

try:
    from PIL import Image, ImageFont
except ImportError:
    raise SystemExit('Pillow is required; use the bundled Python runtime or install Pillow.')

NS = {
    'p':'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm':'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'a14':'http://schemas.microsoft.com/office/drawing/2010/main',
    'a16':'http://schemas.microsoft.com/office/drawing/2014/main',
    'p14':'http://schemas.microsoft.com/office/powerpoint/2010/main',
    'mc':'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'rel':'http://schemas.openxmlformats.org/package/2006/relationships',
    'ct':'http://schemas.openxmlformats.org/package/2006/content-types',
}
for prefix,uri in NS.items(): E.register_namespace(prefix,uri)
REL=NS['r']+'/'
EMU_PT=12700
SCALE=4


class ComposeError(ValueError): pass


def q(prefix,name): return '{'+NS[prefix]+'}'+name
def node(prefix,name,attrs=None,children=()):
    out=E.Element(q(prefix,name),attrs or {});out.extend(children);return out
def xml(root):
    # The OPC reader used by Office/Artifact Tool requires unprefixed Types.
    # Namespace-equivalent ct:Types parses as XML but is rejected by that reader.
    if root.tag==q('ct','Types'):
        E.register_namespace('',NS['ct'])
        try:return E.tostring(root,encoding='utf-8',xml_declaration=True)
        finally:E.register_namespace('ct',NS['ct'])
    return E.tostring(root,encoding='utf-8',xml_declaration=True)
def sha(data): return hashlib.sha256(data).hexdigest()
def local(tag): return tag.rsplit('}',1)[-1]
def fail(message): raise ComposeError(message)
def check_keys(value,allowed,label):
    if not isinstance(value,dict):fail(f'{label} must be an object')
    unknown=set(value)-set(allowed)
    if unknown:fail(f'Unknown fields in {label}: {sorted(unknown)}')
def text_of(shape): return ''.join(t.text or '' for t in shape.findall('.//a:t',NS)+shape.findall('.//m:t',NS))
def props(shape):
    for path in ('p:nvSpPr/p:cNvPr','p:nvPicPr/p:cNvPr','p:nvGraphicFramePr/p:cNvPr'):
        found=shape.find(path,NS)
        if found is not None:return found
    return None
def named(shape):
    pr=props(shape);return '' if pr is None else pr.get('name','')
def rel_path(part): return posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
def resolve(part,target):
    target=unquote(target.split('#')[0])
    return target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(part),target))


class Fonts:
    """Actual glyph advances from locally installed files, oversampled 4x."""
    def __init__(self,config,base,report=None):
        self.config=config;self.base=base;self.cache={};self.used={};self.report=report
    def get(self,family,size,bold=False):
        item=self.config.get(family)
        if item is None:
            win=Path(os.environ.get('WINDIR','C:/Windows'))/'Fonts'
            options={
                '微软雅黑':{'regular':str(win/'msyh.ttc'),'bold':str(win/'msyhbd.ttc')},
                'Microsoft YaHei':{'regular':str(win/'msyh.ttc'),'bold':str(win/'msyhbd.ttc')},
                'Cambria Math':{'regular':str(win/'cambria.ttc'),'index':1},
                'Times New Roman':{'regular':str(win/'times.ttf'),'bold':str(win/'timesbd.ttf')},
            }
            item=options.get(family)
        if item is None:fail(f'No local font file configured for {family!r}; provide font_files.')
        if isinstance(item,str):item={'regular':item}
        check_keys(item,{'regular','bold','index'},f'font_files[{family}]')
        filepath=Path(item.get('bold') if bold and item.get('bold') else item.get('regular',''))
        if not filepath.is_absolute():filepath=self.base/filepath
        if not filepath.is_file():fail(f'Font file unavailable for {family!r}: {filepath}')
        index=int(item.get('index',0));key=(str(filepath.resolve()),index,round(size*SCALE))
        if key not in self.cache:
            self.cache[key]=ImageFont.truetype(str(filepath),key[2],index=index)
        if str(filepath.resolve()) not in self.used:
            self.used[str(filepath.resolve())]={'family':family,'index':index,'actual_family':self.cache[key].getname()[0],'sha256':sha(filepath.read_bytes())}
            if self.report is not None:self.report['fonts']=list(self.used.values())
        return self.cache[key]
    def measure(self,text,family,size,bold=False):
        f=self.get(family,size,bold);box=f.getbbox(text)
        # Glyph ink height, not font-wide ascender+descender: the latter is not
        # a baseline distance and overstates two-line CJK text substantially.
        return f.getlength(text)/SCALE,(box[3]-box[1])/SCALE


def source_style(body):
    base=body.find('.//a:rPr',NS)
    if base is None:base=body.find('.//a:defRPr',NS)
    if base is None:base=node('a','rPr',{'sz':'1800'})
    base=copy.deepcopy(base);base.tag=q('a','rPr')
    return base


def style_family(style,default='微软雅黑'):
    ea=style.find('a:ea',NS);latin=style.find('a:latin',NS)
    f=ea.get('typeface') if ea is not None else latin.get('typeface') if latin is not None else default
    return default if not f or f.startswith('+') else f


def run_style(base,size,family=None,bold=None,color=None):
    out=copy.deepcopy(base);out.tag=q('a','rPr');out.set('sz',str(round(size*100)))
    if bold is not None:out.set('b','1' if bold else '0')
    if family:
        for kind in ('latin','ea','cs'):
            child=out.find('a:'+kind,NS)
            if child is None:child=node('a',kind);out.append(child)
            child.set('typeface',family)
    if color:
        if not re.fullmatch(r'#?[0-9A-Fa-f]{6}',color):fail(f'Invalid color {color}')
        for c in list(out):
            if local(c.tag) in {'solidFill','noFill','gradFill','pattFill','blipFill','grpFill'}:out.remove(c)
        out.insert(0,node('a','solidFill',children=[node('a','srgbClr',{'val':color.lstrip('#').upper()})]))
    # CT_TextCharacterProperties is ordered: outline, fill, effects, underline,
    # fonts, hyperlinks, extensions. Office can silently ignore a late fill.
    groups=[('ln',),('noFill','solidFill','gradFill','blipFill','pattFill','grpFill'),
            ('effectLst','effectDag'),('highlight',),('uLnTx','uLn'),('uFillTx','uFill'),
            ('latin',),('ea',),('cs',),('sym',),('hlinkClick',),('hlinkMouseOver',),('extLst',)]
    order={tag:i for i,group in enumerate(groups) for tag in group}
    out[:]=sorted(out,key=lambda child:order.get(local(child.tag),len(groups)))
    return out


def math_run(value,size,base,plain=False):
    r=node('m','r')
    if plain:r.append(node('m','rPr',children=[node('m','sty',{q('m','val'):'p'})]))
    r.append(run_style(base,size,'Cambria Math',False));t=node('m','t');t.text=str(value)
    if str(value)!=str(value).strip():t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
    r.append(t);return r


def math_nodes(expr,size,base):
    """A deliberately small, explicit AST; this does not interpret arbitrary LaTeX."""
    if isinstance(expr,(str,int,float)):
        if '\\' in str(expr):fail('Math strings are literal Unicode, not LaTeX; use the documented AST.')
        return [math_run(expr,size,base)]
    if isinstance(expr,list):return [n for part in expr for n in math_nodes(part,size,base)]
    if not isinstance(expr,dict) or len(expr)!=1:fail('Math AST node must contain exactly one supported operator')
    kind,value=next(iter(expr.items()))
    def cell(name,item):return node('m',name,children=math_nodes(item,size,base))
    def control():return node('m','ctrlPr',children=[run_style(base,size,'Cambria Math',False)])
    if kind=='text':return [math_run(value,size,base,True)]
    if kind in {'sub','sup','subsup'}:
        expected=3 if kind=='subsup' else 2
        if not isinstance(value,list) or len(value)!=expected:fail(f'{kind} expects {expected} elements')
        tag={'sub':'sSub','sup':'sSup','subsup':'sSubSup'}[kind]
        children=[node('m',tag+'Pr',children=[control()]),cell('e',value[0])]
        if kind in {'sub','subsup'}:children.append(cell('sub',value[1]))
        if kind in {'sup','subsup'}:children.append(cell('sup',value[-1]))
        return [node('m',tag,children=children)]
    if kind=='frac':
        if not isinstance(value,list) or len(value)!=2:fail('frac expects [numerator, denominator]')
        return [node('m','f',children=[node('m','fPr',children=[control()]),cell('num',value[0]),cell('den',value[1])])]
    if kind=='sqrt':
        return [node('m','rad',children=[node('m','radPr',children=[node('m','degHide',{q('m','val'):'1'}),control()]),node('m','deg'),cell('e',value)])]
    if kind=='hat':
        return [node('m','acc',children=[node('m','accPr',children=[node('m','chr',{q('m','val'):'̂'}),control()]),cell('e',value)])]
    if kind=='delim':
        check_keys(value,{'left','right','body'},'math delim')
        pr=node('m','dPr',children=[node('m','begChr',{q('m','val'):value.get('left','(')}),node('m','endChr',{q('m','val'):value.get('right',')')}),control()])
        return [node('m','d',children=[pr,cell('e',value['body'])])]
    if kind=='nary':
        check_keys(value,{'symbol','sub','sup','body'},'math nary')
        pr=node('m','naryPr',children=[node('m','chr',{q('m','val'):value.get('symbol','∑')}),node('m','limLoc',{q('m','val'):'undOvr'}),control()])
        for side in ('sub','sup'):
            if side not in value:pr.append(node('m',side+'Hide',{q('m','val'):'1'}))
        return [node('m','nary',children=[pr,cell('sub',value.get('sub','')),cell('sup',value.get('sup','')),cell('e',value['body'])])]
    fail(f'Unsupported math operator: {kind}')


def math_box(expr,size,fonts):
    """Conservative structural estimate from real Cambria Math glyph advances."""
    if isinstance(expr,(str,int,float)):return fonts.measure(str(expr),'Cambria Math',size)
    if isinstance(expr,list):
        boxes=[math_box(x,size,fonts) for x in expr];return sum(x[0] for x in boxes),max([x[1] for x in boxes]+[size])
    kind,value=next(iter(expr.items()))
    if kind=='text':return fonts.measure(str(value),'Cambria Math',size)
    if kind in {'sub','sup','subsup'}:
        main=math_box(value[0],size,fonts);minor=[math_box(x,size*.72,fonts) for x in value[1:]]
        return main[0]+max(x[0] for x in minor),main[1]+max(x[1] for x in minor)*(.7 if kind=='subsup' else .42)
    if kind=='frac':
        a,b=[math_box(x,size,fonts) for x in value];return max(a[0],b[0])+size*.35,a[1]+b[1]+size*.25
    if kind in {'sqrt','hat'}:
        a,b=math_box(value,size,fonts);return a+(size*.7 if kind=='sqrt' else size*.1),b+size*.25
    if kind=='delim':
        a,b=math_box(value['body'],size,fonts);return a+size,b
    if kind=='nary':
        body=math_box(value['body'],size,fonts);lo=math_box(value.get('sub',''),size*.65,fonts);hi=math_box(value.get('sup',''),size*.65,fonts)
        return body[0]+max(size,lo[0],hi[0])+size*.3,max(body[1],size+lo[1]+hi[1])
    fail(f'Unsupported math operator: {kind}')


def normalize_text(value):
    if isinstance(value,str):value={'text':value}
    check_keys(value,{'text','paragraphs','font_size_pt','min_font_size_pt','font','bold','color','align','vertical_align'},'text value')
    if ('text' in value)==('paragraphs' in value):fail('Text value requires exactly one of text or paragraphs')
    paragraphs=[[{'text':p}] for p in str(value['text']).split('\n')] if 'text' in value else value['paragraphs']
    if not isinstance(paragraphs,list) or not paragraphs:fail('paragraphs must be a nonempty array')
    result=[]
    for paragraph in paragraphs:
        if not isinstance(paragraph,list):fail('Each paragraph must be an array of runs')
        runs=[]
        for r in paragraph:
            if isinstance(r,str):r={'text':r}
            check_keys(r,{'text','math','bold','color','font'},'inline run')
            if ('text' in r)==('math' in r):fail('Each inline run requires exactly one of text or math')
            runs.append(r)
        result.append(runs)
    return value,result


def measured_lines(paragraphs,size,family,bold,width,fonts,line_spacing_pt):
    max_width=0;count=0;all_heights=[]
    for paragraph in paragraphs:
        line_w=0;line_h=0;heights=[]
        def flush():
            nonlocal line_w,line_h,max_width,count
            heights.append(line_h or fonts.measure('国',family,size,bold)[1]);max_width=max(max_width,line_w);count+=1;line_w=0;line_h=0
        for span in paragraph:
            if 'math' in span:
                tokens=[math_box(span['math'],size,fonts)]
            else:
                s=str(span['text']);f=span.get('font',family);b=span.get('bold',bold)
                # CJK characters can wrap individually; English words stay together.
                words=re.findall(r'\n|[\u2e80-\u9fff\uf900-\ufaff]|[^\s\u2e80-\u9fff\uf900-\ufaff]+|[ \t]+',s)
                tokens=[]
                for word in words:
                    if word=='\n':tokens.append(None)
                    else:
                        box=fonts.measure(word,f,size,b)
                        if box[0]>width and len(word)>1:
                            tokens.extend(fonts.measure(ch,f,size,b) for ch in word)
                        else:tokens.append(box)
            for box in tokens:
                if box is None:flush();continue
                w,h=box
                if line_w and line_w+w>width:flush()
                line_w+=w;line_h=max(line_h,h)
        flush()
        all_heights.extend(heights)
    # Plain text uses the native-Office-verified 1.16 * font-size point spacing.
    # Tall math/inline constructs raise the SAME point value that is serialized.
    actual_spacing=math.ceil(max([line_spacing_pt]+all_heights)*100)/100
    return max_width,actual_spacing*count,count,actual_spacing


def geometry(shape,catalog=None):
    xfrm=shape.find('p:spPr/a:xfrm',NS)
    if xfrm is None:xfrm=shape.find('p:xfrm',NS)
    if xfrm is not None:
        off=xfrm.find('a:off',NS);ext=xfrm.find('a:ext',NS)
        if off is not None and ext is not None:return tuple(int(v) for v in (off.get('x'),off.get('y'),ext.get('cx'),ext.get('cy')))
    if catalog:return tuple(round(catalog[k]*9525) for k in ('x','y','w','h'))
    fail(f'Missing explicit geometry: {named(shape)}')


def text_dimensions(body,width,height):
    bp=body.find('a:bodyPr',NS);attrs={} if bp is None else bp.attrib
    # DrawingML defaults: 0.1 inch horizontal and 0.05 inch vertical.
    left=int(attrs.get('lIns','91440'));right=int(attrs.get('rIns','91440'))
    top=int(attrs.get('tIns','45720'));bottom=int(attrs.get('bIns','45720'))
    return (width-left-right)/EMU_PT,(height-top-bottom)/EMU_PT


def fit_text(body,paragraphs,config,slot,width,height,fonts,label,report):
    base=source_style(body);family=config.get('font',slot.get('font',style_family(base)))
    bold=config.get('bold',base.get('b')=='1')
    normal=float(slot.get('fontSizePt',float(base.get('sz','1800'))/100))
    lo=float(slot.get('minFontSizePt',normal));hi=float(slot.get('maxFontSizePt',normal))
    maximum=float(config.get('font_size_pt',min(hi,normal)));minimum=float(config.get('min_font_size_pt',lo))
    if not 0<lo<=minimum<=maximum<=hi:fail(f'Font range outside template bounds for {label}: allowed {lo}–{hi} pt, requested {minimum}–{maximum} pt')
    aw,ah=text_dimensions(body,width,height)
    size=maximum;chosen=None
    while size>=minimum-0.001:
        mw,mh,count,spacing_pt=measured_lines(paragraphs,size,family,bold,aw,fonts,size*1.16)
        if mw<=aw+.15 and mh<=ah+.15:chosen=(size,mw,mh,count,spacing_pt);break
        size=round(size-.5,3)
    if chosen is None:
        report['overflows'].append({'slot':label,'minimum_pt':minimum,'available_pt':[round(aw,2),round(ah,2)],'measured_pt':[round(mw,2),round(mh,2)],'action':'Shorten text or split across a second template slide.'})
        fail(f'Text overflow in {label} at minimum {minimum:g} pt; shorten or split the slide.')
    size,mw,mh,count,spacing_pt=chosen
    report['text_fits'].append({'slot':label,'font':family,'font_size_pt':size,'lines':count,'line_spacing_pt':spacing_pt,'paragraph_before_pt':0,'paragraph_after_pt':0,'available_pt':[round(aw,2),round(ah,2)],'measured_pt':[round(mw,2),round(mh,2)],'measurement':'local glyph advances/ink; fixed point line spacing written identically; structural math is conservatively estimated'})
    return base,family,bold,size,spacing_pt


def write_text(body,paragraphs,config,slot,width,height,fonts,label,report):
    alignments={'left':'l','center':'ctr','right':'r','justify':'just'}
    anchors={'top':'t','middle':'ctr','bottom':'b'}
    for field,allowed in (('align',alignments),('vertical_align',anchors)):
        if field in config and (not isinstance(config[field],str) or config[field] not in allowed):fail(f'{label}: {field} must be one of {", ".join(allowed)}')
    base,family,bold,size,spacing_pt=fit_text(body,paragraphs,config,slot,width,height,fonts,label,report)
    report['text_fits'][-1].update({'align':config.get('align'),'vertical_align':config.get('vertical_align')})
    original=body.find('a:p/a:pPr',NS)
    default_ppr=copy.deepcopy(original) if original is not None else node('a','pPr')
    if 'align' in config:default_ppr.set('algn',alignments[config['align']])
    # No hidden auto-shrink: actual explicit font size is recorded and written.
    bp=body.find('a:bodyPr',NS)
    if bp is None:bp=node('a','bodyPr');body.insert(0,bp)
    if 'vertical_align' in config:bp.set('anchor',anchors[config['vertical_align']])
    for old in list(bp):
        if local(old.tag) in {'normAutofit','spAutoFit','noAutofit'}:bp.remove(old)
    bp.append(node('a','noAutofit'))
    for old in list(body.findall('a:p',NS)):body.remove(old)
    for paragraph in paragraphs:
        pp=copy.deepcopy(default_ppr)
        for child in list(pp):
            if local(child.tag) in {'lnSpc','spcBef','spcAft'}:pp.remove(child)
        for index,(kind,value) in enumerate([('lnSpc',round(spacing_pt*100)),('spcBef',0),('spcAft',0)]):
            pp.insert(index,node('a',kind,children=[node('a','spcPts',{'val':str(value)})]))
        dr=pp.find('a:defRPr',NS)
        if dr is not None:pp.remove(dr)
        dr=run_style(base,size,family,bold,config.get('color'));dr.tag=q('a','defRPr');pp.append(dr)
        p=node('a','p',children=[pp])
        for span in paragraph:
            style=run_style(base,size,span.get('font',family),span.get('bold',bold),span.get('color',config.get('color')))
            if 'math' in span:
                p.append(node('a14','m',children=[node('m','oMath',children=math_nodes(span['math'],size,style))]))
            else:
                # Explicit newlines in runs are represented as native DrawingML line breaks.
                parts=str(span['text']).split('\n')
                for i,part in enumerate(parts):
                    if i:p.append(node('a','br'))
                    t=node('a','t');t.text=part
                    if part!=part.strip():t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
                    p.append(node('a','r',children=[copy.deepcopy(style),t]))
        ep=run_style(base,size,family,bold);ep.tag=q('a','endParaRPr');p.append(ep);body.append(p)


def fill_shape(shape,value,slot,fonts,label,report,as_math=False):
    body=shape.find('p:txBody',NS)
    if body is None:fail(f'Slot {label} is not a text shape')
    if as_math:
        check_keys(value,{'expr','font_size_pt','min_font_size_pt','align','vertical_align'},'math value')
        if 'expr' not in value:fail('math value requires expr')
        # Build first, so unsupported syntax fails explicitly before fit estimation.
        math_nodes(value['expr'],slot['fontSizePt'],source_style(body))
        config={k:v for k,v in value.items() if k!='expr'};config['font']='Cambria Math'
        paragraphs=[[{'math':value['expr']}]]
    else:config,paragraphs=normalize_text(value)
    _,_,w,h=geometry(shape,slot)
    write_text(body,paragraphs,config,slot,w,h,fonts,label,report)


def new_rid(rels,prefix='rIdPpt2405'):
    used={x.get('Id') for x in rels};i=1
    while prefix+str(i) in used:i+=1
    return prefix+str(i)


def image_shape(original,config,slot,rels,parts,content_types,base,label,report):
    check_keys(config,{'path','source','alt','fit'},'image value')
    if config.get('fit','contain')!='contain':fail(f'Only contain is supported for {label}; crop/cover is forbidden.')
    path=Path(config['path']);path=path if path.is_absolute() else base/path
    data=path.read_bytes()
    with Image.open(io.BytesIO(data)) as im:
        fmt=im.format;pw,ph=im.size;orientation=im.getexif().get(274,1)
        if orientation not in (None,1):fail(f'{label}: normalize EXIF orientation before use; pixel bounds must be unambiguous.')
        if getattr(im,'n_frames',1)!=1:fail(f'{label}: animated/multipage images are not supported')
    formats={'PNG':('png','image/png'),'JPEG':('jpg','image/jpeg')}
    if fmt not in formats:fail(f'{label}: use PNG or JPEG for deterministic complete-image embedding')
    ext,ctype=formats[fmt];media='ppt/media/ppt2405-'+sha(data)+'.'+ext;parts[media]=data
    if not any(x.get('Extension')==ext for x in content_types):content_types.append(node('ct','Default',{'Extension':ext,'ContentType':ctype}))
    rid=new_rid(rels);rels.append(node('rel','Relationship',{'Id':rid,'Type':REL+'image','Target':'/'+media}))
    x,y,w,h=geometry(original,slot);scale=min(w/pw,h/ph);iw=round(pw*scale);ih=round(ph*scale);ix=x+(w-iw)//2;iy=y+(h-ih)//2
    oldpr=props(original);nv=copy.deepcopy(oldpr);nv.set('descr',config.get('alt',slot.get('name',label)))
    nvpr=original.find('p:nvSpPr/p:nvPr',NS)
    if nvpr is None:nvpr=original.find('p:nvPicPr/p:nvPr',NS)
    nvpr=copy.deepcopy(nvpr) if nvpr is not None else node('p','nvPr')
    pic=node('p','pic',children=[
        node('p','nvPicPr',children=[nv,node('p','cNvPicPr',children=[node('a','picLocks',{'noChangeAspect':'1'})]),nvpr]),
        node('p','blipFill',children=[node('a','blip',{q('r','embed'):rid}),node('a','stretch',children=[node('a','fillRect')])]),
        node('p','spPr',children=[node('a','xfrm',children=[node('a','off',{'x':str(ix),'y':str(iy)}),node('a','ext',{'cx':str(iw),'cy':str(ih)})]),node('a','prstGeom',{'prst':'rect'},[node('a','avLst')]),node('a','ln',children=[node('a','noFill')])]),
    ])
    report['images'].append({'slot':label,'path':str(path.resolve()),'sha256':sha(data),'pixels':[pw,ph],'slot_emu':[x,y,w,h],'placed_emu':[ix,iy,iw,ih],'contain':True,'cropped':False,'source':config.get('source'),'source_missing':not bool(config.get('source'))})
    return pic


def fill_table(frame,config,fonts,label,report):
    check_keys(config,{'rows','font_size_pt','min_font_size_pt','source'},'table value')
    table=frame.find('.//a:tbl',NS)
    if table is None:fail(f'{label} is not a native table')
    rows=config.get('rows')
    if not isinstance(rows,list) or len(rows)<2:fail(f'{label}: rows must include a header and at least one data row')
    widths=[int(x.get('w')) for x in table.findall('a:tblGrid/a:gridCol',NS)];cols=len(widths)
    if any(not isinstance(r,list) or len(r)!=cols for r in rows):fail(f'{label}: expected exactly {cols} cells in every row')
    old=table.findall('a:tr',NS)
    if not old:fail(f'{label}: no source row styles')
    if any(any(k in c.attrib for k in ('gridSpan','rowSpan','hMerge','vMerge')) for c in table.findall('.//a:tc',NS)):fail(f'{label}: merged cells are outside the supported table schema')
    total_h=sum(int(r.get('h')) for r in old);row_h=total_h//len(rows)
    for r in old:table.remove(r)
    for ri,values in enumerate(rows):
        prototype=old[0] if ri==0 else old[min(ri,len(old)-1)]
        row=copy.deepcopy(prototype);height=row_h if ri<len(rows)-1 else total_h-row_h*(len(rows)-1);row.set('h',str(height))
        cells=row.findall('a:tc',NS)
        for ci,value in enumerate(values):
            if isinstance(value,dict):
                check_keys(value,{'text','bold','color'},f'{label}: cell [{ri},{ci}]')
                if 'text' not in value:fail(f'{label}: styled cells require text')
                cfg=dict(value)
                if 'bold' in cfg and not isinstance(cfg['bold'],bool):fail(f'{label}: cell bold must be true or false')
                if 'color' in cfg and (not isinstance(cfg['color'],str) or not re.fullmatch(r'#?[0-9A-Fa-f]{6}',cfg['color'])):fail(f'{label}: invalid cell color')
            else:cfg={'text':value}
            if not isinstance(cfg['text'],(str,int,float)) or isinstance(cfg['text'],bool):fail(f'{label}: cell text must be a string or number')
            cfg['text']=str(cfg['text'])
            body=cells[ci].find('a:txBody',NS);style=source_style(body);original_pt=float(style.get('sz','1800'))/100
            tcpr=cells[ci].find('a:tcPr',NS);bp=body.find('a:bodyPr',NS)
            if tcpr is not None and bp is not None:
                for tcattr,bpattr in [('marL','lIns'),('marR','rIns'),('marT','tIns'),('marB','bIns')]:
                    if tcattr in tcpr.attrib:bp.set(bpattr,tcpr.get(tcattr))
            slot={'font':style_family(style),'fontSizePt':original_pt,'minFontSizePt':original_pt,'maxFontSizePt':original_pt}
            # Table style size is fixed unless the catalog-independent source cell allows explicit shrink.
            for k in ('font_size_pt','min_font_size_pt'):
                if k in config:cfg[k]=config[k]
            normalized,paras=normalize_text(cfg)
            write_text(body,paras,normalized,slot,widths[ci],height,fonts,f'{label}[{ri},{ci}]',report)
        table.append(row)
    report['tables'].append({'slot':label,'rows':len(rows),'columns':cols,'native':True,'source':config.get('source'),'source_missing':not bool(config.get('source'))})


def notes_body(root):
    for sp in root.findall('.//p:sp',NS):
        ph=sp.find('p:nvSpPr/p:nvPr/p:ph',NS)
        if ph is not None and ph.get('type')=='body':
            body=sp.find('p:txBody',NS)
            if body is not None:return body
    fail('Source notes have no native body placeholder')


def note_lines(lines):
    if isinstance(lines,str):lines=[lines]
    if not isinstance(lines,list) or any(not isinstance(x,str) for x in lines):fail('notes must be a string or array of strings')
    return [paragraph for line in lines for paragraph in line.split('\n')]


def update_notes(root,lines):
    lines=note_lines(lines);body=notes_body(root)
    for old in list(body.findall('a:p',NS)):body.remove(old)
    for line in lines:
        t=node('a','t');t.text=line
        body.append(node('a','p',children=[node('a','r',children=[t])]))


def notes_source_slot(name,slot):
    return name.startswith('source-') or name=='source' or slot.get('deliveryRole')=='notes-source'


def source_note(value):
    return value if isinstance(value,str) else json.dumps(value,ensure_ascii=False,sort_keys=True)


def legacy_source_text(value):
    _,paragraphs=normalize_text(value)
    if any('math' in run for paragraph in paragraphs for run in paragraph):fail('Source-only slots accept plain citation text, not math')
    return '\n'.join(''.join(str(run['text']) for run in paragraph) for paragraph in paragraphs).strip()


def omit_content(tree,target):
    """Override inherited placeholder content without deleting its matching ph."""
    ph=target.find('p:nvSpPr/p:nvPr/p:ph',NS)
    if ph is None:
        tree.remove(target)
        return 'removed slide-level object'
    body=target.find('p:txBody',NS)
    if body is not None:
        for child in list(body.findall('a:p',NS)):body.remove(child)
        # An explicit invisible glyph overrides layout prompt inheritance in renderers.
        t=node('a','t');t.text='\u200b'
        body.append(node('a','p',children=[node('a','r',children=[node('a','rPr',{'sz':'100'}),t])]))
    sppr=target.find('p:spPr',NS)
    if sppr is not None:
        for child in list(sppr):
            if local(child.tag) in {'solidFill','gradFill','pattFill','noFill','blipFill','ln'}:sppr.remove(child)
        sppr.append(node('a','noFill'));sppr.append(node('a','ln',{'w':'0'},[node('a','noFill')]))
    return 'cleared placeholder content, fill, and border; fixed layout artwork retained'


def stamp_ids(root,seed):
    for i,x in enumerate(root.findall('.//a16:creationId',NS)):
        x.set('id','{'+str(uuid.uuid5(uuid.NAMESPACE_URL,seed+':'+str(i))).upper()+'}')
    for i,x in enumerate(root.findall('.//p14:creationId',NS)):
        x.set('val',str(int(sha((seed+':p14:'+str(i)).encode())[:8],16)))


def validate_relationships(parts):
    for name,data in parts.items():
        if name.endswith(('.xml','.rels')):E.fromstring(data)
        if not name.endswith('.rels'):continue
        source=posixpath.join(posixpath.dirname(posixpath.dirname(name)),posixpath.basename(name)[:-5])
        if name=='_rels/.rels':source=''
        for rel in E.fromstring(data):
            if rel.get('TargetMode')=='External':continue
            destination=resolve(source,rel.get('Target',''))
            if destination not in parts:fail(f'Broken package relationship: {name} -> {destination}')


def compose(args,report):
    report.setdefault('sources_to_notes',[])
    plan_path=Path(args.plan).resolve();plan=json.loads(plan_path.read_text(encoding='utf-8-sig'));base=plan_path.parent
    check_keys(plan,{'version','template_sha256','catalog_sha256','mode','font_files','slides','source'},'plan')
    if plan.get('version')!=1:fail('Plan version must equal 1')
    if plan.get('mode','strict') not in {'strict','template'}:fail('mode must be strict or template')
    mode=plan.get('mode','strict');template=Path(args.template).resolve();catalog_path=Path(args.catalog).resolve();output=Path(args.output).resolve()
    if output in {template,catalog_path,plan_path}:fail('Output cannot overwrite a source, catalog, or plan')
    if output.suffix.lower()!='.pptx':fail('Output must have a .pptx extension')
    if output.exists() and not args.overwrite:fail('Output exists; use --overwrite to replace the generated file')
    template_data=template.read_bytes();catalog_data=catalog_path.read_bytes();template_sha=sha(template_data)
    if plan.get('template_sha256','').lower()!=template_sha:fail('Template SHA256 mismatch; bind the plan to the exact reviewed template')
    if plan.get('catalog_sha256') and plan['catalog_sha256'].lower()!=sha(catalog_data):fail('Catalog SHA256 mismatch')
    catalog=json.loads(catalog_data);by_id={str(x['id']):x for x in catalog}
    entries=plan.get('slides')
    if not isinstance(entries,list) or not entries:fail('slides must be a nonempty array')
    fonts=Fonts(plan.get('font_files',{}),base,report)
    report.update({'mode':mode,'template':str(template),'template_sha256':template_sha,'catalog':str(catalog_path),'catalog_sha256':sha(catalog_data),'plan_sha256':sha(plan_path.read_bytes()),'source':plan.get('source'),'old_unreferenced_parts':'Retained; only selected copies occur in the presentation slide order.','limitations':['Glyph fit is conservative preflight, not a native PowerPoint layout engine. Render every final slide.','Math accepts the documented AST, not arbitrary LaTeX.','Native table column count and source style are retained; merged cells are unsupported.']})
    with ZipFile(io.BytesIO(template_data)) as z:parts={n:z.read(n) for n in z.namelist()}
    presentation=E.fromstring(parts['ppt/presentation.xml']);prels=E.fromstring(parts['ppt/_rels/presentation.xml.rels']);ct=E.fromstring(parts['[Content_Types].xml'])
    ordered=[];mapping={r.get('Id'):r.get('Target') for r in prels}
    old_list=presentation.find('p:sldIdLst',NS)
    if old_list is None:fail('Template has no slide list')
    for sid in old_list:ordered.append(resolve('ppt/presentation.xml',mapping[sid.get(q('r','id'))]))
    maxslide=max(int(re.search(r'slide(\d+)\.xml$',n).group(1)) for n in parts if re.fullmatch(r'ppt/slides/slide\d+\.xml',n))
    maxnotes=max([int(re.search(r'notesSlide(\d+)\.xml$',n).group(1)) for n in parts if re.fullmatch(r'ppt/notesSlides/notesSlide\d+\.xml',n)]+[0])
    maxid=max([int(x.get('id')) for x in old_list]+[255]);old_list.clear()
    for r in list(prels):
        if r.get('Type')==REL+'slide':prels.remove(r)
    generated=[]
    for order,entry in enumerate(entries,1):
        check_keys(entry,{'layout','text','math','images','tables','notes','source','omit'},f'slides[{order-1}]')
        layout=str(entry.get('layout',''));layout=layout.zfill(2) if layout.isdigit() else layout
        if layout not in by_id:fail(f'Unknown layout: {layout}')
        record=by_id[layout];source=ordered[int(record['slide'])-1];slide=E.fromstring(parts[source]);srels=E.fromstring(parts[rel_path(source)]);tree=slide.find('p:cSld/p:spTree',NS)
        if tree is None:fail('Missing slide shape tree')
        slot_catalog={s['name']:s for s in record['slots']};lookup={};kind_map={}
        for sname,slot in slot_catalog.items():
            full=f'ppt2405.{layout}.{sname}';matches=[x for x in tree if named(x)==full]
            if len(matches)!=1:fail(f'Template/catalog mismatch: expected one named slot {full}, found {len(matches)}')
            lookup[sname]=matches[0];lookup[full]=matches[0];kind_map[sname]='image' if slot.get('type')=='picture' else 'text'
        actual_tables=[x for x in tree if x.find('.//a:tbl',NS) is not None]
        # Slide-local branding stays editable and is not a research-image slot.
        actual_images=[x for x in tree if x.tag==q('p','pic') and not named(x).startswith('装饰.')]
        table_map={};image_map={k:v for k,v in lookup.items() if v.tag in {q('p','sp'),q('p','pic')} and (k.startswith('ppt2405.') and k.rsplit('.',1)[-1] in slot_catalog and slot_catalog[k.rsplit('.',1)[-1]].get('type')=='picture' or k in kind_map and kind_map[k]=='image')}
        for i,item in enumerate(actual_tables,1):
            table_map['table-'+str(i)]=item
            if named(item):table_map[named(item)]=item;table_map[named(item).rsplit('.',1)[-1]]=item
        for i,item in enumerate(actual_images,1):
            image_map['image-'+str(i)]=item
            if named(item):image_map[named(item)]=item;image_map[named(item).rsplit('.',1)[-1]]=item
        used=set();filled_names=set();automatic_notes=[]
        def preserve_source(kind,slotname,value):
            if not value:return
            line=f'来源 [{slotname}]：{source_note(value)}'
            if line not in automatic_notes:automatic_notes.append(line)
            report['sources_to_notes'].append({'slide':order,'layout':layout,'kind':kind,'slot':slotname,'source':value})
        preserve_source('slide','slide',entry.get('source'))
        omitted=entry.get('omit',[])
        if not isinstance(omitted,list) or any(not isinstance(x,str) for x in omitted):fail('omit must be an array of known slot names')
        all_objects={**lookup,**image_map,**table_map}
        for key in omitted:
            if key not in all_objects:fail(f'Unknown slot {key!r} for layout {layout} in omit')
            target=all_objects[key]
            if id(target) in used:fail(f'Duplicate assignments to slot {key}')
            used.add(id(target));short=named(target).rsplit('.',1)[-1] or key;filled_names.add(short)
            action=omit_content(tree,target)
            report['omitted'].append({'slide':order,'layout':layout,'slot':short,'action':action})
        for category in ('text','math','images','tables'):
            inputs=entry.get(category,{})
            if not isinstance(inputs,dict):fail(f'{category} must be an object')
            candidates=table_map if category=='tables' else image_map if category=='images' else lookup
            for key,value in inputs.items():
                if key not in candidates:fail(f'Unknown slot {key!r} for layout {layout} in {category}')
                target=candidates[key]
                if id(target) in used:fail(f'Duplicate assignments to slot {key}')
                used.add(id(target));short=named(target).rsplit('.',1)[-1] or key;filled_names.add(short);label=f'slide {order} / {layout}.{short}'
                meta=slot_catalog.get(short,{'name':short})
                if notes_source_slot(short,meta):
                    if category!='text':fail(f'Source-only slot {key} accepts text; provenance belongs in notes')
                    preserve_source('legacy-source-slot',short,legacy_source_text(value))
                    action=omit_content(tree,target)
                    report['omitted'].append({'slide':order,'layout':layout,'slot':short,'action':action,'automatic':True,'reason':'Source citation moved to notes'})
                    continue
                if category in {'text','math'}:
                    if meta.get('type')=='picture':fail(f'Picture slot {key} must be filled through images')
                    if meta.get('nativeMath') and category=='text':fail(f'Native math slot {key} must be filled through math')
                    fill_shape(target,value,meta,fonts,label,report,category=='math')
                    if not text_of(target).strip():report['unfilled'].append({'slide':order,'layout':layout,'slot':short,'reason':'Assigned empty text'})
                elif category=='images':
                    pic=image_shape(target,value,meta,srels,parts,ct,base,label,report);idx=list(tree).index(target);tree.remove(target);tree.insert(idx,pic)
                    preserve_source('image',short,value.get('source'))
                else:
                    fill_table(target,value,fonts,label,report)
                    preserve_source('table',short,value.get('source'))
        for sname,slot in slot_catalog.items():
            if notes_source_slot(sname,slot) and sname not in filled_names:
                target=lookup[sname];used.add(id(target));filled_names.add(sname)
                action=omit_content(tree,target)
                report['omitted'].append({'slide':order,'layout':layout,'slot':sname,'action':action,'automatic':True,'reason':'Source citations are notes-only'})
        for sname,slot in slot_catalog.items():
            required=slot.get('type')=='picture' or '[' in slot.get('content','') or slot.get('nativeMath') or slot.get('font')=='Cambria Math'
            if required and sname not in filled_names:report['unfilled'].append({'slide':order,'layout':layout,'slot':sname,'reason':'Template example/placeholder retained'})
        for i,target in enumerate(actual_tables,1):
            if id(target) not in used:report['unfilled'].append({'slide':order,'layout':layout,'slot':named(target) or f'table-{i}','reason':'Example native table retained'})
        for i,target in enumerate(actual_images,1):
            if id(target) not in used:report['unfilled'].append({'slide':order,'layout':layout,'slot':named(target) or f'image-{i}','reason':'Example slide image retained'})
        new_slide=f'ppt/slides/slide{maxslide+order}.xml';new_notes=f'ppt/notesSlides/notesSlide{maxnotes+order}.xml'
        seed=template_sha+':'+report['plan_sha256']+':'+str(order);stamp_ids(slide,seed)
        nrel=next((r for r in srels if r.get('Type')==REL+'notesSlide'),None)
        if nrel is not None:
            old_note=resolve(source,nrel.get('Target'));notes=E.fromstring(parts[old_note]);nrels=E.fromstring(parts[rel_path(old_note)])
            if 'notes' in entry:lines=note_lines(entry['notes'])
            elif mode=='strict':lines=[]
            else:lines=[''.join(t.text or '' for t in p.findall('.//a:t',NS)) for p in notes_body(notes).findall('a:p',NS)]
            for line in automatic_notes:
                if line not in lines:lines.append(line)
            if 'notes' in entry or mode=='strict' or automatic_notes:update_notes(notes,lines or [''])
            stamp_ids(notes,seed+':notes')
            for r in nrels:
                if r.get('Type')==REL+'slide':r.set('Target','/'+new_slide)
            parts[new_notes]=xml(notes);parts[rel_path(new_notes)]=xml(nrels);nrel.set('Target','/'+new_notes)
            ct.append(node('ct','Override',{'PartName':'/'+new_notes,'ContentType':'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'}))
        elif 'notes' in entry or automatic_notes:fail(f'Layout {layout} has no notes to preserve citations or speaker notes')
        parts[new_slide]=xml(slide);parts[rel_path(new_slide)]=xml(srels)
        ct.append(node('ct','Override',{'PartName':'/'+new_slide,'ContentType':'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'}))
        rid=new_rid(prels);prels.append(node('rel','Relationship',{'Id':rid,'Type':REL+'slide','Target':'/'+new_slide}));old_list.append(node('p','sldId',{'id':str(maxid+order),q('r','id'):rid}))
        generated.append({'order':order,'layout':layout,'source_part':source,'part':new_slide,'source':entry.get('source'),'filled_slots':sorted(filled_names)})
    report['slides']=generated;report['fonts']=list(fonts.used.values())
    if mode=='strict' and report['unfilled']:fail(f'{len(report["unfilled"])} unfilled template slots; see report and complete them before strict export.')
    if mode=='strict' and any(x.get('source_missing') for x in report['images']+report['tables']):fail('Evidence source provenance is required for every strict-mode image/table.')
    parts['ppt/presentation.xml']=xml(presentation);parts['ppt/_rels/presentation.xml.rels']=xml(prels)
    for x in ct:
        if x.get('PartName')=='/ppt/presentation.xml':x.set('ContentType','application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml')
    parts['[Content_Types].xml']=xml(ct)
    if 'docProps/app.xml' in parts:
        app=E.fromstring(parts['docProps/app.xml'])
        for x in app.iter():
            if local(x.tag)=='Slides':x.text=str(len(generated))
        parts['docProps/app.xml']=xml(app)
    validate_relationships(parts)
    output.parent.mkdir(parents=True,exist_ok=True);temp=output.with_name(output.name+'.tmp')
    try:
        with ZipFile(temp,'w',ZIP_DEFLATED,compresslevel=9) as z:
            for name in sorted(parts):
                info=ZipInfo(name,date_time=(1980,1,1,0,0,0));info.compress_type=ZIP_DEFLATED;z.writestr(info,parts[name])
        os.replace(temp,output)
    finally:
        if temp.exists():temp.unlink()
    report.update({'status':'success','output':str(output),'output_sha256':sha(output.read_bytes()),'slide_count':len(generated),'relationship_check':'All internal package relationships resolve.','render_review_required':True})


def main():
    home=Path(__file__).resolve().parents[1]
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--template',default=str(home/'assets/ppt2405.pptx'))
    parser.add_argument('--catalog',default=str(home/'references/layout-catalog.json'))
    parser.add_argument('--plan',required=True,help='JSON plan; relative image/font paths resolve beside this file')
    parser.add_argument('--output',required=True,help='New .pptx file')
    parser.add_argument('--report',help='JSON audit report (default: OUTPUT.report.json)')
    parser.add_argument('--overwrite',action='store_true',help='Replace an existing generated output only after a complete successful build')
    args=parser.parse_args();report={'status':'failed','errors':[],'text_fits':[],'overflows':[],'images':[],'tables':[],'unfilled':[],'omitted':[],'fonts':[]}
    report_path=Path(args.report or (args.output+'.report.json')).resolve()
    if report_path in {Path(p).resolve() for p in (args.template,args.catalog,args.plan,args.output)}:
        print('ppt2405: Report path must differ from source, catalog, plan, and output.',file=sys.stderr);return 2
    code=0
    try:compose(args,report)
    except (ComposeError,OSError,KeyError,ValueError,E.ParseError) as error:
        report['errors'].append(str(error));print(f'ppt2405: {error}',file=sys.stderr);code=2
    report_path.parent.mkdir(parents=True,exist_ok=True)
    report_path.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    if not code:print(json.dumps({'output':report['output'],'slides':report['slide_count'],'unfilled':len(report['unfilled']),'report':str(report_path)},ensure_ascii=False))
    return code


if __name__=='__main__':sys.exit(main())
