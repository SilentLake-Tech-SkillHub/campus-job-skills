"""Deterministic Excalidraw subset renderer for this repo's static flow sources.
Supports text, rectangle, ellipse, line and arrow; rejects unsupported geometry.
No screenshot composition: PNG is generated entirely from the source elements.
"""
import json, math, hashlib, argparse, os
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, PngImagePlugin
FONT=os.environ.get('FLOW_FONT', '/System/Library/Fonts/Supplemental/Arial Unicode.ttf')

def render(path):
 raw=path.read_bytes(); data=json.loads(raw)
 elements=[e for e in data['elements'] if not e.get('isDeleted')]
 assert all(e['type'] in {'text','rectangle','ellipse','line','arrow'} and not e.get('angle') for e in elements)
 lowx=min(e['x'] for e in elements)-35; lowy=min(e['y'] for e in elements)-35
 maxx=max(e['x']+e['width'] for e in elements)+35; maxy=max(e['y']+e['height'] for e in elements)+35
 im=Image.new('RGB',(math.ceil(maxx-lowx),math.ceil(maxy-lowy)),data.get('appState',{}).get('viewBackgroundColor','#ffffff'))
 d=ImageDraw.Draw(im)
 for e in elements:
  x,y=e['x']-lowx,e['y']-lowy;w,h=e['width'],e['height'];c=e.get('strokeColor','#1e1e1e');bg=e.get('backgroundColor','transparent');fill=None if bg=='transparent' else bg
  if e['type']=='text':
   font=ImageFont.truetype(FONT,round(e['fontSize'])); lh=e.get('lineHeight',1.25)*e['fontSize']
   for i,line in enumerate(e['text'].split('\n')):
    length=d.textlength(line,font=font);align=e.get('textAlign','left')
    tx=x+(w-length)/2 if align=='center' else x+w-length if align=='right' else x
    d.text((tx,y+i*lh),line,font=font,fill=c,anchor='lt')
  elif e['type']=='rectangle':
   d.rounded_rectangle((x,y,x+w,y+h),radius=12 if e.get('roundness') else 0,fill=fill,outline=c,width=round(e.get('strokeWidth',2)))
  elif e['type']=='ellipse':d.ellipse((x,y,x+w,y+h),fill=fill,outline=c,width=round(e.get('strokeWidth',2)))
  else:
   pts=[(x+a,y+b) for a,b in e['points']];d.line(pts,fill=c,width=round(e.get('strokeWidth',2)))
   if e.get('endArrowhead'):
    px,py=pts[-1];ax,ay=pts[-2];theta=math.atan2(py-ay,px-ax);size=14
    tip=[(px-size*math.cos(theta-off),py-size*math.sin(theta-off)) for off in [-.45,.45]]
    d.line([tip[0],(px,py),tip[1]],fill=c,width=2)
 meta=PngImagePlugin.PngInfo();meta.add_text('excalidraw_sha256',hashlib.sha256(raw).hexdigest());meta.add_text('renderer','source-only deterministic flow subset v1')
 im.save(path.with_suffix('.png'),pnginfo=meta)

if __name__=='__main__':
 parser=argparse.ArgumentParser();parser.add_argument('sources',nargs='+',type=Path);args=parser.parse_args()
 for p in args.sources:render(p)
