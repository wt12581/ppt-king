"""Portable reference/component catalog. Python standard library only."""
import argparse,base64,hashlib,html,json,re,shutil,sys,uuid,zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
ROOT=Path(__file__).resolve().parents[1]
CAT=ROOT/'assets/catalog.json'
NS={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
def read(p):return json.loads(Path(p).read_text(encoding='utf-8-sig'))
def emit(x):print(json.dumps(x,ensure_ascii=False,indent=2))
def digest(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def valid_id(s):
    if not re.fullmatch(r'[a-z0-9]+(?:-[a-z0-9]+)*',s) or len(s)>64:raise ValueError('ID must be lowercase ASCII, digits and hyphens; max 64 characters')
def contained(root,p):
    f=(root/p).resolve()
    if not f.is_relative_to(root.resolve()):raise ValueError('Path escapes resource directory')
    return f
def native_counts(p):
    result={'slides':0,'pictures':0,'textboxes':0,'groups':0,'freeforms':0,'attached_connectors':0}
    with zipfile.ZipFile(p) as z:
        if z.testzip():raise ValueError('Corrupt PPTX')
        for n in z.namelist():
            if re.fullmatch(r'ppt/slides/slide\d+\.xml',n):
                t=ET.fromstring(z.read(n));result['slides']+=1
                result['pictures']+=len(t.findall('.//p:pic',NS))
                result['groups']+=len(t.findall('.//p:grpSp',NS))
                result['freeforms']+=len(t.findall('.//a:custGeom',NS))
                result['textboxes']+=sum(bool(v.findall('.//a:t',NS)) for v in t.findall('.//p:sp',NS))
                result['attached_connectors']+=sum(v.find('.//a:stCxn',NS) is not None or v.find('.//a:endCxn',NS) is not None for v in t.findall('.//p:cxnSp',NS))
    return result
def image_check(p):
    # Verify signature; visual/content inspection remains the agent's responsibility.
    h=Path(p).read_bytes()[:64]
    if not (h.startswith(b'\x89PNG\r\n\x1a\n') or h.startswith(b'\xff\xd8\xff') or h[:4]==b'RIFF' and h[8:12]==b'WEBP'):raise ValueError('Reference must be PNG, JPEG or WebP')
def save_catalog(c):
    tmp=CAT.with_name('catalog-'+uuid.uuid4().hex+'.tmp')
    tmp.write_text(json.dumps(c,ensure_ascii=False,indent=2)+'\n',encoding='utf8');tmp.replace(CAT)
def register(collection,item,files):
    valid_id(item['id']);c=read(CAT)
    if any(v['id']==item['id'] for v in c[collection]):raise ValueError('ID already exists; choose a new version ID')
    dest=ROOT/'assets'/collection/item['id']
    if dest.exists():raise ValueError('Resource destination already exists')
    dest.mkdir(parents=True)
    try:
        item['files']={}
        for source,name in files:
            target=contained(dest,name);target.parent.mkdir(parents=True,exist_ok=True)
            shutil.copy2(source,target)
            rel=target.relative_to(ROOT).as_posix();item['files'][rel]=digest(target)
        c[collection].append(item);save_catalog(c)
    except Exception:
        # Leave partial folder inspectable; catalog is only committed after successful copies.
        raise
    emit(item)
def main():
    parser=argparse.ArgumentParser(description=__doc__)
    sub=parser.add_subparsers(dest='cmd',required=True)
    sub.add_parser('list-styles')
    p=sub.add_parser('list-components');p.add_argument('--query',default='')
    p=sub.add_parser('show-style');p.add_argument('id')
    p=sub.add_parser('browse');p.add_argument('--out',required=True)
    sub.add_parser('check')
    p=sub.add_parser('new-project');p.add_argument('--out',required=True);p.add_argument('--style',required=True);p.add_argument('--purpose',default='group-meeting')
    p=sub.add_parser('add-style');p.add_argument('--id',required=True);p.add_argument('--name',required=True);p.add_argument('--source-dir',required=True);p.add_argument('--tags',default='')
    p=sub.add_parser('add-component');p.add_argument('--id',required=True);p.add_argument('--name',required=True);p.add_argument('--pptx',required=True);p.add_argument('--preview',required=True);p.add_argument('--category',required=True);p.add_argument('--tags',default='');p.add_argument('--allow-pictures',action='store_true')
    args=parser.parse_args();c=read(CAT)
    if args.cmd=='list-styles':
        emit([{'id':v['id'],'name':v['name'],'type':v['type'],'reference_images':len(v.get('images',[])),'description':v['description']} for v in c['styles']]);return
    if args.cmd=='list-components':
        q=args.query.casefold();emit([v for v in c['components'] if q in json.dumps({k:v.get(k) for k in ['id','name','category','tags','description']},ensure_ascii=False).casefold()]);return
    if args.cmd=='show-style':
        v=next(v for v in c['styles'] if v['id']==args.id);emit(v);print(contained(ROOT,v['brief']).read_text(encoding='utf8'));return
    if args.cmd=='check':
        issues=[];ids=set()
        for collection in ['styles','components']:
            for v in c[collection]:
                valid_id(v['id']);key=(collection,v['id'])
                if key in ids:issues.append('duplicate '+str(key))
                ids.add(key)
                for rel,sha in v.get('files',{}).items():
                    p=contained(ROOT,rel)
                    if not p.is_file() or digest(p)!=sha:issues.append('missing/changed '+rel)
                if collection=='styles':
                    if not contained(ROOT,v['brief']).is_file():issues.append('missing brief '+v['id'])
                    for im in v.get('images',[]):image_check(contained(ROOT,im['file']))
                else:
                    n=native_counts(contained(ROOT,v['pptx']))
                    if n['slides']!=1 or n['attached_connectors'] or n['pictures'] and v.get('native_only',True):issues.append('component invariant '+v['id'])
                    image_check(contained(ROOT,v['preview']))
        emit({'styles':len(c['styles']),'components':len(c['components']),'issues':issues,'passed':not issues})
        if issues:sys.exit(1)
        return
    if args.cmd=='new-project':
        next(v for v in c['styles'] if v['id']==args.style);presets=read(ROOT/'references/presets.json')
        if args.purpose not in presets:raise ValueError('Unknown purpose preset')
        out=Path(args.out).resolve();out.mkdir(parents=True,exist_ok=True);p=out/'project.json'
        if p.exists():raise ValueError('Project already exists; use resume')
        d={'schema_version':3,'workflow':'complete-design-api-rebuild','implementation_mode':'lightweight-mixed','generation_tool':None,'render_tool':None,'visual_review':{'method':'per-slide-and-object-comparison','no_simplified_substitutes':True},'stage':'outline','style_id':args.style,'purpose':args.purpose,'sources':[],'canvas':None,'typography':{'font':None,'minimum_body_pt':None},'institution':None,'logo':None,'slide_count':None,'arrows':{'independent':True,'attached':False},'slides':[]}
        p.write_text(json.dumps(d,ensure_ascii=False,indent=2)+'\n',encoding='utf8');emit({'project':str(p)});return
    if args.cmd=='add-style':
        folder=Path(args.source_dir).resolve();brief=folder/'brief.md';ims=read(folder/'images.json')
        if not brief.is_file() or not brief.read_text(encoding='utf8').strip() or not ims:raise ValueError('Nonempty brief and actual images are required')
        files=[(brief,'brief.md')];refs=[];seen=set()
        for im in ims:
            f=contained(folder,im['file']);image_check(f);name=f.name
            if name in seen:raise ValueError('Duplicate reference image filename')
            seen.add(name);files.append((f,name));refs.append({'file':f'assets/styles/{args.id}/{name}','role':im['role']})
        if (folder/'reference.pptx').is_file():native_counts(folder/'reference.pptx');files.append((folder/'reference.pptx','reference.pptx'))
        item={'id':args.id,'name':args.name,'type':'image-reference','description':brief.read_text(encoding='utf8').splitlines()[0].lstrip('# '),'tags':args.tags.split(',') if args.tags else [],'brief':f'assets/styles/{args.id}/brief.md','images':refs,'provenance':'user-added; inspect sources and images before first use'}
        register('styles',item,files);return
    if args.cmd=='add-component':
        n=native_counts(args.pptx);image_check(args.preview)
        if n['slides']!=1 or n['attached_connectors']:raise ValueError('Require one slide and no attached connectors')
        if n['pictures'] and not args.allow_pictures:raise ValueError('Raster pictures found; not a native-only component')
        item={'id':args.id,'name':args.name,'category':args.category,'tags':args.tags.split(',') if args.tags else [],'description':'用户添加；需在实际导出PPTX中检查视觉和编辑效果','native_only':not n['pictures'],'counts':n,'pptx':f'assets/components/{args.id}/component.pptx','preview':f'assets/components/{args.id}/preview.png','parameters':['position','size','color','text if present']}
        register('components',item,[(Path(args.pptx),'component.pptx'),(Path(args.preview),'preview.png')]);return
    if args.cmd=='browse':
        out=Path(args.out).resolve();out.parent.mkdir(parents=True,exist_ok=True)
        cards=[]
        def image_uri(rel):
            p=contained(ROOT,rel);mime={'.png':'image/png','.jpg':'image/jpeg','.jpeg':'image/jpeg','.webp':'image/webp'}[p.suffix.lower()]
            return 'data:'+mime+';base64,'+base64.b64encode(p.read_bytes()).decode()
        for v in c['styles']:
            ims=v.get('images',[]);pictures=''.join('<img loading="lazy" src="'+image_uri(i['file'])+'" alt="'+html.escape(i['role'])+'">' for i in ims)
            cards.append('<section><h2>'+html.escape(v['name'])+' <code>'+v['id']+'</code></h2><p>'+html.escape(v['description'])+'</p><p>'+('实际图片参考 · '+str(len(ims))+'张' if ims else '文字方向预设 · 无参考图片')+'</p><div class="grid">'+pictures+'</div></section>')
        comps=''.join('<article><img loading="lazy" src="'+image_uri(v['preview'])+'" alt="'+html.escape(v['name'])+'"><h3>'+html.escape(v['name'])+'</h3><code>'+v['id']+'</code><p>'+html.escape(v.get('description',''))+'</p></article>' for v in c['components'])
        body='<!doctype html><html lang="zh-CN"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>逐页生图PPT资源库</title><style>body{font-family:"Microsoft YaHei",Arial,sans-serif;margin:24px;color:#243448;background:#f4f7fa}section{margin:24px 0}h1{font-size:26px}h2{font-size:20px}h3{font-size:16px}code{font-size:14px;overflow-wrap:anywhere}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:16px}.grid img{width:100%;background:white}.components{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:16px}article{background:white;padding:16px}article img{width:100%;height:160px;object-fit:contain}p{line-height:1.7}dialog{width:90vw;border:0}dialog img{width:100%;max-height:85vh;object-fit:contain}img{cursor:zoom-in}</style><h1>逐页生图与可编辑PPT · 资源库</h1><p>图片参考用于视觉设计，机构、数据、模型和字体不自动继承。点击图片放大。</p>'+''.join(cards)+'<h2>精选原生组件</h2><div class="components">'+comps+'</div><dialog><button>关闭</button><img alt="放大预览"></dialog><script>const d=document.querySelector("dialog");document.querySelectorAll("img").forEach(i=>{if(i.parentElement!==d)i.onclick=()=>{d.querySelector("img").src=i.src;d.showModal()}});d.querySelector("button").onclick=()=>d.close();</script></html>'
        out.write_text(body,encoding='utf8');emit({'browse':str(out)})
if __name__=='__main__':
    try:main()
    except (ValueError,StopIteration,OSError,KeyError,zipfile.BadZipFile) as e:print('ERROR: '+str(e),file=sys.stderr);sys.exit(2)
