import sys,bpy,json,re,math
bpy.ops.wm.open_mainfile(filepath=sys.argv[1],load_ui=False,use_scripts=False)
NG=-0.7
def mw(o):
    m=o.matrix_basis.copy()
    if o.parent: m=mw(o.parent)@o.matrix_parent_inverse@m
    return m
def akey(col):
    if col.startswith('01 Aeration') or col.startswith('02 Aeration'): return 'AW-AER'
    m=re.match(r'FST ([A-L])',col)
    if m: return 'AW-FST-'+m.group(1)
    if col.startswith('Pipes'): return 'PIPES'
    return col.split(' | ')[0]
def hull(pts):
    pts=sorted(set(pts))
    if len(pts)<3: return pts
    def cr(o,a,b): return (a[0]-o[0])*(b[1]-o[1])-(a[1]-o[1])*(b[0]-o[0])
    lo=[];up=[]
    for p in pts:
        while len(lo)>=2 and cr(lo[-2],lo[-1],p)<=0: lo.pop()
        lo.append(p)
    for p in reversed(pts):
        while len(up)>=2 and cr(up[-2],up[-1],p)<=0: up.pop()
        up.append(p)
    return lo[:-1]+up[:-1]
def ap(h):
    a=0;p=0
    for i in range(len(h)):
        x1,y1=h[i];x2,y2=h[(i+1)%len(h)]
        a+=x1*y2-x2*y1; p+=math.hypot(x2-x1,y2-y1)
    return abs(a)/2,p
main=bpy.data.scenes['Alex West - combined coordination']
SKIP=re.compile(r'pile|display stub|rail|post|stair|tread|water|WL \+|equipment|cabinet|crane|door|glazing|ladder',re.I)
grp={}; allpts=[]; zmin={}
for o in main.objects:
    if o.type!='MESH': continue
    cols=[c.name for c in o.users_collection if not c.name.startswith('AW-AER Status')]
    if not cols: continue
    k=akey(cols[0]); M=mw(o)
    vs=[M@v.co for v in o.data.vertices]
    if not vs: continue
    step=max(1,len(vs)//400)
    allpts+= [(round(v.x,2),round(v.y,2)) for v in vs[::step]]
    if k=='PIPES' or SKIP.search(o.name) or 'provisional attachments' in cols[0]: continue
    zb=min(v.z for v in vs)
    zmin[k]=min(zmin.get(k,9e9),zb)
    if zb<NG:
        grp.setdefault(k,[]).extend((round(v.x,2),round(v.y,2)) for v in vs[::step])
out={}
for k,pts in grp.items():
    h=hull(pts); a,p=ap(h); out[k]=dict(area=a,perim=p,zmin=zmin[k],cx=sum(x for x,_ in h)/len(h),cy=sum(y for _,y in h)/len(h))
for k in zmin:
    if k not in out: out[k]=dict(area=0,perim=0,zmin=zmin[k])
h=hull(allpts); a,p=ap(h); out['_SITE_ALL']=dict(area=a,perim=p,zmin=0)
json.dump(out,open(sys.argv[2],'w'),indent=0)
for k,v in sorted(out.items()): print(k, round(v['area']), round(v['perim']), round(v['zmin'],2))
