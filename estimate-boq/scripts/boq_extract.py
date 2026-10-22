import sys,bpy,bmesh,json,mathutils,math
bpy.ops.wm.open_mainfile(filepath=sys.argv[1],load_ui=False,use_scripts=False)
def mw(o):
    m=o.matrix_basis.copy()
    if o.parent: m=mw(o.parent)@o.matrix_parent_inverse@m
    return m
main=bpy.data.scenes['Alex West - combined coordination']
rows=[]
KEEP=('Source','Category','Element_ID','Asset_ID','Status','Maturity','Tank_ID','Placement','Coordination_status','Source_status','Approval_status','Elevation_basis','Basin_ID','Unit_ID','Thickener_ID','Layer_ID','tag','diameter','material','system','Component_type','Qualification','Review_status')
for o in main.objects:
    if o.type!='MESH': continue
    cols=[c.name for c in o.users_collection]
    M=mw(o)
    bm=bmesh.new(); bm.from_mesh(o.data); bm.transform(M)
    nonman=sum(1 for e in bm.edges if not e.is_manifold)
    vol=abs(bm.calc_volume(signed=True)) if bm.faces else 0.0
    up=dn=vert=0.0
    for f in bm.faces:
        a=f.calc_area(); nz=f.normal.z
        if nz>0.7: up+=a
        elif nz<-0.7: dn+=a
        elif abs(nz)<0.3: vert+=a
    if bm.verts:
        xs=[v.co.x for v in bm.verts]; ys=[v.co.y for v in bm.verts]; zs=[v.co.z for v in bm.verts]
        bb=(min(xs),max(xs),min(ys),max(ys),min(zs),max(zs))
    else: bb=(0,)*6
    ld=[abs(o.dimensions[i]) for i in range(3)]
    rows.append(dict(name=o.name,cols=cols,nv=len(bm.verts),nf=len(bm.faces),nonman=nonman,vol=vol,up=up,dn=dn,vert=vert,bb=bb,ldim=ld,
                     props={k:(o[k] if isinstance(o[k],(int,float,str)) else str(o[k])) for k in KEEP if k in o.keys()}))
    bm.free()
json.dump(rows,open(sys.argv[2],'w'))
print(len(rows))
