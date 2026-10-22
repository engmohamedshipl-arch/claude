import sys,bpy,json,mathutils
bpy.ops.wm.open_mainfile(filepath=sys.argv[1],load_ui=False,use_scripts=False)
def mw(o):
    return mathutils.Matrix.LocRotScale(o.location,o.rotation_euler,o.scale)
P=[o for o in bpy.data.collections['02 Aeration piles - display stubs'].objects]
xs=[o.location.x for o in P]; ys=[o.location.y for o in P]
print('pile loc range',min(xs),max(xs),min(ys),max(ys))
rows=[]
for o in P:
    rows.append(dict(name=o.name,x=o.location.x,y=o.location.y,z=o.location.z,status=o.get('Tracking_Status'),zone=int(o.get('Tracking_Zone')),pile=o.get('Drawing_Pile_ID'),exe=o.get('Tracking_Execution_Date'),data=o.get('Tracking_Data_Date')))
json.dump(rows,open(sys.argv[2],'w'))
# structures: compute model-frame footprint centroid per top-level collection for reference
