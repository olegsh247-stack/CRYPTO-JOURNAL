import struct
CP='cp1251'
def u16(s): return s.encode('utf-16-le')
# ---------- MS-OVBA compression ----------
def _chunk(chunk):
    body=bytearray(); i=0; n=len(chunk)
    while i<n:
        fp=len(body); body.append(0); flag=0
        for bit in range(8):
            if i>=n: break
            bc=max((i-1).bit_length() if i>1 else 0,4)
            maxlen=(0xFFFF>>bc)+3
            bl=0;bo=0
            c0=chunk[i]
            for j in range(i-1,-1,-1):
                if chunk[j]!=c0: continue
                l=0
                while i+l<n and l<maxlen and chunk[j+l]==chunk[i+l]: l+=1
                if l>bl: bl,bo=l,i-j
                if bl==maxlen: break
            if bl>=3:
                tok=((bo-1)<<(16-bc))|(bl-3)
                body+=struct.pack('<H',tok); flag|=1<<bit; i+=bl
            else:
                body.append(c0); i+=1
        body[fp]=flag
    if len(body)+2<=4098:
        return struct.pack('<H',((len(body)+2-3)&0xFFF)|0x3000|0x8000)+bytes(body)
    assert len(chunk)==4096
    return struct.pack('<H',0x3FFF)+bytes(chunk)
def compress(data):
    out=bytearray(b'\x01'); pos=0
    while pos<len(data):
        out+=_chunk(data[pos:pos+4096]); pos+=4096
    return bytes(out)
# ---------- dir stream ----------
def rec(i,payload): return struct.pack('<HI',i,len(payload))+payload
def build_dir(modules,codepage=1251):
    d=bytearray()
    d+=rec(0x01,struct.pack('<I',1))
    d+=rec(0x02,struct.pack('<I',0x409)); d+=rec(0x14,struct.pack('<I',0x409))
    d+=rec(0x03,struct.pack('<H',codepage))
    d+=rec(0x04,b'VBAProject')
    d+=rec(0x05,b'')+struct.pack('<HI',0x40,0)
    d+=rec(0x06,b'')+struct.pack('<HI',0x3D,0)
    d+=rec(0x07,struct.pack('<I',0)); d+=rec(0x08,struct.pack('<I',0))
    d+=struct.pack('<HIIH',0x09,4,0,0)
    d+=rec(0x0C,b'')+struct.pack('<HI',0x3C,0)
    refs=[('stdole',r'*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation'),
          ('Excel',r'*\G{00020813-0000-0000-C000-000000000046}#1.9#0#C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE#Microsoft Excel 16.0 Object Library'),
          ('Office',r'*\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL#Microsoft Office 16.0 Object Library')]
    for nm,lib in refs:
        d+=rec(0x16,nm.encode(CP))+struct.pack('<HI',0x3E,len(u16(nm)))+u16(nm)
        lb=lib.encode(CP)
        body=struct.pack('<I',len(lb))+lb+struct.pack('<IH',0,0)
        d+=rec(0x0D,body)
    d+=rec(0x0F,struct.pack('<H',len(modules)))
    d+=rec(0x13,struct.pack('<H',0xFFFF))
    for name,kind in modules:
        nb=name.encode(CP)
        d+=rec(0x19,nb)
        d+=rec(0x47,u16(name))
        d+=rec(0x1A,nb)+struct.pack('<HI',0x32,len(u16(name)))+u16(name)
        d+=rec(0x1C,b'')+struct.pack('<HI',0x48,0)
        d+=rec(0x31,struct.pack('<I',0))
        d+=rec(0x1E,struct.pack('<I',0))
        d+=rec(0x2C,struct.pack('<H',0xFFFF))
        d+=struct.pack('<HI',0x22 if kind=='doc' else 0x21,0)
        d+=struct.pack('<HI',0x2B,0)
    d+=struct.pack('<HI',0x10,0)
    return compress(bytes(d))
# ---------- module source ----------
def module_stream(name,kind,code=''):
    attrs='Attribute VB_Name = "%s"\r\n'%name
    if kind=='doc':
        base='{00020819-0000-0000-C000-000000000046}' if name=='ThisWorkbook' else '{00020820-0000-0000-C000-000000000046}'
        attrs+=('Attribute VB_Base = "0%s"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\n'
                'Attribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\nAttribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = True\r\n')%base
    code=code.replace('\r\n','\n').replace('\n','\r\n')
    return compress((attrs+code).encode(CP))
# ---------- minimal CFB writer ----------
FREE=0xFFFFFFFF; END=0xFFFFFFFE; FATSECT=0xFFFFFFFD
class E:
    def __init__(s,name,typ,data=None):
        s.name=name;s.typ=typ;s.data=data;s.kids=[];s.left=s.right=s.child=FREE;s.start=END;s.size=0;s.sid=0
def _bst(nodes):
    if not nodes: return FREE
    nodes=sorted(nodes,key=lambda e:(len(e.name),e.name.upper()))
    def b(l):
        if not l: return FREE
        m=len(l)//2; n=l[m]; n.left=b(l[:m]); n.right=b(l[m+1:]); return n.sid
    return b(nodes)
def build_cfb(root):
    ents=[root]
    def walk(e):
        for k in e.kids: ents.append(k)
        for k in e.kids:
            if k.typ==1: walk(k)
    walk(root)
    for i,e in enumerate(ents): e.sid=i
    for e in ents:
        if e.kids: e.child=_bst(e.kids)
    # mini stream
    mini=bytearray(); minifat=[]
    for e in ents:
        if e.typ==2:
            e.size=len(e.data)
            if e.size<4096:
                e.start=len(mini)//64 if e.size else END
                ns=(e.size+63)//64
                for k in range(ns): minifat.append(e.start+k+1 if k<ns-1 else END)
                mini+=e.data+b'\0'*(ns*64-e.size)
    root.size=len(mini)
    sectors=[]  # list of bytes
    fat=[]
    def alloc(data):
        ns=(len(data)+511)//512
        first=len(sectors)
        for k in range(ns):
            sectors.append(data[k*512:(k+1)*512].ljust(512,b'\0'))
            fat.append(first+k+1 if k<ns-1 else END)
        return first if ns else END
    for e in ents:
        if e.typ==2 and e.size>=4096: e.start=alloc(e.data)
    mf=b''.join(struct.pack('<I',x) for x in minifat)
    mf_first=alloc(mf) if mf else END
    mf_n=(len(mf)+511)//512
    # directory
    def dentry(e):
        nm=u16(e.name)+b'\0\0'
        typ=e.typ
        return (nm.ljust(64,b'\0')+struct.pack('<HBB',len(nm),typ,1)+struct.pack('<III',e.left,e.right,e.child)
                +b'\0'*16+struct.pack('<I',0)+b'\0'*16+struct.pack('<IQ',e.start if e.size or e.typ!=2 else END,e.size)[0:4]+struct.pack('<Q',e.size)[0:8]) if False else None
    dirb=bytearray()
    for e in ents:
        nm=u16(e.name)+b'\0\0'
        dirb+=nm.ljust(64,b'\0')+struct.pack('<HBB',len(nm),e.typ,1)+struct.pack('<III',e.left,e.right,e.child)
        dirb+=b'\0'*16+struct.pack('<I',0)+b'\0'*16
    # starts need root start (ministream) known later -> second pass
    dir_first=len(sectors)
    ndir=(len(ents)*128+511)//512
    for k in range(ndir):
        sectors.append(b'');fat.append(dir_first+k+1 if k<ndir-1 else END)
    mini_first=END
    if mini:
        mini_first=len(sectors); nm_=(len(mini)+511)//512
        for k in range(nm_):
            sectors.append(mini[k*512:(k+1)*512].ljust(512,b'\0')); fat.append(mini_first+k+1 if k<nm_-1 else END)
    root.start=mini_first
    dirb=bytearray()
    for e in ents:
        nm=u16(e.name)+b'\0\0'
        dirb+=nm.ljust(64,b'\0')+struct.pack('<HBB',len(nm),e.typ,1)+struct.pack('<III',e.left,e.right,e.child)
        dirb+=b'\0'*16+struct.pack('<I',0)+b'\0'*16+struct.pack('<I',e.start)+struct.pack('<Q',e.size)
    dirb=bytes(dirb).ljust(ndir*512,b'\0')
    # pad empty entries properly
    pad=bytearray()
    for _ in range(len(ents),ndir*4):
        pad+=b'\0'*64+struct.pack('<HBB',0,0,0)+struct.pack('<III',FREE,FREE,FREE)+b'\0'*(16+4+16)+struct.pack('<I',0)+struct.pack('<Q',0)
    dirb=dirb[:len(ents)*128]+bytes(pad)
    for k in range(ndir): sectors[dir_first+k]=dirb[k*512:(k+1)*512]
    # FAT sectors
    nfat=1
    while True:
        total=len(sectors)+nfat
        if (total+127)//128<=nfat: break
        nfat+=1
    fat_first=len(sectors)
    for k in range(nfat): fat.append(FATSECT); sectors.append(b'')
    fat+= [FREE]*(nfat*128-len(fat))
    for k in range(nfat):
        sectors[fat_first+k]=b''.join(struct.pack('<I',x) for x in fat[k*128:(k+1)*128])
    hdr=bytearray(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'+b'\0'*16)
    hdr+=struct.pack('<HHHHH',0x3E,3,0xFFFE,9,6)+b'\0'*6
    hdr+=struct.pack('<IIIIIIIII',0,nfat,dir_first,0,4096,mf_first,mf_n,END,0)
    difat=[fat_first+k for k in range(nfat)]+[FREE]*(109-nfat)
    hdr+=b''.join(struct.pack('<I',x) for x in difat)
    assert len(hdr)==512
    return bytes(hdr)+b''.join(sectors)

def build_vba_project(sheet_codenames,code):
    mods=[('ThisWorkbook','doc')]+[(n,'doc') for n in sheet_codenames]+[('modTrade','std')]
    proj=['ID="{5B6B8A44-5A1D-4C2F-8B7E-1C0F6B3C2A10}"']
    for n,k in mods:
        proj.append(('Document=%s/&H00000000' if k=='doc' else 'Module=%s')%n)
    proj+=['Name="VBAProject"','HelpContextID="0"','VersionCompatible32="393222000"','',
           '[Host Extender Info]','&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000','','[Workspace]']
    for n,k in mods: proj.append('%s=0, 0, 0, 0, C'%n)
    project=('\r\n'.join(proj)+'\r\n').encode(CP)
    wm=b''.join(n.encode(CP)+b'\0'+u16(n)+b'\0\0' for n,k in mods)+b'\0\0'
    root=E('Root Entry',5)
    root.kids=[E('PROJECT',2,project),E('PROJECTwm',2,wm)]
    v=E('VBA',1)
    v.kids=[E('_VBA_PROJECT',2,bytes([0xCC,0x61,0xFF,0xFF,0,0,0])),E('dir',2,build_dir(mods))]
    for n,k in mods:
        v.kids.append(E(n,2,module_stream(n,k,code if n=='modTrade' else '')))
    root.kids.append(v)
    return build_cfb(root)
