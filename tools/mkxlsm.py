import openpyxl, zipfile, re, sys
sys.path.insert(0,'.')
import vbabuild as v
SRC="/home/claude/crypto-journal/archive/CryptoJournal. Test m2 (исходник без макросов).xlsx"
OUT="/home/claude/crypto-journal/CryptoJournal. Test m2.xlsm"
import tempfile,os
TMP=os.path.join(tempfile.gettempdir(),"tmp_cn.xlsx")
wb=openpyxl.load_workbook(SRC)
wb.code_name="ThisWorkbook"
names=[]
for i,ws in enumerate(wb.worksheets,1):
    ws.sheet_properties.codeName="Sheet%d"%i; names.append("Sheet%d"%i)
inp_idx=wb.sheetnames.index("Ввод")+1
wb.save(TMP)
code=open('/home/claude/crypto-journal/macros/modTrade.bas',encoding='utf-8').read()
vba=v.build_vba_project(names,code,{'Sheet%d'%inp_idx:open('/home/claude/crypto-journal/macros/modInputSheet.bas',encoding='utf-8').read()})
zin=zipfile.ZipFile(TMP)
files={n:zin.read(n) for n in zin.namelist()}
ct=files['[Content_Types].xml'].decode()
ct=ct.replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml','application/vnd.ms-excel.sheet.macroEnabled.main+xml')
ct=ct.replace('</Types>','<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
files['[Content_Types].xml']=ct.encode()
rel=files['xl/_rels/workbook.xml.rels'].decode()
rel=rel.replace('</Relationships>','<Relationship Id="rIdVba1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>')
files['xl/_rels/workbook.xml.rels']=rel.encode()
files['xl/vbaProject.bin']=vba
sheet='xl/worksheets/sheet%d.xml'%inp_idx
sx=files[sheet].decode()
assert '<drawing' not in sx
tail=re.search(r'<(legacyDrawing|tableParts|extLst)[ >/]',sx)
ins='<drawing r:id="rIdDr1"/>'
if 'xmlns:r=' not in sx[:600]:
    sx=sx.replace('<worksheet ','<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',1)
sx=sx[:tail.start()]+ins+sx[tail.start():] if tail else sx.replace('</worksheet>',ins+'</worksheet>')
files[sheet]=sx.encode()
rp='xl/worksheets/_rels/sheet%d.xml.rels'%inp_idx
assert rp not in files
files[rp]=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>').encode()
def btn(id_,name,macro,c1,c2):
    # cells: from (col c1,row 15) to (col c2,row 17) 0-based -> F16:F17 = col5 rows15..16 ; to col6,row17
    return f'''<xdr:twoCellAnchor editAs="absolute"><xdr:from><xdr:col>{c1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>{c2}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>17</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp macro="[0]!{macro}" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{id_}" name="{name}"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="1000"/></a:srgbClr></a:solidFill><a:ln><a:noFill/></a:ln></xdr:spPr><xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ru-RU"/></a:p></xdr:txBody></xdr:sp><xdr:clientData fLocksWithSheet="0"/></xdr:twoCellAnchor>'''
dr='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'+btn(2,'btnWrite','WriteEntry',5,6)+btn(3,'btnClear','ClearInput',6,7)+'</xdr:wsDr>'
files['xl/drawings/drawing1.xml']=dr.encode()
with zipfile.ZipFile(OUT,'w',zipfile.ZIP_DEFLATED) as z:
    order=['[Content_Types].xml']+[n for n in files if n!='[Content_Types].xml']
    for n in order: z.writestr(n,files[n])
print("ok",OUT, inp_idx)
