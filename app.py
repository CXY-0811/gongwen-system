import os, sqlite3, io, re, json
from datetime import date, datetime
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, abort
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

app=Flask(__name__); app.secret_key=os.getenv('SECRET_KEY','change-this-secret-key')
DB=os.getenv('DATABASE_PATH',os.path.join(os.path.dirname(__file__),'data.db'))
SAMPLE={'name':'新民國小跑道整修工程','office':'徐英豪建築師事務所','year':115,'prefix':'一一五所新環字第','recipient':'彰化縣田中鎮新民國民小學','speed':'普通件','security':'普通','address':'臺中市西屯區西屯路三段 159-72 號 5 樓','phone':'04-24631919','fax':'04-24631509','email':'yh6726.mail@msa.hinet.net','contact':'黃先生'}

def conn(): c=sqlite3.connect(DB); c.row_factory=sqlite3.Row; return c

def init_db():
 c=conn(); c.executescript('''
 CREATE TABLE IF NOT EXISTS projects(id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT NOT NULL,office TEXT NOT NULL,year INTEGER NOT NULL,prefix TEXT NOT NULL,default_recipient TEXT,speed TEXT,security TEXT,address TEXT,phone TEXT,fax TEXT,email TEXT,contact TEXT,created_at TEXT NOT NULL);
 CREATE TABLE IF NOT EXISTS documents(id INTEGER PRIMARY KEY AUTOINCREMENT,project_id INTEGER NOT NULL,doc_date TEXT NOT NULL,serial TEXT NOT NULL,recipient TEXT,speed TEXT,security TEXT,attachment TEXT,subject TEXT,explanation TEXT,originals TEXT,copies TEXT,created_at TEXT NOT NULL,updated_at TEXT NOT NULL,FOREIGN KEY(project_id) REFERENCES projects(id));
 CREATE TABLE IF NOT EXISTS templates(id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT NOT NULL,subject_template TEXT,explanation_template TEXT,attachment_template TEXT,originals_template TEXT,copies_template TEXT,use_material_builder INTEGER DEFAULT 0,quick_options TEXT,created_at TEXT NOT NULL,updated_at TEXT NOT NULL);
 ''')
 if c.execute('SELECT COUNT(*) FROM projects').fetchone()[0]==0:
  c.execute('''INSERT INTO projects(name,office,year,prefix,default_recipient,speed,security,address,phone,fax,email,contact,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)''',(SAMPLE['name'],SAMPLE['office'],SAMPLE['year'],SAMPLE['prefix'],SAMPLE['recipient'],SAMPLE['speed'],SAMPLE['security'],SAMPLE['address'],SAMPLE['phone'],SAMPLE['fax'],SAMPLE['email'],SAMPLE['contact'],datetime.now().isoformat(timespec='seconds')))
 if c.execute('SELECT COUNT(*) FROM templates').fetchone()[0]==0:
  now=datetime.now().isoformat(timespec='seconds')
  c.execute('''INSERT INTO templates(name,subject_template,explanation_template,attachment_template,originals_template,copies_template,use_material_builder,quick_options,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?)''',('材料送審','本所辦理 貴單位「{{工程名稱}}」，檢送本案材料送審資料，詳如說明，敬請 鑒核。','依據契約規定辦理。\n{{材料清單}}','','{{受文者}}','',1,'廠商證明文件,規格樣式,試驗報告,防火證明,型錄,出廠證明,CNS證明',now,now))
 c.commit(); c.close()

def get_project(pid): c=conn(); r=c.execute('SELECT * FROM projects WHERE id=?',(pid,)).fetchone(); c.close(); return r
def get_template(tid): c=conn(); r=c.execute('SELECT * FROM templates WHERE id=?',(tid,)).fetchone(); c.close(); return r
def all_projects(): c=conn(); r=c.execute('SELECT * FROM projects ORDER BY id DESC').fetchall(); c.close(); return r
def all_templates(): c=conn(); r=c.execute('SELECT * FROM templates ORDER BY id DESC').fetchall(); c.close(); return r
def project_docs(pid): c=conn(); r=c.execute('SELECT * FROM documents WHERE project_id=? ORDER BY id DESC',(pid,)).fetchall(); c.close(); return r
def all_docs(q=''):
 c=conn()
 if q:r=c.execute('''SELECT d.*,p.name project_name,p.office,p.prefix FROM documents d JOIN projects p ON p.id=d.project_id WHERE d.subject LIKE ? OR d.recipient LIKE ? OR p.name LIKE ? ORDER BY d.id DESC''',(f'%{q}%',f'%{q}%',f'%{q}%')).fetchall()
 else:r=c.execute('''SELECT d.*,p.name project_name,p.office,p.prefix FROM documents d JOIN projects p ON p.id=d.project_id ORDER BY d.id DESC''').fetchall()
 c.close(); return r

def next_serial(pid,d): c=conn(); n=c.execute('SELECT COUNT(*) FROM documents WHERE project_id=? AND doc_date=?',(pid,d)).fetchone()[0]+1; c.close(); return f'{n:02d}'
def doc_no(p,d,serial):
 dt=datetime.strptime(d,'%Y-%m-%d'); roc=f'{dt.year-1911:03d}{dt.month:02d}{dt.day:02d}'; return f'{p["prefix"]} {roc}-{serial} 號'

def vars_for_project(p):
 today=date.today(); return {'工程名稱':p['name'],'專案名稱':p['name'],'事務所':p['office'],'受文者':p['default_recipient'] or '','聯絡人':p['contact'] or '','民國年':str(today.year-1911),'月份':f'{today.month:02d}','發文日期':f'中華民國 {today.year-1911} 年 {today.month:02d} 月 {today.day:02d} 日'}
def placeholders(t):
 text='\n'.join([t['subject_template'] or '',t['explanation_template'] or '',t['attachment_template'] or '',t['originals_template'] or '',t['copies_template'] or '']); return list(dict.fromkeys(re.findall(r'{{\s*([^{}]+?)\s*}}',text)))
def render_tpl(text,vals):
 def sub(m): return str(vals.get(m.group(1).strip(),''))
 return re.sub(r'{{\s*([^{}]+?)\s*}}',sub,text or '')

@app.route('/')
def index(): q=request.args.get('q',''); return render_template('index.html',projects=all_projects(),docs=all_docs(q),templates=all_templates(),q=q)
@app.route('/projects/<int:pid>')
def project_view(pid):
 p=get_project(pid)
 if not p: abort(404)
 return render_template('project_view.html',project=p,docs=project_docs(pid),templates=all_templates())
@app.route('/projects/new',methods=['GET','POST'])
def project_new():
 if request.method=='POST':
  f=request.form;c=conn();cur=c.execute('''INSERT INTO projects(name,office,year,prefix,default_recipient,speed,security,address,phone,fax,email,contact,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)''',(f['name'],f['office'],int(f['year']),f['prefix'],f.get('default_recipient',''),f.get('speed','普通件'),f.get('security','普通'),f.get('address',''),f.get('phone',''),f.get('fax',''),f.get('email',''),f.get('contact',''),datetime.now().isoformat(timespec='seconds')));c.commit();pid=cur.lastrowid;c.close();flash('專案已建立');return redirect(url_for('project_view',pid=pid))
 return render_template('project_form.html',project=None)
@app.route('/projects/<int:pid>/edit',methods=['GET','POST'])
def project_edit(pid):
 p=get_project(pid)
 if not p: abort(404)
 if request.method=='POST':
  f=request.form;c=conn();c.execute('''UPDATE projects SET name=?,office=?,year=?,prefix=?,default_recipient=?,speed=?,security=?,address=?,phone=?,fax=?,email=?,contact=? WHERE id=?''',(f['name'],f['office'],int(f['year']),f['prefix'],f.get('default_recipient',''),f.get('speed','普通件'),f.get('security','普通'),f.get('address',''),f.get('phone',''),f.get('fax',''),f.get('email',''),f.get('contact',''),pid));c.commit();c.close();flash('專案已更新');return redirect(url_for('project_view',pid=pid))
 return render_template('project_form.html',project=p)

@app.route('/templates/new',methods=['GET','POST'])
@app.route('/templates/<int:tid>/edit',methods=['GET','POST'])
def template_edit(tid=None):
 t=get_template(tid) if tid else None
 if request.method=='POST':
  f=request.form;now=datetime.now().isoformat(timespec='seconds');vals=(f['name'],f.get('subject_template',''),f.get('explanation_template',''),f.get('attachment_template',''),f.get('originals_template',''),f.get('copies_template',''),1 if f.get('use_material_builder') else 0,f.get('quick_options',''),now)
  c=conn()
  if tid:c.execute('''UPDATE templates SET name=?,subject_template=?,explanation_template=?,attachment_template=?,originals_template=?,copies_template=?,use_material_builder=?,quick_options=?,updated_at=? WHERE id=?''',vals+(tid,))
  else:c.execute('''INSERT INTO templates(name,subject_template,explanation_template,attachment_template,originals_template,copies_template,use_material_builder,quick_options,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?)''',vals+(now,))
  c.commit();c.close();flash('範本已儲存');return redirect(url_for('index'))
 return render_template('template_form.html',template=t)
@app.post('/templates/<int:tid>/delete')
def template_delete(tid): c=conn();c.execute('DELETE FROM templates WHERE id=?',(tid,));c.commit();c.close();flash('範本已刪除');return redirect(url_for('index'))

@app.route('/projects/<int:pid>/templates/<int:tid>',methods=['GET','POST'])
def template_generate(pid,tid):
 p=get_project(pid);t=get_template(tid)
 if not p or not t: abort(404)
 known=vars_for_project(p); unknown=[x for x in placeholders(t) if x not in known and x!='材料清單']; options=[x.strip() for x in (t['quick_options'] or '').split(',') if x.strip()]
 if request.method=='POST':
  vals=dict(known)
  for x in unknown: vals[x]=request.form.get('var_'+x,'')
  vals['材料清單']=request.form.get('material_list','')
  return render_template('document_form.html',project=p,doc=None,serial=next_serial(pid,date.today().isoformat()),doc_date=date.today().isoformat(),prefill={'subject':render_tpl(t['subject_template'],vals),'explanation':render_tpl(t['explanation_template'],vals),'attachment':render_tpl(t['attachment_template'],vals),'originals':render_tpl(t['originals_template'],vals) or p['default_recipient'],'copies':render_tpl(t['copies_template'],vals)})
 return render_template('template_generate.html',project=p,template=t,unknown=unknown,options=options)

@app.route('/documents/new')
def document_new():
 pid=request.args.get('project_id',type=int); ps=all_projects(); pid=pid or (ps[0]['id'] if ps else None); p=get_project(pid) if pid else None; d=date.today().isoformat(); return render_template('document_form.html',project=p,doc=None,serial=next_serial(pid,d) if pid else '01',doc_date=d,prefill={})
@app.route('/documents/<int:did>/clone')
def document_clone(did):
 c=conn();d=c.execute('SELECT * FROM documents WHERE id=?',(did,)).fetchone();c.close()
 if not d:abort(404)
 today=date.today().isoformat();return render_template('document_form.html',project=get_project(d['project_id']),doc=None,serial=next_serial(d['project_id'],today),doc_date=today,prefill=dict(d))
@app.route('/documents/<int:did>/edit',methods=['GET','POST'])
def document_edit(did):
 c=conn();d=c.execute('SELECT * FROM documents WHERE id=?',(did,)).fetchone();c.close()
 if not d:abort(404)
 if request.method=='POST':return save_document(did)
 return render_template('document_form.html',project=get_project(d['project_id']),doc=d,serial=d['serial'],doc_date=d['doc_date'],prefill={})
def save_document(did=None):
 f=request.form;pid=int(f['project_id']);dd=f['doc_date'];serial=f.get('serial') or next_serial(pid,dd);now=datetime.now().isoformat(timespec='seconds');vals=(pid,dd,serial,f.get('recipient',''),f.get('speed','普通件'),f.get('security','普通'),f.get('attachment',''),f.get('subject',''),f.get('explanation',''),f.get('originals',''),f.get('copies',''));c=conn()
 if did:c.execute('''UPDATE documents SET project_id=?,doc_date=?,serial=?,recipient=?,speed=?,security=?,attachment=?,subject=?,explanation=?,originals=?,copies=?,updated_at=? WHERE id=?''',vals+(now,did));new_id=did
 else:cur=c.execute('''INSERT INTO documents(project_id,doc_date,serial,recipient,speed,security,attachment,subject,explanation,originals,copies,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)''',vals+(now,now));new_id=cur.lastrowid
 c.commit();c.close();flash('公文已儲存');return redirect(url_for('document_view',did=new_id))
@app.post('/documents/save')
def document_save():return save_document()
@app.route('/documents/<int:did>')
def document_view(did):
 c=conn();d=c.execute('''SELECT d.*,p.name project_name,p.office,p.prefix,p.year,p.address,p.phone,p.fax,p.email,p.contact FROM documents d JOIN projects p ON p.id=d.project_id WHERE d.id=?''',(did,)).fetchone();c.close()
 if not d:abort(404)
 return render_template('document_view.html',doc=d,no=doc_no(d,d['doc_date'],d['serial']))

@app.get('/documents/<int:did>/word')
def word(did):
 c=conn();d=c.execute('''SELECT d.*,p.name project_name,p.office,p.prefix,p.year,p.address,p.phone,p.fax,p.email,p.contact FROM documents d JOIN projects p ON p.id=d.project_id WHERE d.id=?''',(did,)).fetchone();c.close()
 if not d:abort(404)
 doc=Document();sec=doc.sections[0];sec.top_margin=Mm(18);sec.bottom_margin=Mm(18);sec.left_margin=Mm(24);sec.right_margin=Mm(20);sec.header_distance=Mm(5)
 normal=doc.styles['Normal'];normal.font.name='標楷體';normal._element.rPr.rFonts.set(qn('w:eastAsia'),'標楷體');normal.font.size=Pt(14);normal.paragraph_format.space_after=Pt(0);normal.paragraph_format.line_spacing=1.35
 def set_run(run,size=14,bold=False):run.font.name='標楷體';run._element.rPr.rFonts.set(qn('w:eastAsia'),'標楷體');run.font.size=Pt(size);run.bold=bold
 def para(text='',size=14,bold=False,align=None,before=0,after=0):
  p=doc.add_paragraph();p.paragraph_format.space_before=Pt(before);p.paragraph_format.space_after=Pt(after)
  if align is not None:p.alignment=align
  set_run(p.add_run(text),size,bold);return p
 def field(label,value,size=12):
  p=doc.add_paragraph();p.paragraph_format.space_after=Pt(0);p.paragraph_format.line_spacing=1.25;set_run(p.add_run(label),size);set_run(p.add_run(str(value or '')),size);return p
 # 裝訂線：放在頁首中的絕對定位浮動文字方塊，不佔正文空間
 hp=sec.header.paragraphs[0]; pict=parse_xml(r'''<w:pict %s><v:shape xmlns:v="urn:schemas-microsoft-com:vml" style="position:absolute;left:7mm;top:45mm;width:9mm;height:185mm;z-index:-1;mso-position-horizontal-relative:page;mso-position-vertical-relative:page" stroked="f" filled="f"><v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:eastAsia="標楷體"/><w:sz w:val="20"/></w:rPr><w:t>裝</w:t><w:br/><w:br/><w:t>訂</w:t><w:br/><w:br/><w:t>線</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict>''' % nsdecls('w'))
 hp._p.append(pict)
 para(f'{d["office"]}　函',20,True,WD_ALIGN_PARAGRAPH.CENTER,after=8)
 t=doc.add_table(rows=1,cols=2);t.alignment=WD_TABLE_ALIGNMENT.RIGHT;t.autofit=False;t.columns[0].width=Mm(68);t.columns[1].width=Mm(92)
 for cell in t.rows[0].cells:
  tcPr=cell._tc.get_or_add_tcPr();borders=OxmlElement('w:tcBorders');tcPr.append(borders)
 right=t.cell(0,1);right.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP;right.text='';info=[f'聯絡住址：{d["address"] or ""}',f'電話：{d["phone"] or ""}　傳真：{d["fax"] or ""}',f'E-Mail：{d["email"] or ""}',f'連絡人：{d["contact"] or ""}']
 for i,x in enumerate(info):pp=right.paragraphs[0] if i==0 else right.add_paragraph();pp.paragraph_format.space_after=Pt(0);pp.paragraph_format.line_spacing=1.05;set_run(pp.add_run(x),10)
 para('',before=5,after=2);field('受文者：',d['recipient'],12);dt=datetime.strptime(d['doc_date'],'%Y-%m-%d');roc=dt.year-1911;field('發文日期：',f'中華民國 {roc} 年 {dt.month:02d} 月 {dt.day:02d} 日');field('發文字號：',doc_no(d,d['doc_date'],d['serial']));field('速別：',d['speed']);field('密等及解密條件：',d['security']);field('附件：',d['attachment'])
 p=doc.add_paragraph();p.paragraph_format.space_before=Pt(4);p.paragraph_format.space_after=Pt(0);p.paragraph_format.line_spacing=1.35;set_run(p.add_run('主旨：'),14,False);set_run(p.add_run(d['subject'] or ''),14,False)
 p=doc.add_paragraph();p.paragraph_format.space_after=Pt(0);set_run(p.add_run('說明：'),14,False)
 for x in [x for x in (d['explanation'] or '').splitlines() if x.strip()]:pp=doc.add_paragraph();pp.paragraph_format.left_indent=Mm(8);pp.paragraph_format.space_after=Pt(0);pp.paragraph_format.line_spacing=1.35;set_run(pp.add_run(x),14,False)
 para('',before=4,after=2);field('正本：',d['originals'],12);field('副本：',d['copies'],12)
 bio=io.BytesIO();doc.save(bio);bio.seek(0);return send_file(bio,as_attachment=True,download_name=f'{doc_no(d,d["doc_date"],d["serial"]).replace(" ","")}-公文.docx',mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
@app.post('/documents/<int:did>/delete')
def document_delete(did):c=conn();c.execute('DELETE FROM documents WHERE id=?',(did,));c.commit();c.close();flash('公文已刪除');return redirect(url_for('index'))
with app.app_context():init_db()
if __name__=='__main__':app.run(host='0.0.0.0',port=int(os.getenv('PORT',5000)),debug=os.getenv('FLASK_DEBUG')=='1')
