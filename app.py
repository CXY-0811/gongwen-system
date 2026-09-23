import os, sqlite3, io, re
from datetime import date, datetime
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, abort
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'change-this-secret-key')
DB = os.getenv('DATABASE_PATH', os.path.join(os.path.dirname(__file__), 'data.db'))

SAMPLE = {'name':'新民國小跑道整修工程','office':'徐英豪建築師事務所','year':115,'prefix':'一一五所新環字第','recipient':'彰化縣田中鎮新民國民小學','speed':'普通件','security':'普通','address':'臺中市西屯區西屯路三段 159-72 號 5 樓','phone':'04-24631919','fax':'04-24631509','email':'yh6726.mail@msa.hinet.net','contact':'黃先生'}
DEFAULT_TEMPLATE = {
 'name':'施工計畫／品質計畫／職安計畫審查',
 'subject':'本所辦理 貴校「{{工程名稱}}」委託設計監造採購案，檢送本案施工計畫、品質計畫及職業安全衛生管理計畫乙式三份，審查情形詳如說明，敬請 鑒核。',
 'explanation':'1.依據廠商 {{來文日期}} 發文字號：{{來文字號}}及契約規定辦理。\n2.廠商檢送上述文件，經本所審查同意，隨文檢附旨揭文件，詳后附件。',
 'originals':'{{受文者}}','copies':'{{廠商名稱}}、本所設計部'
}

def conn():
    c=sqlite3.connect(DB); c.row_factory=sqlite3.Row; return c

def init_db():
    c=conn(); c.executescript('''
    CREATE TABLE IF NOT EXISTS projects(id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT NOT NULL,office TEXT NOT NULL,year INTEGER NOT NULL,prefix TEXT NOT NULL,default_recipient TEXT,speed TEXT,security TEXT,address TEXT,phone TEXT,fax TEXT,email TEXT,contact TEXT,created_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS documents(id INTEGER PRIMARY KEY AUTOINCREMENT,project_id INTEGER NOT NULL,doc_date TEXT NOT NULL,serial TEXT NOT NULL,recipient TEXT,speed TEXT,security TEXT,attachment TEXT,subject TEXT,explanation TEXT,originals TEXT,copies TEXT,created_at TEXT NOT NULL,updated_at TEXT NOT NULL,FOREIGN KEY(project_id) REFERENCES projects(id));
    CREATE TABLE IF NOT EXISTS templates(id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT NOT NULL,subject TEXT,explanation TEXT,originals TEXT,copies TEXT,created_at TEXT NOT NULL,updated_at TEXT NOT NULL);
    ''')
    if c.execute('SELECT COUNT(*) FROM projects').fetchone()[0]==0:
        c.execute('INSERT INTO projects(name,office,year,prefix,default_recipient,speed,security,address,phone,fax,email,contact,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)',(SAMPLE['name'],SAMPLE['office'],SAMPLE['year'],SAMPLE['prefix'],SAMPLE['recipient'],SAMPLE['speed'],SAMPLE['security'],SAMPLE['address'],SAMPLE['phone'],SAMPLE['fax'],SAMPLE['email'],SAMPLE['contact'],datetime.now().isoformat(timespec='seconds')))
    if c.execute('SELECT COUNT(*) FROM templates').fetchone()[0]==0:
        now=datetime.now().isoformat(timespec='seconds'); t=DEFAULT_TEMPLATE
        c.execute('INSERT INTO templates(name,subject,explanation,originals,copies,created_at,updated_at) VALUES(?,?,?,?,?,?,?)',(t['name'],t['subject'],t['explanation'],t['originals'],t['copies'],now,now))
    c.commit(); c.close()

def get_project(pid): return conn().execute('SELECT * FROM projects WHERE id=?',(pid,)).fetchone()
def get_template(tid): return conn().execute('SELECT * FROM templates WHERE id=?',(tid,)).fetchone()
def all_projects(): return conn().execute('SELECT * FROM projects ORDER BY id DESC').fetchall()
def all_templates(): return conn().execute('SELECT * FROM templates ORDER BY id DESC').fetchall()
def next_serial(pid,d):
    c=conn(); n=c.execute('SELECT COUNT(*) FROM documents WHERE project_id=? AND doc_date=?',(pid,d)).fetchone()[0]+1; c.close(); return f'{n:02d}'
def doc_no(p,d,serial):
    dt=datetime.strptime(d,'%Y-%m-%d'); return f'{p["prefix"]} {dt.year-1911:03d}{dt.month:02d}{dt.day:02d}-{serial} 號'
def all_docs(q=''):
    c=conn()
    if q: r=c.execute('SELECT d.*,p.name project_name,p.office,p.prefix FROM documents d JOIN projects p ON p.id=d.project_id WHERE d.subject LIKE ? OR d.recipient LIKE ? OR p.name LIKE ? ORDER BY d.id DESC',(f'%{q}%',f'%{q}%',f'%{q}%')).fetchall()
    else: r=c.execute('SELECT d.*,p.name project_name,p.office,p.prefix FROM documents d JOIN projects p ON p.id=d.project_id ORDER BY d.id DESC').fetchall()
    c.close(); return r

def render_vars(text, values):
    text=text or ''
    def repl(m): return str(values.get(m.group(1).strip(), m.group(0)))
    return re.sub(r'{{\s*([^{}]+?)\s*}}', repl, text)

def builtins(p, doc_date):
    dt=datetime.strptime(doc_date,'%Y-%m-%d'); roc=dt.year-1911
    return {'工程名稱':p['name'],'專案名稱':p['name'],'事務所':p['office'],'受文者':p['default_recipient'] or '', '發文日期':f'{roc}年{dt.month:02d}月{dt.day:02d}日','民國年':str(roc),'月份':str(dt.month),'聯絡人':p['contact'] or ''}

@app.route('/')
def index():
    q=request.args.get('q',''); return render_template('index.html',projects=all_projects(),templates=all_templates(),docs=all_docs(q),q=q)

@app.route('/projects/new',methods=['GET','POST'])
def project_new():
    if request.method=='POST':
        f=request.form; c=conn(); cur=c.execute('INSERT INTO projects(name,office,year,prefix,default_recipient,speed,security,address,phone,fax,email,contact,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)',(f['name'],f['office'],int(f['year']),f['prefix'],f.get('default_recipient',''),f.get('speed','普通件'),f.get('security','普通'),f.get('address',''),f.get('phone',''),f.get('fax',''),f.get('email',''),f.get('contact',''),datetime.now().isoformat(timespec='seconds'))); c.commit(); pid=cur.lastrowid; c.close(); flash('專案已建立'); return redirect(url_for('document_new',project_id=pid))
    return render_template('project_form.html',project=None)

@app.route('/projects/<int:pid>/edit',methods=['GET','POST'])
def project_edit(pid):
    p=get_project(pid)
    if not p: abort(404)
    if request.method=='POST':
        f=request.form; c=conn(); c.execute('UPDATE projects SET name=?,office=?,year=?,prefix=?,default_recipient=?,speed=?,security=?,address=?,phone=?,fax=?,email=?,contact=? WHERE id=?',(f['name'],f['office'],int(f['year']),f['prefix'],f.get('default_recipient',''),f.get('speed','普通件'),f.get('security','普通'),f.get('address',''),f.get('phone',''),f.get('fax',''),f.get('email',''),f.get('contact',''),pid)); c.commit(); c.close(); flash('專案已更新'); return redirect(url_for('index'))
    return render_template('project_form.html',project=p)

@app.route('/templates/new',methods=['GET','POST'])
@app.route('/templates/<int:tid>/edit',methods=['GET','POST'])
def template_edit(tid=None):
    t=get_template(tid) if tid else None
    if tid and not t: abort(404)
    if request.method=='POST':
        f=request.form; now=datetime.now().isoformat(timespec='seconds'); c=conn()
        vals=(f['name'],f.get('subject',''),f.get('explanation',''),f.get('originals',''),f.get('copies',''))
        if tid: c.execute('UPDATE templates SET name=?,subject=?,explanation=?,originals=?,copies=?,updated_at=? WHERE id=?',vals+(now,tid))
        else: c.execute('INSERT INTO templates(name,subject,explanation,originals,copies,created_at,updated_at) VALUES(?,?,?,?,?,?,?)',vals+(now,now))
        c.commit(); c.close(); flash('範本已儲存'); return redirect(url_for('index'))
    return render_template('template_form.html',template=t)

@app.post('/templates/<int:tid>/delete')
def template_delete(tid):
    c=conn(); c.execute('DELETE FROM templates WHERE id=?',(tid,)); c.commit(); c.close(); flash('範本已刪除'); return redirect(url_for('index'))

@app.route('/generate/<int:tid>',methods=['GET','POST'])
def generate(tid):
    t=get_template(tid)
    if not t: abort(404)
    projects=all_projects(); pid=request.values.get('project_id',type=int) or (projects[0]['id'] if projects else None)
    p=get_project(pid) if pid else None
    d=request.values.get('doc_date') or date.today().isoformat()
    fields=[]
    joined='\n'.join([t['subject'] or '',t['explanation'] or '',t['originals'] or '',t['copies'] or ''])
    reserved=set(['工程名稱','專案名稱','事務所','受文者','發文日期','民國年','月份','聯絡人'])
    for x in re.findall(r'{{\s*([^{}]+?)\s*}}',joined):
        x=x.strip()
        if x not in reserved and x not in fields: fields.append(x)
    if request.method=='POST' and p:
        vals=builtins(p,d)
        for field in fields: vals[field]=request.form.get('var_'+field,'')
        serial=next_serial(pid,d); now=datetime.now().isoformat(timespec='seconds')
        recipient=request.form.get('recipient') or p['default_recipient'] or ''
        vals['受文者']=recipient
        subject=render_vars(t['subject'],vals); explanation=render_vars(t['explanation'],vals); originals=render_vars(t['originals'],vals); copies=render_vars(t['copies'],vals)
        c=conn(); cur=c.execute('INSERT INTO documents(project_id,doc_date,serial,recipient,speed,security,attachment,subject,explanation,originals,copies,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)',(pid,d,serial,recipient,p['speed'],p['security'],request.form.get('attachment',''),subject,explanation,originals,copies,now,now)); c.commit(); did=cur.lastrowid; c.close(); flash('已由範本產生公文'); return redirect(url_for('document_view',did=did))
    return render_template('generate.html',template=t,projects=projects,project=p,fields=fields,doc_date=d)

@app.route('/documents/new')
def document_new():
    pid=request.args.get('project_id',type=int); projects=all_projects(); pid=pid or (projects[0]['id'] if projects else None); p=get_project(pid) if pid else None; d=date.today().isoformat(); serial=next_serial(pid,d) if pid else '01'; return render_template('document_form.html',project=p,doc=None,serial=serial,doc_date=d)

@app.route('/documents/<int:did>/edit',methods=['GET','POST'])
def document_edit(did):
    c=conn(); d=c.execute('SELECT d.*,p.name project_name,p.office,p.prefix,p.year,p.address,p.phone,p.fax,p.email,p.contact FROM documents d JOIN projects p ON p.id=d.project_id WHERE d.id=?',(did,)).fetchone(); c.close()
    if not d: abort(404)
    if request.method=='POST': return save_document(did)
    return render_template('document_form.html',project=get_project(d['project_id']),doc=d,serial=d['serial'],doc_date=d['doc_date'])

def save_document(did=None):
    f=request.form; pid=int(f['project_id']); d=f['doc_date']; serial=f.get('serial') or next_serial(pid,d); now=datetime.now().isoformat(timespec='seconds'); vals=(pid,d,serial,f.get('recipient',''),f.get('speed','普通件'),f.get('security','普通'),f.get('attachment',''),f.get('subject',''),f.get('explanation',''),f.get('originals',''),f.get('copies','')); c=conn()
    if did: c.execute('UPDATE documents SET project_id=?,doc_date=?,serial=?,recipient=?,speed=?,security=?,attachment=?,subject=?,explanation=?,originals=?,copies=?,updated_at=? WHERE id=?',vals+(now,did)); new_id=did
    else: cur=c.execute('INSERT INTO documents(project_id,doc_date,serial,recipient,speed,security,attachment,subject,explanation,originals,copies,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)',vals+(now,now)); new_id=cur.lastrowid
    c.commit(); c.close(); flash('公文已儲存'); return redirect(url_for('document_view',did=new_id))

@app.post('/documents/save')
def document_save(): return save_document()
@app.route('/documents/<int:did>')
def document_view(did):
    c=conn(); d=c.execute('SELECT d.*,p.name project_name,p.office,p.prefix,p.year,p.address,p.phone,p.fax,p.email,p.contact FROM documents d JOIN projects p ON p.id=d.project_id WHERE d.id=?',(did,)).fetchone(); c.close()
    if not d: abort(404)
    return render_template('document_view.html',doc=d,no=doc_no(d,d['doc_date'],d['serial']))

@app.get('/documents/<int:did>/word')
def word(did):
    c=conn(); d=c.execute('SELECT d.*,p.name project_name,p.office,p.prefix,p.year,p.address,p.phone,p.fax,p.email,p.contact FROM documents d JOIN projects p ON p.id=d.project_id WHERE d.id=?',(did,)).fetchone(); c.close()
    if not d: abort(404)
    doc=Document(); sec=doc.sections[0]; sec.top_margin=Mm(18); sec.bottom_margin=Mm(18); sec.left_margin=Mm(25); sec.right_margin=Mm(20)
    normal=doc.styles['Normal']; normal.font.name='標楷體'; normal._element.rPr.rFonts.set(qn('w:eastAsia'),'標楷體'); normal.font.size=Pt(14); normal.paragraph_format.space_after=Pt(0); normal.paragraph_format.line_spacing=1.35
    # Word 裝訂線：頁首中放置固定位置文字框，避免正文長短影響
    header=sec.header; p=header.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.LEFT
    run=p.add_run('裝\n\n訂\n\n線'); run.font.name='標楷體'; run._element.rPr.rFonts.set(qn('w:eastAsia'),'標楷體'); run.font.size=Pt(10)
    p.paragraph_format.left_indent=Mm(-18); p.paragraph_format.space_before=Mm(48)
    def set_run(run,size=14,bold=False): run.font.name='標楷體'; run._element.rPr.rFonts.set(qn('w:eastAsia'),'標楷體'); run.font.size=Pt(size); run.bold=bold
    def para(text='',size=14,bold=False,align=None,before=0,after=0):
        p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(before); p.paragraph_format.space_after=Pt(after); p.alignment=align if align is not None else p.alignment; set_run(p.add_run(text),size,bold); return p
    def field(label,value,size=14):
        p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=1.25; set_run(p.add_run(label),size); set_run(p.add_run(str(value or '')),size); return p
    para(f'{d["office"]}　函',20,True,WD_ALIGN_PARAGRAPH.CENTER,after=6)
    t=doc.add_table(rows=1,cols=2); t.alignment=WD_TABLE_ALIGNMENT.RIGHT; t.autofit=False; right=t.cell(0,1); right.text=''; right.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP
    for cell in t.rows[0].cells:
        tcPr=cell._tc.get_or_add_tcPr(); borders=OxmlElement('w:tcBorders'); tcPr.append(borders)
    for i,x in enumerate([f'聯絡住址：{d["address"] or ""}',f'電話：{d["phone"] or ""}　傳真：{d["fax"] or ""}',f'E-Mail：{d["email"] or ""}',f'連絡人：{d["contact"] or ""}']):
        pp=right.paragraphs[0] if i==0 else right.add_paragraph(); pp.paragraph_format.space_after=Pt(0); set_run(pp.add_run(x),10)
    field('受文者：',d['recipient'],16); dt=datetime.strptime(d['doc_date'],'%Y-%m-%d'); field('發文日期：',f'中華民國 {dt.year-1911} 年 {dt.month:02d} 月 {dt.day:02d} 日'); field('發文字號：',doc_no(d,d['doc_date'],d['serial'])); field('速別：',d['speed']); field('密等及解密條件：',d['security']); field('附件：',d['attachment'])
    p=doc.add_paragraph(); set_run(p.add_run('主旨：'),14,True); set_run(p.add_run(d['subject'] or ''),14)
    p=doc.add_paragraph(); set_run(p.add_run('說明：'),14,True)
    for x in [x for x in (d['explanation'] or '').splitlines() if x.strip()]:
        pp=doc.add_paragraph(); pp.paragraph_format.left_indent=Mm(8); pp.paragraph_format.space_after=Pt(0); set_run(pp.add_run(x),14)
    field('正本：',d['originals']); field('副本：',d['copies'])
    bio=io.BytesIO(); doc.save(bio); bio.seek(0); return send_file(bio,as_attachment=True,download_name=f'{doc_no(d,d["doc_date"],d["serial"]).replace(" ","")}-公文.docx',mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

@app.post('/documents/<int:did>/delete')
def document_delete(did):
    c=conn(); c.execute('DELETE FROM documents WHERE id=?',(did,)); c.commit(); c.close(); flash('公文已刪除'); return redirect(url_for('index'))

with app.app_context(): init_db()
if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.getenv('PORT',5000)),debug=os.getenv('FLASK_DEBUG')=='1')
