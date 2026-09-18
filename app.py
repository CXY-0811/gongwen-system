import os, sqlite3, io
from datetime import date, datetime
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, abort
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'change-this-secret-key')
DB = os.getenv('DATABASE_PATH', os.path.join(os.path.dirname(__file__), 'data.db'))

SAMPLE = {
 'name':'新民國小跑道整修工程', 'office':'徐英豪建築師事務所', 'year':115,
 'prefix':'一一五所新環字第', 'recipient':'彰化縣田中鎮新民國民小學',
 'speed':'普通件','security':'普通','address':'臺中市西屯區西屯路三段 159-72 號 5 樓',
 'phone':'04-24631919','fax':'04-24631509','email':'yh6726.mail@msa.hinet.net','contact':'黃先生'
}

def conn():
    c=sqlite3.connect(DB); c.row_factory=sqlite3.Row; return c

def init_db():
    c=conn()
    c.executescript('''
    CREATE TABLE IF NOT EXISTS projects(
      id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, office TEXT NOT NULL,
      year INTEGER NOT NULL, prefix TEXT NOT NULL, default_recipient TEXT, speed TEXT,
      security TEXT, address TEXT, phone TEXT, fax TEXT, email TEXT, contact TEXT,
      created_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS documents(
      id INTEGER PRIMARY KEY AUTOINCREMENT, project_id INTEGER NOT NULL,
      doc_date TEXT NOT NULL, serial TEXT NOT NULL, recipient TEXT, speed TEXT,
      security TEXT, attachment TEXT, subject TEXT, explanation TEXT,
      originals TEXT, copies TEXT, created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
      FOREIGN KEY(project_id) REFERENCES projects(id)
    );
    ''')
    if c.execute('SELECT COUNT(*) FROM projects').fetchone()[0]==0:
        c.execute('''INSERT INTO projects(name,office,year,prefix,default_recipient,speed,security,address,phone,fax,email,contact,created_at)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)''', (SAMPLE['name'],SAMPLE['office'],SAMPLE['year'],SAMPLE['prefix'],SAMPLE['recipient'],SAMPLE['speed'],SAMPLE['security'],SAMPLE['address'],SAMPLE['phone'],SAMPLE['fax'],SAMPLE['email'],SAMPLE['contact'],datetime.now().isoformat(timespec='seconds')))
    c.commit(); c.close()

def get_project(pid):
    r=conn().execute('SELECT * FROM projects WHERE id=?',(pid,)).fetchone()
    return r

def zh_year(y):
    digits='〇一二三四五六七八九'
    return ''.join(digits[int(x)] for x in str(y))

def next_serial(pid, d):
    c=conn(); n=c.execute('SELECT COUNT(*) FROM documents WHERE project_id=? AND doc_date=?',(pid,d)).fetchone()[0]+1; c.close(); return f'{n:02d}'

def doc_no(p, d, serial):
    compact=d.replace('-','')
    y=int(d[:4])-1911
    return f'{p["prefix"]} {compact}-{serial} 號'

def all_projects(): return conn().execute('SELECT * FROM projects ORDER BY id DESC').fetchall()
def all_docs(q=''):
    if q:
      return conn().execute('''SELECT d.*,p.name project_name,p.office,p.prefix FROM documents d JOIN projects p ON p.id=d.project_id
      WHERE d.subject LIKE ? OR d.recipient LIKE ? OR p.name LIKE ? ORDER BY d.id DESC''',(f'%{q}%',f'%{q}%',f'%{q}%')).fetchall()
    return conn().execute('''SELECT d.*,p.name project_name,p.office,p.prefix FROM documents d JOIN projects p ON p.id=d.project_id ORDER BY d.id DESC''').fetchall()

@app.route('/')
def index():
    q=request.args.get('q',''); return render_template('index.html', projects=all_projects(), docs=all_docs(q), q=q)

@app.route('/projects/new', methods=['GET','POST'])
def project_new():
    if request.method=='POST':
      f=request.form
      c=conn(); c.execute('''INSERT INTO projects(name,office,year,prefix,default_recipient,speed,security,address,phone,fax,email,contact,created_at)
      VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)''',(f['name'],f['office'],int(f['year']),f['prefix'],f.get('default_recipient',''),f.get('speed','普通件'),f.get('security','普通'),f.get('address',''),f.get('phone',''),f.get('fax',''),f.get('email',''),f.get('contact',''),datetime.now().isoformat(timespec='seconds'))); c.commit(); pid=c.lastrowid; c.close(); flash('專案已建立'); return redirect(url_for('document_new',project_id=pid))
    return render_template('project_form.html', project=None)

@app.route('/projects/<int:pid>/edit', methods=['GET','POST'])
def project_edit(pid):
    p=get_project(pid)
    if not p: abort(404)
    if request.method=='POST':
      f=request.form; c=conn(); c.execute('''UPDATE projects SET name=?,office=?,year=?,prefix=?,default_recipient=?,speed=?,security=?,address=?,phone=?,fax=?,email=?,contact=? WHERE id=?''',(f['name'],f['office'],int(f['year']),f['prefix'],f.get('default_recipient',''),f.get('speed','普通件'),f.get('security','普通'),f.get('address',''),f.get('phone',''),f.get('fax',''),f.get('email',''),f.get('contact',''),pid)); c.commit(); c.close(); flash('專案已更新'); return redirect(url_for('index'))
    return render_template('project_form.html', project=p)

@app.route('/documents/new')
def document_new():
    pid=request.args.get('project_id', type=int)
    if not pid and all_projects(): pid=all_projects()[0]['id']
    p=get_project(pid) if pid else None
    d=date.today().isoformat(); serial=next_serial(pid,d) if pid else '01'
    return render_template('document_form.html', project=p, doc=None, serial=serial, doc_date=d)

@app.route('/documents/<int:did>/edit', methods=['GET','POST'])
def document_edit(did):
    c=conn(); d=c.execute('SELECT d.*,p.name project_name,p.office,p.prefix,p.year,p.address,p.phone,p.fax,p.email,p.contact FROM documents d JOIN projects p ON p.id=d.project_id WHERE d.id=?',(did,)).fetchone(); c.close()
    if not d: abort(404)
    if request.method=='POST': return save_document(did)
    return render_template('document_form.html', project=get_project(d['project_id']), doc=d, serial=d['serial'], doc_date=d['doc_date'])

def save_document(did=None):
    f=request.form; pid=int(f['project_id']); d=f['doc_date']; serial=f.get('serial') or next_serial(pid,d); now=datetime.now().isoformat(timespec='seconds')
    vals=(pid,d,serial,f.get('recipient',''),f.get('speed','普通件'),f.get('security','普通'),f.get('attachment',''),f.get('subject',''),f.get('explanation',''),f.get('originals',''),f.get('copies',''))
    c=conn()
    if did: c.execute('''UPDATE documents SET project_id=?,doc_date=?,serial=?,recipient=?,speed=?,security=?,attachment=?,subject=?,explanation=?,originals=?,copies=?,updated_at=? WHERE id=?''',vals+(now,did));
    else: c.execute('''INSERT INTO documents(project_id,doc_date,serial,recipient,speed,security,attachment,subject,explanation,originals,copies,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)''',vals+(now,now))
    c.commit(); new_id=did or c.lastrowid; c.close(); flash('公文已儲存'); return redirect(url_for('document_view',did=new_id))

@app.post('/documents/save')
def document_save(): return save_document()

@app.route('/documents/<int:did>')
def document_view(did):
    c=conn(); d=c.execute('''SELECT d.*,p.name project_name,p.office,p.prefix,p.year,p.address,p.phone,p.fax,p.email,p.contact FROM documents d JOIN projects p ON p.id=d.project_id WHERE d.id=?''',(did,)).fetchone(); c.close()
    if not d: abort(404)
    return render_template('document_view.html', doc=d, no=doc_no(d,d['doc_date'],d['serial']))

@app.get('/documents/<int:did>/word')
def word(did):
    c=conn(); d=c.execute('''SELECT d.*,p.name project_name,p.office,p.prefix,p.year,p.address,p.phone,p.fax,p.email,p.contact FROM documents d JOIN projects p ON p.id=d.project_id WHERE d.id=?''',(did,)).fetchone(); c.close()
    if not d: abort(404)
    doc=Document(); sec=doc.sections[0]; sec.top_margin=Mm(20); sec.bottom_margin=Mm(20); sec.left_margin=Mm(25); sec.right_margin=Mm(25)
    styles=doc.styles; styles['Normal'].font.name='標楷體'; styles['Normal']._element.rPr.rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia','標楷體'); styles['Normal'].font.size=Pt(14)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=p.add_run(f'{d["office"]} 函'); r.bold=True; r.font.size=Pt(20); r.font.name='標楷體'
    def line(label,val):
      p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(0); p.add_run(label).bold=True; p.add_run(str(val or ''))
    line('受文者：',d['recipient']); line('發文日期：',f'中華民國 {d["year"]} 年 {d["doc_date"][5:7]} 月 {d["doc_date"][8:10]} 日'); line('發文字號：',doc_no(d,d['doc_date'],d['serial'])); line('速別：',d['speed']); line('密等及解密條件：',d['security']); line('附件：',d['attachment'])
    p=doc.add_paragraph(); p.add_run('主旨：').bold=True; p.add_run(d['subject'] or '')
    p=doc.add_paragraph(); p.add_run('說明：').bold=True
    for x in (d['explanation'] or '').splitlines():
      if x.strip(): doc.add_paragraph(x.strip())
    line('正本：',d['originals']); line('副本：',d['copies']); line('聯絡住址：',d['address']); line('電話：',d['phone']); line('傳真：',d['fax']); line('E-Mail：',d['email']); line('連絡人：',d['contact'])
    bio=io.BytesIO(); doc.save(bio); bio.seek(0); return send_file(bio,as_attachment=True,download_name=f'{d["doc_date"].replace("-","")}-{d["serial"]}-公文.docx',mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

@app.post('/documents/<int:did>/delete')
def document_delete(did):
    c=conn(); c.execute('DELETE FROM documents WHERE id=?',(did,)); c.commit(); c.close(); flash('公文已刪除'); return redirect(url_for('index'))

with app.app_context(): init_db()

if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.getenv('PORT',5000)),debug=os.getenv('FLASK_DEBUG')=='1')
