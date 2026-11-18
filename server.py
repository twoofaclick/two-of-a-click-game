#!/usr/bin/env python3
"""Wedding Expo game: standard-library HTTP server, SQLite source of truth, Excel exports."""
import os, json, sqlite3, secrets, random, re, hmac, io, sys, threading
from datetime import datetime, date, timedelta
from pathlib import Path
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from http.cookies import SimpleCookie

ROOT=Path(__file__).resolve().parent
DB=Path(os.getenv('GAME_DB', str(ROOT/'wedding_game.sqlite3')))
PORT=int(os.getenv('PORT','8080'))
ADMIN_PASSWORD=os.getenv('ADMIN_PASSWORD','')
EXPO_DATE=os.getenv('EXPO_DATE', date.today().isoformat())
PRIZES=[('Free couples photoshoot',1),('20% off packages',3),('15% off packages',6),('10% off packages',20),('Free event trailer',2),('Free DJI behind-the-scenes',3),('Free 30-minute extra coverage',5)]
lock=threading.Lock()

def connect():
    c=sqlite3.connect(DB, timeout=15)
    c.row_factory=sqlite3.Row
    c.execute('PRAGMA journal_mode=WAL')
    c.execute('PRAGMA busy_timeout=15000')
    return c

def init():
    with connect() as c:
        c.execute('''CREATE TABLE IF NOT EXISTS entries (id TEXT PRIMARY KEY, created_at TEXT NOT NULL, full_name TEXT NOT NULL, phone TEXT NOT NULL, phone_key TEXT NOT NULL, email TEXT NOT NULL, email_key TEXT NOT NULL, social TEXT NOT NULL, event_date TEXT NOT NULL, prize TEXT NOT NULL, claim_status TEXT NOT NULL DEFAULT 'Unclaimed', booking_status TEXT NOT NULL DEFAULT 'Not booked', marketing_opt_in INTEGER NOT NULL DEFAULT 0, follow_attested INTEGER NOT NULL DEFAULT 1, follow_verified INTEGER NOT NULL DEFAULT 0, notes TEXT NOT NULL DEFAULT '')''')
        c.execute('CREATE UNIQUE INDEX IF NOT EXISTS unique_email_event ON entries(email_key,event_date)')
        c.execute('CREATE UNIQUE INDEX IF NOT EXISTS unique_phone_event ON entries(phone_key,event_date)')
        c.execute('CREATE TABLE IF NOT EXISTS inventory (prize TEXT PRIMARY KEY, initial_count INTEGER NOT NULL, remaining INTEGER NOT NULL)')
        for name,n in PRIZES: c.execute('INSERT OR IGNORE INTO inventory VALUES (?,?,?)',(name,n,n))

def as_json(handler,code,data):
    b=json.dumps(data).encode()
    handler.send_response(code);handler.send_header('Content-Type','application/json');handler.send_header('Cache-Control','no-store');handler.send_header('Content-Length',str(len(b)));handler.end_headers();handler.wfile.write(b)

def valid_admin(handler):
    value=handler.headers.get('Authorization','')
    return bool(ADMIN_PASSWORD) and value.startswith('Basic ') and hmac.compare_digest(__import__('base64').b64decode(value[6:]).decode(errors='replace').split(':',1)[-1],ADMIN_PASSWORD)

class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        p=urlparse(self.path).path
        if p in ('/','/index.html'):
            b=(ROOT/'index.html').read_bytes();self.send_response(200);self.send_header('Content-Type','text/html; charset=utf-8');self.send_header('Content-Length',str(len(b)));self.end_headers();self.wfile.write(b);return
        if p=='/api/config':
            with connect() as c:
                rows=c.execute('SELECT prize,remaining FROM inventory').fetchall()
            return as_json(self,200,{'prizes':[dict(r) for r in rows],'expo_date':EXPO_DATE})
        if p.startswith('/admin'):
            if not valid_admin(self):
                self.send_response(401);self.send_header('WWW-Authenticate','Basic realm="Wedding Expo Admin"');self.end_headers();return
            if p=='/admin/entries':
                with connect() as c: entries=[dict(r) for r in c.execute('SELECT * FROM entries ORDER BY created_at DESC')]
                return as_json(self,200,{'entries':entries})
            if p=='/admin/export.xlsx':
                try:
                    from artifact_tool import Workbook, SpreadsheetFile
                    with connect() as c: entries=[dict(r) for r in c.execute('SELECT * FROM entries ORDER BY created_at DESC')]
                    wb=Workbook.create();sh=wb.worksheets.add('Client Register')
                    headers=['Entry ID','Registered UTC','Full name','Phone','Email','Social username','2027 event date','Prize won','Claim status','Booking status','Marketing opt-in','Follow attested','Follow staff verified','Claim by','Booking by','Contact at least 1 month before','Notes']
                    sh.get_range('A1:Q1').values=[headers]
                    data=[]
                    try: expo=date.fromisoformat(EXPO_DATE)
                    except: expo=date.today()
                    for r in entries:
                        d=date.fromisoformat(r['event_date'])
                        month=d.month-1 or 12;year=d.year-(1 if d.month==1 else 0)
                        import calendar
                        contact=date(year,month,min(d.day,calendar.monthrange(year,month)[1]))
                        data.append([r['id'],r['created_at'],r['full_name'],r['phone'],r['email'],r['social'],r['event_date'],r['prize'],r['claim_status'],r['booking_status'],'Yes' if r['marketing_opt_in'] else 'No','Yes' if r['follow_attested'] else 'No','Yes' if r['follow_verified'] else 'No',str(expo+timedelta(days=14)),str(expo+timedelta(days=30)),str(contact),r['notes']])
                    if data: sh.get_range_by_indexes(1,0,len(data),len(headers)).values=data
                    sh.freeze_panes.freeze_rows(1)
                    sh.get_range('A1:Q1').format.fill='#2F2340';sh.get_range('A1:Q1').format.font.color='#FFFFFF'
                    sh.get_range('A:Q').format.column_width=19
                    sh.get_range('C:C').format.column_width=26;sh.get_range('H:H').format.column_width=30
                    inv=wb.worksheets.add('Prize Inventory')
                    inv.get_range('A1:C1').values=[['Prize','Initial allocation','Remaining']]
                    with connect() as c: invrows=[list(r) for r in c.execute('SELECT prize,initial_count,remaining FROM inventory')]
                    if invrows: inv.get_range_by_indexes(1,0,len(invrows),3).values=invrows
                    inv.get_range('A:A').format.column_width=35
                    out=ROOT/'latest_client_register.xlsx';SpreadsheetFile.export_xlsx(wb).save(str(out));b=out.read_bytes()
                    self.send_response(200);self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');self.send_header('Content-Disposition','attachment; filename="Wedding_Expo_Client_Register.xlsx"');self.send_header('Content-Length',str(len(b)));self.end_headers();self.wfile.write(b);return
                except Exception as e: return as_json(self,500,{'error':'Export failed: '+str(e)})
        return as_json(self,404,{'error':'Not found'})
    def do_POST(self):
        p=urlparse(self.path).path
        if p=='/admin/update':
            if not valid_admin(self):return as_json(self,401,{'error':'Admin login required'})
            try:
                n=int(self.headers.get('Content-Length','0'));payload=json.loads(self.rfile.read(n))
                if payload.get('claim_status') not in ('Unclaimed','Claimed','Forfeited') or payload.get('booking_status') not in ('Not booked','Contacted','Booked','Cancelled'):raise ValueError('Invalid status')
                with connect() as c:c.execute('UPDATE entries SET claim_status=?,booking_status=?,follow_verified=?,notes=? WHERE id=?',(payload['claim_status'],payload['booking_status'],int(bool(payload.get('follow_verified'))),str(payload.get('notes',''))[:500],payload['id']))
                return as_json(self,200,{'ok':True})
            except Exception as e:return as_json(self,400,{'error':str(e)})
        if p!='/api/play':return as_json(self,404,{'error':'Not found'})
        try:
            n=int(self.headers.get('Content-Length','0'))
            if n>10000 or n<=0:raise ValueError('Invalid submission')
            v=json.loads(self.rfile.read(n));name=str(v.get('name','')).strip();phone=str(v.get('phone','')).strip();email=str(v.get('email','')).strip().lower();social=str(v.get('social','')).strip();event=str(v.get('event_date',''))
            if not 2<=len(name)<=100 or not re.match(r'^[^@\s]+@[^@\s]+\.[^@\s]+$',email) or not 7<=len(re.sub(r'\D','',phone))<=16 or not 2<=len(social)<=80:raise ValueError('Please enter valid contact details')
            if not v.get('followed') or not v.get('terms'):raise ValueError('Confirm the follow and giveaway terms')
            d=date.fromisoformat(event)
            if d.year!=2027:raise ValueError('Event must take place in 2027')
            phone_key=re.sub(r'\D','',phone);eid=secrets.token_hex(6).upper()
            with lock:
                with connect() as c:
                    c.execute('BEGIN IMMEDIATE')
                    old=c.execute('SELECT id FROM entries WHERE (email_key=? OR phone_key=?) AND event_date=?',(email,phone_key,event)).fetchone()
                    if old:return as_json(self,409,{'error':'This client/event has already played. Please ask the team for help.'})
                    available=c.execute('SELECT prize,remaining FROM inventory WHERE remaining>0').fetchall()
                    if not available:return as_json(self,409,{'error':'All prizes have been allocated.'})
                    prize=random.choices([r['prize'] for r in available],weights=[r['remaining'] for r in available],k=1)[0]
                    c.execute('UPDATE inventory SET remaining=remaining-1 WHERE prize=? AND remaining>0',(prize,))
                    c.execute('INSERT INTO entries (id,created_at,full_name,phone,phone_key,email,email_key,social,event_date,prize,marketing_opt_in) VALUES (?,?,?,?,?,?,?,?,?,?,?)',(eid,datetime.now(__import__('datetime').timezone.utc).isoformat(),name,phone,phone_key,email,email,social,event,prize,int(bool(v.get('marketing')))))
            return as_json(self,200,{'entry_id':eid,'prize':prize,'message':'Saved successfully'})
        except ValueError as e:return as_json(self,400,{'error':str(e)})
        except sqlite3.IntegrityError:return as_json(self,409,{'error':'Duplicate entry. Please ask the team for help.'})
        except Exception as e:return as_json(self,500,{'error':'Unable to save. Please try again or ask the team.'})

if __name__=='__main__':
    if not ADMIN_PASSWORD:sys.exit('Set ADMIN_PASSWORD before starting the game.')
    init();print(f'Game running at http://localhost:{PORT} | admin export /admin/export.xlsx')
    ThreadingHTTPServer(('0.0.0.0',PORT),Handler).serve_forever()
