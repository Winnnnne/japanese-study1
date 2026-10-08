import os, json, zipfile, html, re
from lxml import etree
from docx import Document

BASE='/mnt/data'
OUT='/mnt/data/japanese-study-site'
NS={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

# ---------- DOCX helpers ----------
def cell_text_xml(tc):
    return ''.join(tc.xpath('.//w:t/text()', namespaces=NS)).strip()

def read_tables_xml(path):
    with zipfile.ZipFile(path) as z:
        root=etree.fromstring(z.read('word/document.xml'))
    tables=[]
    for tbl in root.xpath('.//w:body//w:tbl', namespaces=NS):
        rows=[]
        for tr in tbl.xpath('./w:tr', namespaces=NS):
            cells=[]
            for tc in tr.xpath('./w:tc', namespaces=NS):
                cells.append(cell_text_xml(tc))
            rows.append(cells)
        if rows:
            tables.append(rows)
    return tables

def read_paragraphs_xml(path):
    with zipfile.ZipFile(path) as z:
        root=etree.fromstring(z.read('word/document.xml'))
    out=[]
    for p in root.xpath('.//w:body//w:p', namespaces=NS):
        t=''.join(p.xpath('.//w:t/text()', namespaces=NS)).strip()
        if t:
            out.append(t)
    return out

# ---------- Vocabulary ----------
vocab_path=os.path.join(BASE,'Tổng hợp từ vựng 4-7(Romaji).docx')
vocab=[]
for ti, rows in enumerate(read_tables_xml(vocab_path)):
    if not rows: continue
    first=' '.join(rows[0])
    m=re.search(r'Unit\s*(\d+)(?:\.(\d+))?', first, re.I)
    if not m: continue
    unit=int(m.group(1)); sub=int(m.group(2) or 0)
    subkey=f'{unit}.{sub}' if sub else str(unit)
    for row in rows[1:]:
        row=[x.strip() for x in row]
        if not any(row): continue
        # Most tables: [kanji, japanese, Vietnamese, romaji] or [japanese, reading, romaji]
        if len(row)>=4:
            kanji, japanese, meaning, romaji = row[:4]
        elif len(row)==3:
            japanese, meaning, romaji = row
            kanji=''
        else:
            continue
        if not (japanese or meaning or romaji): continue
        vocab.append({
            'unit':unit, 'sub':sub, 'section':subkey,
            'kanji':kanji, 'japanese':japanese, 'meaning':meaning, 'romaji':romaji
        })

# Include any data in individual word files that might not occur in aggregate file.
# Deduplicate by (unit, japanese, meaning, romaji)
seen=set(); vocab2=[]
for x in vocab:
    k=(x['unit'],x['section'],x['japanese'],x['meaning'],x['romaji'])
    if k not in seen:
        seen.add(k); vocab2.append(x)
vocab=vocab2

# ---------- Kanji ----------
kanji_path=os.path.join(BASE,'Tổng hợp Kanji 4-7.docx')
kanji=[]
for rows in read_tables_xml(kanji_path):
    if not rows: continue
    first=' '.join(rows[0])
    m=re.search(r'UNIT\s*(\d+)', first, re.I)
    if not m: continue
    unit=int(m.group(1))
    # second row is heading; remaining rows are actual data
    for row in rows[2:]:
        if len(row)<4: continue
        hanviet, kj, hira, meaning = [x.strip() for x in row[:4]]
        if not any([hanviet,kj,hira,meaning]): continue
        kanji.append({'unit':unit,'hanviet':hanviet,'kanji':kj,'hiragana':hira,'meaning':meaning})

# ---------- Grammar ----------
def sectionize(lines):
    title=lines[0] if lines else ''
    body=lines[1:]
    sections=[]
    current=None
    pat=re.compile(r'^(\d+)[\.）\)]')
    for line in body:
        m=pat.match(line)
        if m:
            if current: sections.append(current)
            current={'no':int(m.group(1)),'title':line,'lines':[]}
        elif line.strip():
            if current is None:
                current={'no':0,'title':'Ghi chú','lines':[]}
            current['lines'].append(line)
    if current: sections.append(current)
    return {'title':title,'sections':sections}

grammar={}
for unit in range(4,8):
    path=os.path.join(BASE,f'Ngữ pháp {unit}.docx')
    grammar[str(unit)]=sectionize(read_paragraphs_xml(path))

# ---------- Labels / stats ----------
units=[]
for u in range(4,8):
    units.append({
        'id':u,
        'name':f'Bài {u}',
        'vocabCount':sum(1 for x in vocab if x['unit']==u),
        'kanjiCount':sum(1 for x in kanji if x['unit']==u),
        'grammarCount':len(grammar[str(u)]['sections'])
    })

data={'units':units,'vocab':vocab,'kanji':kanji,'grammar':grammar}

with open(os.path.join(OUT,'data.js'),'w',encoding='utf-8') as f:
    f.write('window.JP_DATA = ')
    json.dump(data,f,ensure_ascii=False,separators=(',',':'))
    f.write(';')

# ---------- HTML / CSS / JS ----------
html_doc=r'''<!doctype html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Japanese Study — Bài 4–7</title>
<link rel="stylesheet" href="styles.css">
</head>
<body>
<div class="app-shell">
  <header class="topbar">
    <div class="brand" onclick="showHome()"><span class="brand-icon">日</span><div><strong>Japanese Study</strong><small>Bài 4–7</small></div></div>
    <div class="top-actions">
      <button class="ghost" onclick="toggleSearch()">⌕ Tìm kiếm</button>
      <button class="ghost" onclick="showStats()">◫ Tiến độ</button>
    </div>
  </header>
  <div id="searchBar" class="searchbar hidden">
    <input id="searchInput" placeholder="Tìm từ vựng, Kanji, mẫu ngữ pháp..." oninput="renderSearch(this.value)">
  </div>
  <main id="app"></main>
  <footer>Japanese Study • Nội dung dựa trên tài liệu bạn cung cấp • Bài 4–7</footer>
</div>
<script src="data.js"></script>
<script src="app.js"></script>
</body>
</html>'''
open(os.path.join(OUT,'index.html'),'w',encoding='utf-8').write(html_doc)

css=r'''
:root{--bg:#fff7f8;--card:#fff;--text:#20283a;--muted:#6e7484;--line:#ffd1da;--pink:#f43d61;--purple:#9b4df3;--blue:#5965ed;--softpink:#fff3f6;--softpurple:#f9f1ff;--softblue:#f2f5ff;--green:#28a879;--shadow:0 10px 28px rgba(54,31,45,.08)}
*{box-sizing:border-box}body{margin:0;font-family:Inter,system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;background:linear-gradient(#fffafb,#fff7f8 35%,#fff);color:var(--text)}button,input{font:inherit}.app-shell{min-height:100vh;max-width:1480px;margin:auto}.topbar{height:74px;padding:12px 24px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid #f1dfe3;background:rgba(255,255,255,.82);backdrop-filter:blur(9px);position:sticky;top:0;z-index:20}.brand{display:flex;gap:11px;align-items:center;cursor:pointer}.brand-icon{width:42px;height:42px;border-radius:12px;display:grid;place-items:center;background:var(--pink);color:#fff;font-size:22px;font-weight:800}.brand strong{display:block;font-size:18px}.brand small{color:var(--muted)}.top-actions{display:flex;gap:8px}.ghost{border:1px solid #e9dfe4;background:#fff;border-radius:10px;padding:9px 13px;cursor:pointer}.ghost:hover{border-color:#f0a3b2}.searchbar{padding:12px 24px;background:#fff;border-bottom:1px solid #f1dfe3}.searchbar input{width:100%;max-width:760px;padding:12px 14px;border:1px solid #e3dfe4;border-radius:12px;outline:none}.hidden{display:none}main{padding:30px 24px 48px}.hero{text-align:center;padding:18px 0 24px}.hero h1{font-size:clamp(34px,4vw,54px);margin:0 0 8px}.hero p{color:var(--muted);margin:0}.unit-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:18px}.unit-card{background:#fff;border:1px solid #eadfe4;border-radius:18px;padding:18px;box-shadow:var(--shadow)}.unit-card h2{font-size:28px;margin:2px 0 14px;text-align:center}.unit-card .sub{color:var(--muted);text-align:center;font-size:13px;margin-top:-8px;margin-bottom:16px}.module{border:1px solid var(--line);border-radius:13px;padding:13px;margin:10px 0}.module.pink{background:var(--softpink)}.module.purple{border-color:#ead4ff;background:var(--softpurple)}.module.blue{border-color:#d5ddff;background:var(--softblue)}.module-title{font-size:17px;font-weight:800;margin-bottom:9px}.module-buttons{display:grid;gap:8px}.action{border:none;border-radius:10px;padding:10px 12px;font-weight:750;cursor:pointer;background:#fff;color:var(--text);border:1px solid #ecdfe4}.action.primary{color:#fff;background:var(--pink);border-color:var(--pink)}.action.purple{color:#fff;background:var(--purple);border-color:var(--purple)}.action.blue{color:#fff;background:var(--blue);border-color:var(--blue)}.page-head{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:20px}.back{border:1px solid #eadfe4;background:#fff;border-radius:10px;padding:9px 13px;cursor:pointer}.page-head h1{margin:0;font-size:34px}.stats{color:var(--muted);font-size:14px}.toolbar{display:flex;gap:9px;flex-wrap:wrap;margin:12px 0 18px}.pill{padding:8px 11px;border-radius:999px;border:1px solid #eadfe4;background:#fff;cursor:pointer}.pill.active{background:#20283a;color:#fff;border-color:#20283a}.table-wrap{overflow:auto;border:1px solid #eadfe4;border-radius:15px;background:#fff;box-shadow:var(--shadow)}table{width:100%;border-collapse:collapse;min-width:720px}th,td{text-align:left;padding:12px 14px;border-bottom:1px solid #f0e9ec}th{background:#fff7f8;position:sticky;top:0}.jp{font-size:22px}.muted{color:var(--muted)}.section-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}.grammar-card{background:#fff;border:1px solid #eadfe4;border-radius:15px;padding:17px;box-shadow:var(--shadow)}.grammar-card h3{margin:0 0 12px}.line{padding:7px 0;border-bottom:1px dashed #eee9ec;white-space:pre-wrap}.line:last-child{border-bottom:0}.jp-line{font-size:18px}.flash-wrap{display:grid;grid-template-columns:minmax(0,760px) 260px;gap:20px;align-items:start;justify-content:center}.flash{min-height:390px;background:#fff;border:1px solid #eadfe4;border-radius:22px;box-shadow:var(--shadow);display:grid;place-items:center;padding:30px;text-align:center;cursor:pointer;perspective:1000px}.flash-inner{transition:transform .45s;transform-style:preserve-3d;width:100%}.flash.flipped .flash-inner{transform:rotateY(180deg)}.face{backface-visibility:hidden}.backface{transform:rotateY(180deg);margin-top:-160px}.flash .jp-big{font-size:68px;font-weight:850;margin-bottom:14px}.flash .meaning{font-size:28px;font-weight:800}.flash .romaji{color:var(--muted);font-size:20px;margin-top:12px}.flash-side{background:#fff;border:1px solid #eadfe4;border-radius:15px;padding:15px}.flash-side button{width:100%;margin-bottom:8px}.progress-bar{height:8px;background:#eee9ec;border-radius:99px;overflow:hidden}.progress-bar>div{height:100%;background:var(--pink)}.quiz{max-width:880px;margin:auto}.quiz-card{background:#fff;border:1px solid #eadfe4;border-radius:18px;padding:22px;box-shadow:var(--shadow)}.q-label{color:var(--muted);font-size:14px}.question{font-size:26px;font-weight:800;margin:13px 0 18px}.options{display:grid;gap:10px}.opt{padding:14px;border:1px solid #e7dfe3;background:#fff;border-radius:12px;text-align:left;cursor:pointer}.opt:hover{border-color:#f195a7}.opt.correct{background:#e9faf3;border-color:#55bd95}.opt.wrong{background:#fff0f2;border-color:#f28196}.result{text-align:center;padding:30px}.score-big{font-size:50px;font-weight:900}.search-results{display:grid;gap:10px;max-width:900px;margin:auto}.result-card{background:#fff;border:1px solid #eadfe4;border-radius:12px;padding:13px;cursor:pointer}.result-card:hover{border-color:#e8a1b0}.empty{text-align:center;padding:50px;color:var(--muted)}.note{background:#fff8e7;border:1px solid #f3d990;border-radius:12px;padding:12px;margin:12px 0}.source{font-size:12px;color:#8a8390;margin-top:8px}footer{border-top:1px solid #f1dfe3;padding:18px 24px;color:#8a8390;font-size:13px;text-align:center}
@media(max-width:1100px){.unit-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.section-grid{grid-template-columns:1fr}.flash-wrap{grid-template-columns:1fr}}
@media(max-width:640px){.topbar{padding:10px 14px}.top-actions .ghost{padding:8px 9px;font-size:13px}.brand strong{font-size:15px}main{padding:20px 14px 38px}.unit-grid{grid-template-columns:1fr}.unit-card{padding:15px}.page-head{align-items:flex-start;flex-direction:column}.flash{min-height:340px}.flash .jp-big{font-size:48px}.question{font-size:21px}}
'''
open(os.path.join(OUT,'styles.css'),'w',encoding='utf-8').write(css)

js=r'''
const D=window.JP_DATA;
const $=s=>document.querySelector(s);
let state={unit:null,type:null,mode:null,idx:0,flipped:false,quiz:null};
const esc=s=>String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
const unitObj=u=>D.units.find(x=>x.id===Number(u));
function showHome(){state={unit:null,type:null,mode:null};renderHome();scrollTo(0,0)}
function renderHome(){
  $('#app').innerHTML=`<section class="hero"><h1>日本語 学習</h1><p>Học tiếng Nhật theo Bài 4–7 • Từ vựng · Kanji · Ngữ pháp</p></section>
  <div class="unit-grid">${D.units.map(u=>`
  <article class="unit-card"><h2>Bài ${u.id}</h2><div class="sub">${u.vocabCount} từ vựng · ${u.kanjiCount} Kanji · ${u.grammarCount} chủ đề</div>
    <div class="module pink"><div class="module-title">📖 Từ vựng</div><div class="module-buttons">
      <button class="action" onclick="openModule(${u.id},'vocab','list')">Danh sách</button><button class="action" onclick="openModule(${u.id},'vocab','flash')">Flashcard</button><button class="action primary" onclick="openModule(${u.id},'vocab','quiz')">Kiểm tra</button></div></div>
    <div class="module purple"><div class="module-title">🈶 Kanji</div><div class="module-buttons">
      <button class="action" onclick="openModule(${u.id},'kanji','list')">Danh sách</button><button class="action" onclick="openModule(${u.id},'kanji','flash')">Flashcard</button><button class="action purple" onclick="openModule(${u.id},'kanji','quiz')">Kiểm tra</button></div></div>
    <div class="module blue"><div class="module-title">⚖️ Ngữ pháp</div><div class="module-buttons">
      <button class="action blue" onclick="openModule(${u.id},'grammar','view')">Xem Ngữ pháp</button><button class="action" onclick="openModule(${u.id},'grammar','quiz')">Kiểm tra</button></div></div>
  </article>`).join('')}</div>`;
}
function openModule(unit,type,mode){state={unit,type,mode,idx:0,flipped:false,quiz:null};renderModule();scrollTo(0,0)}
function backUnit(){showHome()}
function dataFor(){if(state.type==='vocab')return D.vocab.filter(x=>x.unit===state.unit);if(state.type==='kanji')return D.kanji.filter(x=>x.unit===state.unit);return D.grammar[String(state.unit)].sections}
function titleFor(){return state.type==='vocab'?'Từ vựng':state.type==='kanji'?'Kanji':'Ngữ pháp'}
function renderModule(){
 const u=unitObj(state.unit); let controls=`<button class="back" onclick="backUnit()">← Danh sách bài</button>`;
 $('#app').innerHTML=`<div class="page-head"><div>${controls}</div><div><h1>Bài ${u.id} · ${titleFor()}</h1><div class="stats">${u.vocabCount} từ vựng · ${u.kanjiCount} Kanji · ${u.grammarCount} chủ đề</div></div><div></div></div>${moduleBody()}`;
}
function tabs(kind){return `<div class="toolbar"><button class="pill ${state.mode==='list'?'active':''}" onclick="openModule(${state.unit},'${kind}','list')">Danh sách</button><button class="pill ${state.mode==='flash'?'active':''}" onclick="openModule(${state.unit},'${kind}','flash')">Flashcard</button><button class="pill ${state.mode==='quiz'?'active':''}" onclick="openModule(${state.unit},'${kind}','quiz')">Kiểm tra</button></div>`}
function moduleBody(){
 if(state.type==='vocab') return tabs('vocab')+(state.mode==='list'?vocabList():state.mode==='flash'?flashView():quizView('vocab'));
 if(state.type==='kanji') return tabs('kanji')+(state.mode==='list'?kanjiList():state.mode==='flash'?flashView():quizView('kanji'));
 return `<div class="toolbar"><button class="pill ${state.mode==='view'?'active':''}" onclick="openModule(${state.unit},'grammar','view')">Xem ngữ pháp</button><button class="pill ${state.mode==='quiz'?'active':''}" onclick="openModule(${state.unit},'grammar','quiz')">Kiểm tra</button></div>`+(state.mode==='view'?grammarView():quizView('grammar'));
}
function vocabList(){const rows=dataFor();return `<div class="table-wrap"><table><thead><tr><th>#</th><th>Tiếng Nhật</th><th>Romaji</th><th>Nghĩa</th></tr></thead><tbody>${rows.map((x,i)=>`<tr><td>${i+1}</td><td><div class="jp">${esc(x.japanese||x.kanji||'—')}</div>${x.kanji?`<div class="muted">${esc(x.kanji)}</div>`:''}</td><td>${esc(x.romaji||'—')}</td><td>${esc(x.meaning||'—')}</td></tr>`).join('')}</tbody></table></div>`}
function kanjiList(){const rows=dataFor();return `<div class="table-wrap"><table><thead><tr><th>#</th><th>Kanji</th><th>Hán Việt</th><th>Hiragana</th><th>Ý nghĩa</th></tr></thead><tbody>${rows.map((x,i)=>`<tr><td>${i+1}</td><td class="jp">${esc(x.kanji||'—')}</td><td>${esc(x.hanviet||'—')}</td><td>${esc(x.hiragana||'—')}</td><td>${esc(x.meaning||'—')}</td></tr>`).join('')}</tbody></table></div>`}
function grammarView(){const secs=dataFor();return `<div class="section-grid">${secs.map(s=>`<article class="grammar-card"><h3>${esc(s.title)}</h3>${s.lines.map(line=>`<div class="line ${/[ぁ-んァ-ン一-龯]/.test(line)?'jp-line':''}">${esc(line)}</div>`).join('')}</article>`).join('')}</div><div class="note">Mẹo: ưu tiên đọc mẫu câu tiếng Nhật trước, sau đó xem phần giải thích tiếng Việt và tự đặt thêm 1 câu.</div>`}
function flashItems(){const rows=dataFor(); return state.type==='grammar'?[]:rows}
function flashView(){const rows=flashItems(); if(!rows.length)return '<div class="empty">Chưa có flashcard.</div>'; const x=rows[state.idx%rows.length]; const isK=state.type==='kanji'; const front=isK?(x.kanji||x.hiragana||x.hanviet):(x.japanese||x.kanji); const back=isK?`${x.hanviet||''} · ${x.hiragana||''}<br><span class="muted">${x.meaning||''}</span>`:`${x.meaning||''}<div class="romaji">${x.romaji||''}</div>`; return `<div class="flash-wrap"><div><div class="flash ${state.flipped?'flipped':''}" onclick="flipCard()"><div class="flash-inner"><div class="face"><div class="jp-big">${esc(front||'—')}</div><div class="muted">Nhấn để lật thẻ</div></div><div class="face backface"><div class="meaning">${back}</div></div></div></div><div class="toolbar" style="justify-content:center"><button class="pill" onclick="prevCard()">← Trước</button><button class="pill" onclick="nextCard()">Tiếp →</button><button class="pill" onclick="speakCard()">🔊 Phát âm</button></div></div><aside class="flash-side"><strong>Thẻ ${state.idx%rows.length+1}/${rows.length}</strong><div style="margin:10px 0" class="progress-bar"><div style="width:${((state.idx%rows.length+1)/rows.length)*100}%"></div></div><div class="muted">Lật thẻ để xem nghĩa. Dùng nút phát âm để nghe tiếng Nhật bằng giọng hệ thống.</div></aside></div>`}
function flipCard(){state.flipped=!state.flipped;renderModule()}
function nextCard(){const n=dataFor().length;state.idx=(state.idx+1)%n;state.flipped=false;renderModule()}
function prevCard(){const n=dataFor().length;state.idx=(state.idx-1+n)%n;state.flipped=false;renderModule()}
function speakCard(){const rows=dataFor(),x=rows[state.idx%rows.length];const text=state.type==='kanji'?(x.kanji||x.hiragana):(x.japanese||x.kanji); if('speechSynthesis'in window){const u=new SpeechSynthesisUtterance(text);u.lang='ja-JP';speechSynthesis.cancel();speechSynthesis.speak(u)}}
function quizView(kind){if(!state.quiz||state.quiz.kind!==kind) startQuiz(kind); const q=state.quiz; if(q.done)return `<div class="quiz"><div class="quiz-card result"><div class="score-big">${q.score}/${q.total}</div><h2>Hoàn thành!</h2><p class="muted">Bạn có thể làm lại để luyện tiếp.</p><button class="action primary" onclick="startQuiz('${kind}')">Làm lại</button></div></div>`; return `<div class="quiz"><div class="quiz-card"><div class="q-label">Câu ${q.pos+1}/${q.total} · ${titleFor()}</div><div class="question">${q.question}</div><div class="options">${q.options.map((o,i)=>`<button class="opt" ${q.answered?'disabled':''} onclick="answerQuiz(${i})">${o}</button>`).join('')}</div>${q.answered?`<div class="toolbar"><button class="action primary" onclick="nextQuiz()">${q.pos+1===q.total?'Xem kết quả':'Câu tiếp theo →'}</button></div>`:''}</div></div>`}
function startQuiz(kind){let items=[]; if(kind==='grammar'){items=D.grammar[String(state.unit)].sections.map(s=>({title:s.title,pattern:s.lines.find(x=>/[ぁ-んァ-ン一-龯]/.test(x))||s.lines[0]||''}));} else items=dataFor().slice(); shuffle(items); const total=Math.min(10,items.length);items=items.slice(0,total); state.quiz={kind,items,total,pos:0,score:0,answered:false,done:false};makeQuestion();renderModule()}
function makeQuestion(){const q=state.quiz, item=q.items[q.pos]; if(q.kind==='grammar'){const correct=item.title; const pool=D.grammar[String(state.unit)].sections.map(x=>x.title).filter(t=>t!==correct);const opts=shuffle([correct,...shuffle(pool).slice(0,3)]);q.question=`Mẫu câu dưới đây thuộc chủ đề nào?<br><span class="jp-line">${esc(item.pattern)}</span>`;q.options=opts;}else if(q.kind==='vocab'){const correct=item.meaning||item.japanese;const pool=dataFor().filter(x=>x!==item).map(x=>x.meaning||x.japanese);q.question=`Từ <span class="jp-line">${esc(item.japanese||item.kanji||'—')}</span> có nghĩa là gì?`;q.options=shuffle([correct,...shuffle(pool).slice(0,3)]);}else{const correct=item.kanji||item.hiragana;const pool=dataFor().filter(x=>x!==item).map(x=>x.kanji||x.hiragana||x.hanviet);q.question=`Từ Hán Việt <b>${esc(item.hanviet||item.meaning)}</b> tương ứng với Kanji nào?`;q.options=shuffle([correct,...shuffle(pool).slice(0,3)]);}}
function answerQuiz(i){const q=state.quiz;if(q.answered)return;const correct=q.kind==='grammar'?q.items[q.pos].title:(q.kind==='vocab'?(q.items[q.pos].meaning||q.items[q.pos].japanese):(q.items[q.pos].kanji||q.items[q.pos].hiragana));if(q.options[i]===correct)q.score++;q.answered=true;renderModule()}
function nextQuiz(){const q=state.quiz;if(q.pos+1>=q.total){q.done=true;saveBest();renderModule();return;}q.pos++;q.answered=false;makeQuestion();renderModule()}
function saveBest(){const key=`best_${state.unit}_${state.type}`;const old=Number(localStorage.getItem(key)||0);localStorage.setItem(key,String(Math.max(old,state.quiz.score)))}
function shuffle(a){return a.slice().sort(()=>Math.random()-0.5)}
function showStats(){const list=D.units.map(u=>{const a=['vocab','kanji','grammar'].map(t=>Number(localStorage.getItem(`best_${u.id}_${t}`)||0));return `<div class="result-card"><b>Bài ${u.id}</b><div class="muted">Từ vựng: ${a[0]} · Kanji: ${a[1]} · Ngữ pháp: ${a[2]}</div></div>`}).join('');$('#app').innerHTML=`<div class="page-head"><button class="back" onclick="showHome()">← Về trang chủ</button><div><h1>Tiến độ</h1><div class="stats">Điểm cao nhất được lưu trên trình duyệt này</div></div><div></div></div><div class="search-results">${list}</div>`;scrollTo(0,0)}
function toggleSearch(){const el=$('#searchBar');el.classList.toggle('hidden');if(!el.classList.contains('hidden'))$('#searchInput').focus()}
function renderSearch(term){const t=term.trim().toLowerCase(); if(!t){$('#app').innerHTML='<div class="empty">Nhập từ khóa để tìm.</div>';return} const res=[];for(const x of D.vocab){const hay=[x.japanese,x.romaji,x.meaning,x.kanji].join(' ').toLowerCase();if(hay.includes(t))res.push({type:'Từ vựng',unit:x.unit,title:x.japanese||x.kanji,sub:`${x.romaji||''} · ${x.meaning||''}`})}for(const x of D.kanji){const hay=[x.kanji,x.hanviet,x.hiragana,x.meaning].join(' ').toLowerCase();if(hay.includes(t))res.push({type:'Kanji',unit:x.unit,title:x.kanji||x.hanviet,sub:`${x.hanviet||''} · ${x.hiragana||''} · ${x.meaning||''}`})}for(const [u,g] of Object.entries(D.grammar)){for(const s of g.sections){const hay=[s.title,...s.lines].join(' ').toLowerCase();if(hay.includes(t))res.push({type:'Ngữ pháp',unit:Number(u),title:s.title,sub:s.lines.find(x=>/[ぁ-んァ-ン一-龯]/.test(x))||s.lines[0]||''})}} $('#app').innerHTML=`<div class="page-head"><div><h1>Kết quả tìm kiếm</h1><div class="stats">${res.length} kết quả</div></div><button class="back" onclick="showHome()">← Về trang chủ</button><div></div></div><div class="search-results">${res.length?res.slice(0,80).map(r=>`<div class="result-card" onclick="openSearchResult('${r.type}',${r.unit})"><b>${esc(r.title)}</b><div class="muted">${esc(r.type)} · Bài ${r.unit}</div><div>${esc(r.sub)}</div></div>`).join(''):'<div class="empty">Không tìm thấy.</div>'}</div>`}
function openSearchResult(type,unit){if(type==='Từ vựng'||type==='Kanji')openModule(unit,type==='Từ vựng'?'vocab':'kanji','list');else openModule(unit,'grammar','view')}
renderHome();
'''
open(os.path.join(OUT,'app.js'),'w',encoding='utf-8').write(js)

readme=r'''# Japanese Study — Bài 4–7

Website tĩnh học tiếng Nhật, thiết kế theo phong cách thẻ Bài / Từ vựng / Kanji / Ngữ pháp.

## Chạy thử trên máy
Mở `index.html` bằng trình duyệt. Hoặc chạy:

```bash
python -m http.server 8000
```

sau đó vào `http://localhost:8000`.

## Đưa lên GitHub Pages
1. Tạo một repository mới trên GitHub.
2. Upload toàn bộ `index.html`, `styles.css`, `app.js`, `data.js`.
3. Vào **Settings → Pages**.
4. Chọn **Deploy from a branch**, branch `main`, folder `/root`.
5. GitHub sẽ tạo một URL dạng `https://<username>.github.io/<repository>/`.

## Nội dung
Dữ liệu được xây từ các tài liệu Unit 4–7 mà người dùng cung cấp. Các phần quiz được tạo từ chính dữ liệu từ vựng/Kanji/ngữ pháp trong bộ tài liệu.
'''
open(os.path.join(OUT,'README.md'),'w',encoding='utf-8').write(readme)

# Create a zip
import shutil
zip_path='/mnt/data/Japanese_Study_Bai4-7.zip'
if os.path.exists(zip_path): os.remove(zip_path)
shutil.make_archive('/mnt/data/Japanese_Study_Bai4-7','zip',OUT)
print(json.dumps({'vocab':len(vocab),'kanji':len(kanji),'grammar':{k:len(v['sections']) for k,v in grammar.items()},'zip':zip_path},ensure_ascii=False))
