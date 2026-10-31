# -*- coding: utf-8 -*-
"""AcoustiLab : observations acoustiques et probabilités, outil pédagogique."""
from __future__ import annotations

import csv
import hashlib
import io
import json
import math
import re
import unicodedata
import zipfile
from datetime import date, datetime, time, timezone
from html import escape
from pathlib import Path

import altair as alt
import numpy as np
import pandas as pd
from scipy import stats
import streamlit as st
from openpyxl import Workbook

VERSION='1.0'
COLS=['id','date','heure','point','activite','protocole','type_niveau','niveau_db','duree_min','note']
LABELS={'id':'Observation','date':'Date','heure':'Heure de début','point':'Point de mesure','activite':'Activité','protocole':'Protocole / appareil','type_niveau':'Type de niveau','niveau_db':'Niveau en dB(A)','duree_min':'Durée (min)','note':'Remarque'}
LEQ='LAeq de période';SPOT='Lecture ponctuelle';TYPES=[LEQ,SPOT]
MAX_ROWS=10000;MAX_BYTES=5*1024*1024
BG='#14151D';PANEL='#1F202D';LINE='#383949';TEXT='#F5F3FF';MUTED='#C2BFD3';VIOLET='#AC8CFF';LIME='#D8F36A';PINK='#FF9DAE'
SOURCES={'INRS : décibels et addition énergétique':'https://www.inrs.fr/risques/bruit/definitions.html','ACNAW : définition du niveau équivalent':'https://acnaw.be/home/le-bruit/le-bruit-en-general/ses-indicateurs/niveau-equivalent--leqt.html','NIST : intervalle de Wilson':'https://www.itl.nist.gov/div898/handbook/prc/section2/prc241.htm'}


def fmt(x,d=2):
    if x is None or pd.isna(x):return 'Non calculable'
    return f'{float(x):,.{d}f}'.replace(',',' ').replace('.',',')


def key_name(v):
    return re.sub('[^a-z0-9]','',unicodedata.normalize('NFKD',str(v)).encode('ascii','ignore').decode().lower())


def standardize(frame):
    if not isinstance(frame,pd.DataFrame) or len(frame)>MAX_ROWS or len(frame.columns)>35:
        raise ValueError('Le tableau doit contenir au plus 10 000 lignes et 35 colonnes.')
    aliases={key_name(k):k for k in COLS}
    aliases.update({key_name(v):k for k,v in LABELS.items()})
    aliases.update(niveau='niveau_db',db='niveau_db',dba='niveau_db',duree='duree_min',identifiant='id',commentaire='note',type='type_niveau',lieu='point',heuredebut='heure')
    mapping={};ignored=[]
    for c in frame.columns:
        k=aliases.get(key_name(c))
        if k:
            if k in mapping.values():raise ValueError('Deux colonnes représentent la même variable : '+k)
            mapping[c]=k
        else:ignored.append(str(c))
    if not {'point','activite','type_niveau','niveau_db'}<=set(mapping.values()):
        raise ValueError('Colonnes indispensables : point, activite, type_niveau, niveau_db. Ajouter duree_min pour des LAeq.')
    d=frame.rename(columns=mapping).reindex(columns=COLS).copy()
    for k in COLS:
        def convert(v):
            if pd.isna(v):return ''
            if k=='date' and isinstance(v,(date,datetime,pd.Timestamp)):return v.strftime('%Y-%m-%d')
            if k=='heure' and isinstance(v,(time,datetime,pd.Timestamp)):return v.strftime('%H:%M')
            if k=='heure' and isinstance(v,str) and re.fullmatch(r'(?:[01]\d|2[0-3]):[0-5]\d:00',v.strip()):return v.strip()[:5]
            return str(v).strip()
        d[k]=d[k].map(convert)
        if d[k].str.encode('utf-8').str.len().gt(1200 if k=='note' else 180).any():raise ValueError('Texte trop long dans '+k+'.')
    if len(d.to_json(orient='records',force_ascii=False).encode())>10*1024*1024:raise ValueError('Carnet trop volumineux : réduire les remarques ou le nombre de lignes.')
    return d.reset_index(drop=True),ignored


def numbers(s):
    return pd.to_numeric(s.astype(str).str.replace('\u00a0','',regex=False).str.replace(' ','',regex=False).str.replace(',','.',regex=False),errors='coerce').replace([np.inf,-np.inf],np.nan)


def clean_data(raw,dedup=False):
    if not len(raw):return pd.DataFrame(columns=COLS+['ligne','instant','creneau']),pd.DataFrame(columns=['ligne','décision','motif']),dict(raw=0,valid=0,excluded=0,duplicates=0)
    d=raw.copy();errors=[[] for _ in range(len(d))]
    d['niveau_db']=numbers(raw.niveau_db);d['duree_min']=numbers(raw.duree_min)
    for i,row in raw.iterrows():
        if not row.point:errors[i].append('Point manquant')
        if not row.activite:errors[i].append('Activité manquante')
        if row.type_niveau not in TYPES:errors[i].append('Type attendu : '+LEQ+' ou '+SPOT)
        level=d.loc[i,'niveau_db'];duration=d.loc[i,'duree_min']
        if pd.isna(level) or not -50<=level<=200:errors[i].append('Niveau fini entre -50 et 200 dB(A) requis (limites informatiques)')
        if row.duree_min and (pd.isna(duration) or not 0<duration<=10080):errors[i].append('Durée positive, au plus 10 080 min')
        if row.type_niveau==LEQ and not row.duree_min:errors[i].append('Durée indispensable pour un LAeq de période')
        if row.type_niveau==SPOT and row.duree_min:errors[i].append('Laisser la durée vide pour une lecture ponctuelle')
        if row.date:
            try:
                if date.fromisoformat(row.date).isoformat()!=row.date:raise ValueError()
            except ValueError:errors[i].append('Date attendue AAAA-MM-JJ')
        if row.heure and not re.fullmatch(r'(?:[01]\d|2[0-3]):[0-5]\d',row.heure):errors[i].append('Heure attendue HH:MM')
    duplicate=raw.duplicated(keep='first').to_numpy()&np.array([not e for e in errors])
    if dedup:
        for i in np.flatnonzero(duplicate):errors[i].append('Doublon complet retiré sur demande')
    d['id']=[v or f'L-{i+1:04d}' for i,v in enumerate(d.id)];d['protocole']=d.protocole.replace('','Non renseigné')
    d['ligne']=np.arange(1,len(raw)+1)
    d['instant']=pd.to_datetime(raw.date+' '+raw.heure,format='%Y-%m-%d %H:%M',errors='coerce')
    def slot(h):
        if not re.fullmatch(r'(?:[01]\d|2[0-3]):[0-5]\d',h):return 'Heure inconnue'
        hr=int(h[:2]);return 'Matin · 06–12 h' if 6<=hr<12 else 'Après-midi · 12–18 h' if 12<=hr<18 else 'Soir · 18–22 h' if 18<=hr<22 else 'Nuit · 22–06 h'
    d['creneau']=raw.heure.map(slot)
    keep=np.array([not e for e in errors]);audit=raw.copy();audit.insert(0,'ligne',np.arange(1,len(raw)+1));audit['décision']=['Exclue' if e else 'Conservée' for e in errors];audit['motif']=[' ; '.join(e) for e in errors]
    return d.loc[keep].reset_index(drop=True),audit,dict(raw=len(raw),valid=int(keep.sum()),excluded=int((~keep).sum()),duplicates=int(duplicate.sum()))


def exceeds(values,target):
    return np.asarray(values,dtype=float)>float(target)


def describe(values,target=75.):
    x=np.asarray(values,dtype=float)
    if x.ndim!=1 or not np.isfinite(x).all():raise ValueError('Valeurs finies attendues.')
    n=len(x)
    if not n:return dict(n=0,mean=None,median=None,std=None,variance=None,min=None,max=None,q90=None,k=0,p=None)
    s=float(x.std(ddof=1)) if n>1 else None;k=int(exceeds(x,target).sum())
    return dict(n=n,mean=float(x.mean()),median=float(np.median(x)),std=s,variance=s*s if s is not None else None,min=float(x.min()),max=float(x.max()),q90=float(np.quantile(x,.9)),k=k,p=k/n)


def equivalent(levels,durations):
    x=np.asarray(levels,dtype=float);t=np.asarray(durations,dtype=float)
    if x.ndim!=1 or not len(x) or x.shape!=t.shape or not np.isfinite(x).all() or not np.isfinite(t).all() or (t<=0).any() or (np.abs(x)>300).any():raise ValueError('Niveaux finis et durées positives de même taille requis.')
    m=float(x.max());return m+10*math.log10(float(np.sum(t*10**((x-m)/10))/t.sum()))


def simultaneous(levels):
    x=np.asarray(levels,dtype=float)
    if x.ndim!=1 or not len(x) or not np.isfinite(x).all() or (np.abs(x)>300).any():raise ValueError('Au moins un niveau fini requis.')
    m=float(x.max());return m+10*math.log10(float(np.sum(10**((x-m)/10))))


def energy_status(df):
    if not len(df):return None,'Aucune observation.'
    if df.type_niveau.ne(LEQ).any():return None,'Le regroupement énergétique exige uniquement des LAeq de période.'
    if df.point.nunique()!=1 or df.protocole.nunique()!=1 or (df.protocole=='Non renseigné').any():return None,'Sélectionner un seul point et un protocole renseigné et constant.'
    if df.instant.isna().any():return None,'Date et heure de début indispensables pour vérifier les chevauchements.'
    ordered=df.sort_values('instant');end=None
    for row in ordered.itertuples():
        if end is not None and row.instant<end:return None,'Des périodes se chevauchent : le regroupement est bloqué.'
        end=row.instant+pd.Timedelta(minutes=row.duree_min)
    return equivalent(df.niveau_db,df.duree_min),'LAeq regroupé sur les seules périodes renseignées ; les intervalles non observés ne sont pas inclus.'


def group_table(df,target,by='activite'):
    rows=[]
    for label,g in df.groupby(by,sort=True):
        s=describe(g.niveau_db,target);rows.append({by:label,'n':s['n'],'moyenne_db':s['mean'],'mediane_db':s['median'],'ecart_type_db':s['std'],'maximum_db':s['max'],'depassements':s['k'],'frequence_pct':100*s['p']})
    return pd.DataFrame(rows,columns=[by,'n','moyenne_db','mediane_db','ecart_type_db','maximum_db','depassements','frequence_pct'])


def cohorts(df):
    known=df.loc[df.protocole.ne('Non renseigné')].copy()
    if not len(known):return []
    known['unite']=known.apply(lambda r:'ponctuel' if r.type_niveau==SPOT else repr(float(r.duree_min))+' min',axis=1)
    return [(f'{p} · {a} · {prot} · {kind} · {unit}',g.drop(columns='unite')) for (p,a,prot,kind,unit),g in known.groupby(['point','activite','protocole','type_niveau','unite'],sort=True)]


def infer(values,target,future=10,minimum=6):
    s=describe(values,target)
    if not s['n'] or not 1<=future<=500 or not 0<=minimum<=future:raise ValueError('Effectif et paramètres binomiaux invalides.')
    ci=stats.binomtest(s['k'],s['n']).proportion_ci(confidence_level=.95,method='wilson')
    return dict(n=s['n'],k=s['k'],p=s['p'],low=float(ci.low),high=float(ci.high),future=int(future),minimum=int(minimum),tail=float(stats.binom.sf(minimum-1,future,s['p'])))


@st.cache_data(ttl=1800,max_entries=8,show_spinner=False)
def simulate(p,n,minimum,repeats=10000,seed=42):
    if not np.isfinite(p) or not 0<=p<=1 or not 1<=n<=500 or not 0<=minimum<=n or repeats not in [1000,10000,50000] or not 0<=seed<=2147483647:raise ValueError('Paramètres de simulation invalides.')
    values=np.random.default_rng(seed).binomial(n,p,size=repeats);rate=float(np.mean(values>=minimum));se=math.sqrt(rate*(1-rate)/repeats)
    return dict(values=values,rate=rate,se=se,repeats=repeats,seed=seed)


def demo_data():
    rows=[]
    for i,x in enumerate([68,70,72,74,76,78,80,82]):
        rows.append(dict(id=f'F-{i+1:02d}',date='2026-10-31',heure=f'{8+i//4:02d}:{(i%4)*15:02d}',point='P-01 · exemple',activite='Préparation' if i%2==0 else 'Terrassement',protocole='FICTIF-01',type_niveau=LEQ,niveau_db=str(x),duree_min='10',note='Observation entièrement fictive pour démonstration.'))
    return pd.DataFrame(rows,columns=COLS)


def metadata(name,kind):return dict(name=name,kind=kind,created=datetime.now(timezone.utc).isoformat(timespec='seconds'))


def activate(raw,meta,settings=None):
    st.session_state.al_raw=raw;st.session_state.al_meta=meta;st.session_state.al_revision+=1
    st.session_state.al_defaults=settings or dict(target=75.,points=[],activities=[],dedup=False)
    for k in ['al_risk','al_sim','al_report']:st.session_state.pop(k,None)


def load_demo():activate(demo_data(),metadata('Huit observations de démonstration','Exemple fictif'))


def signature(ctx):
    payload=dict(raw=ctx['raw'].to_dict('records'),meta=ctx['meta'],settings=ctx['settings'])
    return hashlib.sha256(json.dumps(payload,sort_keys=True,ensure_ascii=False).encode()).hexdigest()


def safe_csv(df):
    d=df.copy()
    for c in d.select_dtypes(include=['object','string']).columns:
        def safe(v):
            if c in ['niveau_db','duree_min'] and isinstance(v,str):
                try:
                    if math.isfinite(float(v.replace(',','.'))):return v
                except ValueError:pass
            return "'"+v if isinstance(v,str) and v.lstrip().startswith(('=','+','-','@')) else v
        d[c]=d[c].map(safe)
    return d.to_csv(index=False,sep=';',decimal=',',float_format='%.12g').encode('utf-8-sig')


def read_table(data,name):
    if not data or len(data)>MAX_BYTES:raise ValueError('Fichier vide ou supérieur à 5 Mo.')
    suffix=Path(name).suffix.lower()
    if suffix=='.xlsx':
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                if len(z.infolist())>500 or sum(e.file_size for e in z.infolist())>30*1024*1024:raise ValueError('Classeur trop volumineux après décompression.')
            frame=pd.read_excel(io.BytesIO(data),dtype=object,engine='openpyxl',nrows=MAX_ROWS+1)
        except ValueError:raise
        except Exception as e:raise ValueError('Classeur XLSX illisible.') from e
    elif suffix=='.csv':
        try:text=data.decode('utf-8-sig')
        except UnicodeDecodeError:text=data.decode('cp1252')
        try:
            dialect=csv.Sniffer().sniff(text[:8192],delimiters=';,\t');reader=csv.reader(io.StringIO(text),dialect);header=next(reader)
            if len(set(header))!=len(header):raise ValueError('En-têtes dupliqués.')
            frame=pd.read_csv(io.StringIO(text),sep=dialect.delimiter,dtype=str,keep_default_na=False,nrows=MAX_ROWS+1)
        except Exception as e:raise ValueError('CSV illisible : vérifier séparateur, en-têtes et cellules.') from e
    else:raise ValueError('Formats acceptés : CSV ou XLSX.')
    return standardize(frame)


def backup_bytes(ctx):
    return json.dumps(dict(version=1,raw=ctx['raw'].to_dict('records'),meta=ctx['meta'],settings=ctx['settings']),ensure_ascii=False,allow_nan=False,indent=2).encode()


def restore(data):
    if not data or len(data)>12*1024*1024:raise ValueError('Sauvegarde vide ou supérieure à 12 Mo.')
    try:
        obj=json.loads(data);assert obj['version']==1 and isinstance(obj['raw'],list) and len(obj['raw'])<=MAX_ROWS
        assert all(isinstance(r,dict) and set(r)==set(COLS) and all(isinstance(v,str) for v in r.values()) for r in obj['raw'])
        raw,_=standardize(pd.DataFrame(obj['raw'],columns=COLS));settings=obj['settings'];meta=obj['meta']
        assert set(settings)=={'target','points','activities','dedup'} and isinstance(settings['dedup'],bool)
        assert type(settings['target']) in [int,float] and math.isfinite(settings['target']) and -50<=settings['target']<=200
        assert all(isinstance(settings[k],list) and len(settings[k])<=MAX_ROWS and all(isinstance(v,str) and len(v)<=180 for v in settings[k]) for k in ['points','activities'])
        assert set(meta)=={'name','kind','created'} and all(isinstance(v,str) and len(v)<=300 for v in meta.values())
        clean,_,_=clean_data(raw);assert set(settings['points'])<=set(clean.point) and set(settings['activities'])<=set(clean.activite)
        if meta['kind']=='Exemple fictif' and not raw.equals(demo_data()):meta['kind']='Origine déclarée dans le fichier'
        return raw,meta,settings
    except Exception as e:raise ValueError('Sauvegarde AcoustiLab invalide ou réglages incohérents.') from e


def excel_bytes(ctx):
    wb=Workbook();wb.remove(wb.active)
    frames={'Selection':ctx['df'].reindex(columns=COLS),'Activites':group_table(ctx['df'],ctx['settings']['target']),'Nettoyage':ctx['audit'],'Brut':ctx['raw'],'Contexte':pd.DataFrame([{'champ':k,'valeur':json.dumps(v,ensure_ascii=False) if isinstance(v,(list,dict)) else str(v)} for k,v in {**ctx['meta'],**ctx['settings']}.items()])}
    for name,frame in frames.items():
        ws=wb.create_sheet(name)
        for row in [list(frame.columns)]+frame.astype(object).where(pd.notna(frame),None).values.tolist():
            ws.append(row)
            for cell,v in zip(ws[ws.max_row],row):
                if isinstance(v,str):cell.data_type='s'
        ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
        for col in ws.columns:ws.column_dimensions[col[0].column_letter].width=min(46,max(15,len(str(col[0].value))+2))
    out=io.BytesIO();wb.save(out);return out.getvalue()


def chart(c,alt_text,height=290):
    c=c.properties(height=height).configure(background=BG,view={'stroke':None},axis={'labelColor':MUTED,'titleColor':TEXT,'gridColor':'#30313D','domainColor':LINE,'labelFontSize':11,'titlePadding':12},legend={'labelColor':TEXT,'titleColor':TEXT,'orient':'bottom'})
    st.altair_chart(c,theme=None,width='stretch',alt=alt_text)


def threshold_line(target):return alt.Chart(pd.DataFrame({'objectif':[target]})).mark_rule(color=LIME,strokeDash=[5,4]).encode(y='objectif:Q')


def note(title,text):st.html('<aside class="al-note"><strong>'+escape(title)+'</strong><p>'+escape(text)+'</p></aside>')


def heading(number,title,description):st.html('<div class="al-heading"><span>'+escape(number)+'</span><h1>'+escape(title)+'</h1><p>'+escape(description)+'</p></div>')


def kpis(items):
    st.html('<div class="al-kpis">'+''.join('<article><span>'+escape(label)+'</span><b>'+escape(value)+'</b><small>'+escape(caption)+'</small></article>' for label,value,caption in items)+'</div>')


def current():return st.session_state.al_context


def data_view(df,alt_text):
    st.dataframe(df.rename(columns=LABELS),hide_index=True,width='stretch',alt=alt_text)


def no_data():
    if not len(current()['df']):
        note('Commencer par un carnet','Charger les huit observations fictives ou saisir vos observations. Si le carnet existe déjà, vérifier les filtres et le journal de nettoyage.')
        st.button('Charger les 8 observations fictives',on_click=load_demo,type='primary',icon=':material/science:')
        return True
    return False


def dashboard():
    ctx=current();d=ctx['df'];target=ctx['settings']['target']
    waveform=''.join(f'<i style="--h:{v}%"></i>' for v in [12,23,38,62,42,72,92,57,80,100,69,44,82,57,36,18,34,60,45,23,12,28,44,70,48,25,15])
    st.html('<section class="al-hero"><div><span class="al-eyebrow">ACOUSTIQUE DE CHANTIER / CARNET D’OBSERVATIONS</span><h1>Donner une lecture<br>claire au bruit.</h1><p>Observer. Comparer. Comprendre l’incertitude.</p><div class="al-tag">AcoustiLab · Laboratoire pédagogique</div></div><div class="al-wave" aria-hidden="true">'+waveform+'</div></section>')
    if not len(ctx['raw']):
        c1,c2=st.columns([1.2,1])
        with c1:
            st.subheader('Un premier résultat en un clic')
            st.markdown('Huit observations fictives, deux activités et un objectif d’exercice de **75 dB(A)**. Aucune donnée de démonstration n’est ajoutée automatiquement à votre carnet.')
            st.button('Charger les 8 observations fictives',on_click=load_demo,type='primary',icon=':material/play_arrow:',wrap=True)
        with c2:
            note('Votre carnet, votre point de départ','Saisir des niveaux en dB(A), préciser le point, l’activité, le type de mesure et le protocole. Les LAeq demandent une durée.')
            st.button('Saisir mes observations',on_click=lambda:st.switch_page(st.session_state.al_pages['observations']),icon=':material/edit_note:',wrap=True)
        kpis([('01','Observer','Saisie et import CSV / Excel'),('02','Comparer','Activités, points et créneaux'),('03','Estimer','Fréquences et incertitudes'),('04','Présenter','Rapport PDF et sauvegarde')]);return
    if no_data():return
    s=describe(d.niveau_db,target)
    kpis([('Observations',str(s['n']),f'{d.point.nunique()} point(s) · {d.activite.nunique()} activité(s)'),('Niveau médian',fmt(s['median'])+' dB(A)','La moitié des valeurs de chaque côté'),('Au-dessus de l’objectif',f"{s['k']} / {s['n']}",fmt(100*s['p'],1)+' % des observations'),('Maximum observé',fmt(s['max'])+' dB(A)','Maximum des valeurs saisies, pas LAmax mesuré')])
    c1,c2=st.columns([1.55,1])
    with c1:
        st.subheader('La séquence des observations')
        seq=d.copy();seq['rang']=np.arange(1,len(d)+1);seq['Lecture']=np.where(exceeds(seq.niveau_db,target),'Au-dessus','Au plus à l’objectif')
        base=alt.Chart(seq).encode(x=alt.X('rang:Q',title='Ordre dans le carnet',axis=alt.Axis(tickMinStep=1)),y=alt.Y('niveau_db:Q',title='Niveau en dB(A)',scale=alt.Scale(zero=False)))
        chart(base.mark_line(color=VIOLET,opacity=.65)+base.mark_circle(size=110).encode(color=alt.Color('Lecture:N',scale=alt.Scale(domain=['Au-dessus','Au plus à l’objectif'],range=[PINK,VIOLET])),tooltip=[alt.Tooltip('id:N',title='Observation'),alt.Tooltip('activite:N',title='Activité'),alt.Tooltip('niveau_db:Q',title='dB(A)',format='.2f')])+threshold_line(target),'Valeurs observées dans l’ordre du carnet et objectif choisi',320)
        st.caption('La ligne relie les observations dans l’ordre du carnet ; elle ne représente pas un enregistrement sonore continu.')
    with c2:
        st.subheader('Lire l’objectif')
        st.html('<div class="al-objective"><span>OBJECTIF D’EXERCICE</span><b>'+fmt(target,1)+'<small> dB(A)</small></b><p>Dépassement strict : niveau &gt; objectif.<br>L’égalité n’est pas un dépassement.</p></div>')
        st.markdown('**'+str(s['k'])+' observation(s)** appellent une lecture plus attentive : vérifier le contexte, la comparabilité et la qualité de mesure avant de proposer une action.')
    st.subheader('Les activités du carnet')
    groups=group_table(d,target)
    for start in range(0,min(len(groups),12),3):
        cols=st.columns(min(3,len(groups)-start))
        for col,(_,row) in zip(cols,groups.iloc[start:start+3].iterrows()):
            with col,st.container(border=True):
                st.markdown('**'+escape(str(row.activite))+'**');st.metric('Médiane',fmt(row.mediane_db)+' dB(A)');st.caption(f"{row.n} observations · {row.depassements} dépassement(s)")
    if len(groups)>12:st.caption('Les 12 premières activités sont affichées ici ; toutes restent disponibles dans Comparer.')
    note('Une observation n’est pas une durée d’exposition','50 % d’observations au-dessus d’un objectif ne signifie pas 50 % du temps. Les mesures doivent être représentatives pour aller au-delà du carnet.')


def observations():
    heading('01 / COLLECTE','Un carnet qui garde le contexte.','Saisir, importer ou corriger les données ; chaque exclusion reste visible dans le journal.')
    ctx=current();raw=ctx['raw']
    with st.expander('Ajouter une observation',expanded=True):
        kind=st.segmented_control('Nature de la mesure',TYPES,default=LEQ,key='al_kind',persist_state='session')
        with st.form('al_add'):
            a,b=st.columns(2)
            with a:
                ident=st.text_input('Identifiant',max_chars=120)
                point=st.text_input('Point de mesure *',max_chars=120,placeholder='P-01 · limite du chantier')
                activity=st.text_input('Activité *',max_chars=120,placeholder='Terrassement')
                protocol=st.text_input('Protocole / appareil',max_chars=120,help='Identifiant de campagne et réglages comparables. Ne pas inventer une mesure étalonnée.')
            with b:
                day=st.date_input('Date de l’observation',value=date.today())
                hour=st.time_input('Heure de début',value=time(8,0),step=60)
                level=st.number_input('Niveau en dB(A)',min_value=-50.,max_value=200.,value=70.,step=1.)
                duration=st.number_input('Durée (min)',min_value=.01,max_value=10080.,value=10.,step=1.,disabled=kind!=LEQ,help='Pour un LAeq : durée réelle de la période intégrée. Vide pour une lecture ponctuelle.')
            comment=st.text_area('Remarque',max_chars=1000,height=85)
            submitted=st.form_submit_button('Enregistrer l’observation',type='primary',icon=':material/add:')
        if submitted:
            try:
                row=dict(id=ident,date=day.isoformat(),heure=hour.strftime('%H:%M'),point=point,activite=activity,protocole=protocol,type_niveau=kind,niveau_db=str(level),duree_min=str(duration) if kind==LEQ else '',note=comment)
                frame,_=standardize(pd.DataFrame([row]));_,audit,q=clean_data(frame)
                if q['excluded']:raise ValueError(audit.iloc[0]['motif'])
                combined,_=standardize(pd.concat([raw,frame],ignore_index=True));activate(combined,metadata('Carnet saisi ou modifié par l’utilisateur','Données utilisateur'));st.session_state.al_flash='Observation ajoutée.';st.rerun()
            except ValueError as e:st.error(str(e))
    with st.expander('Importer un CSV ou un classeur XLSX'):
        st.caption('5 Mo maximum ; première feuille Excel ; un import remplace le carnet après confirmation. Les cellules invalides ne sont pas corrigées automatiquement.')
        st.download_button('Télécharger le modèle CSV fictif',safe_csv(demo_data()),'acoustilab-exemple-fictif.csv','text/csv',on_click='ignore')
        upload=st.file_uploader('Fichier d’observations',type=['csv','xlsx'],key='al_upload')
        if upload:
            try:
                incoming,ignored=read_table(upload.getvalue(),upload.name);valid,audit,q=clean_data(incoming)
                st.markdown(f"**{q['raw']} lignes ; {q['valid']} analysables ; {q['excluded']} exclues.**")
                if ignored:st.caption('Colonnes ignorées : '+', '.join(ignored))
                data_view(incoming.head(25),'Aperçu du fichier avant remplacement')
                if q['excluded']:data_view(audit.loc[audit['décision']=='Exclue'].head(25),'Lignes exclues dans le fichier importé')
                confirm=st.checkbox('Remplacer le carnet actif par ce fichier',key='al_confirm_'+hashlib.sha256(upload.getvalue()).hexdigest()[:12])
                if st.button('Utiliser ce fichier',disabled=not confirm or not q['valid'],type='primary'):
                    activate(incoming,metadata(upload.name[:180],'Fichier utilisateur'));st.session_state.al_flash='Import effectué.';st.rerun()
            except (ValueError,UnicodeError) as e:st.error(str(e))
    if len(raw):
        st.subheader('Contrôle de la qualité')
        q=ctx['quality'];st.caption(f"{q['raw']} lignes brutes · {q['valid']} conservées avant filtres · {q['excluded']} exclues · {q['duplicates']} doublon(s) complet(s) détecté(s).")
        with st.expander('Journal complet du nettoyage'):
            data_view(ctx['audit'],'Journal des décisions de nettoyage');st.download_button('Exporter le journal CSV',safe_csv(ctx['audit']),'acoustilab-nettoyage.csv','text/csv',on_click='ignore')
        with st.expander('Modifier ou supprimer des lignes'):
            if len(raw)<=1500:
                with st.form('al_edit_'+str(st.session_state.al_revision)):
                    edited=st.data_editor(raw,num_rows='dynamic',hide_index=True,width='stretch',column_config={c:st.column_config.TextColumn(LABELS[c]) for c in COLS},alt='Carnet brut modifiable')
                    yes=st.checkbox('Appliquer ces modifications au carnet')
                    save=st.form_submit_button('Enregistrer les modifications')
                if save:
                    if not yes:st.error('Cocher la confirmation des modifications.')
                    else:
                        try:
                            normalized,_=standardize(edited);activate(normalized,metadata('Carnet corrigé par l’utilisateur','Données utilisateur'));st.session_state.al_flash='Modifications enregistrées ; vérifier le journal.';st.rerun()
                        except ValueError as e:st.error(str(e))
            else:st.caption('Au-delà de 1 500 lignes, corriger le CSV puis le réimporter.')
        st.download_button('Exporter les données brutes CSV',safe_csv(raw),'acoustilab-brut.csv','text/csv',on_click='ignore')
    note('Le type de niveau compte','Un LAeq décrit une énergie moyenne pendant une période. Une lecture ponctuelle ne suffit pas pour calculer l’énergie de dix minutes. Les deux types restent identifiés et séparés pour l’inférence.')


def calculators():
    st.subheader('L’atelier des décibels')
    mode=st.segmented_control('Calcul acoustique',['Périodes successives','Sources simultanées'],default='Périodes successives',key='al_calc_mode',persist_state='session')
    if mode=='Périodes successives':
        st.caption('Entrer des niveaux LAeq de périodes successives au même point. Les durées peuvent être différentes. Les intervalles non observés restent hors calcul.')
        frame=pd.DataFrame({'niveau_db':[60.,70.],'duree_min':[60.,60.]})
        with st.form('al_energy'):
            data=st.data_editor(frame,num_rows='dynamic',hide_index=True,key='al_periods',column_config={'niveau_db':st.column_config.NumberColumn('LAeq en dB(A)',min_value=-50,max_value=200),'duree_min':st.column_config.NumberColumn('Durée (min)',min_value=.01,max_value=10080)},alt='Périodes successives et durées pour le calcul énergétique')
            run=st.form_submit_button('Calculer le niveau équivalent',type='primary')
        if run:
            try:
                if not 1<=len(data)<=50:raise ValueError('Saisir entre 1 et 50 périodes.')
                result=equivalent(data.niveau_db,data.duree_min);st.success('Niveau équivalent : '+fmt(result)+' dB(A), sur '+fmt(data.duree_min.sum(),1)+' min renseignées.')
                st.latex(r'L_{Aeq}=10\log_{10}\left(\frac{\sum_i t_i\,10^{L_i/10}}{\sum_i t_i}\right)')
                st.caption('Test attendu : 60 et 70 dB(A), pendant 60 min chacun → 67,40 dB(A). La moyenne arithmétique serait 65 dB(A).')
            except (ValueError,TypeError) as e:st.error(str(e))
    else:
        st.caption('Niveaux attribuables à chaque source au même récepteur, sur une même période ; sources supposées non corrélées et bruit de fond non compté plusieurs fois. Aucun calcul de propagation ou d’écran.')
        with st.form('al_sources'):
            data=st.data_editor(pd.DataFrame({'niveau_db':[60.,60.]}),num_rows='dynamic',hide_index=True,key='al_sources_edit',column_config={'niveau_db':st.column_config.NumberColumn('Niveau propre de la source en dB(A)',min_value=-50,max_value=200)},alt='Sources non corrélées au même point récepteur')
            run=st.form_submit_button('Combiner les sources',type='primary')
        if run:
            try:
                if not 1<=len(data)<=50:raise ValueError('Saisir entre 1 et 50 sources.')
                result=simultaneous(data.niveau_db);st.success('Niveau combiné : '+fmt(result)+' dB(A).');st.latex(r'L_{total}=10\log_{10}\left(\sum_i10^{L_i/10}\right)');st.caption('Test : deux contributions de 60 dB(A) donnent 63,01 dB(A).')
            except (ValueError,TypeError) as e:st.error(str(e))


def compare():
    heading('02 / COMPARAISONS & CALCULS','Comparer des contextes, expliquer les écarts.','Les écarts décrivent le carnet. Une différence entre activités ne prouve pas que l’activité est la seule cause.')
    ctx=current();d=ctx['df'];target=ctx['settings']['target']
    if len(d):
        by=st.selectbox('Regrouper les observations par',['activite','point','creneau'],format_func=lambda k:dict(activite='Activité',point='Point de mesure',creneau='Créneau horaire')[k],key='al_compare_by',persist_state='session')
        groups=group_table(d,target,by)
        shown=groups[by].head(20).tolist();plotted=d.loc[d[by].isin(shown)]
        rule=alt.Chart(pd.DataFrame({'objectif':[target]})).mark_rule(color=LIME,strokeDash=[5,4]).encode(x='objectif:Q')
        chart(alt.Chart(plotted).mark_boxplot(color=VIOLET,extent=1.5,size=22).encode(y=alt.Y(by+':N',title=None,axis=alt.Axis(labelLimit=135)),x=alt.X('niveau_db:Q',title='Niveau en dB(A)',scale=alt.Scale(zero=False)))+rule,'Dispersion des niveaux par groupe et objectif',min(620,max(260,len(shown)*30)))
        if len(groups)>20:st.caption('Graphique limité aux 20 premiers groupes par ordre alphabétique ; le tableau et les exports conservent tous les groupes.')
        st.dataframe(groups.rename(columns={by:'Groupe','n':'Observations','moyenne_db':'Moyenne arithmétique (dB)','mediane_db':'Médiane (dB)','ecart_type_db':'Écart-type (dB)','maximum_db':'Maximum saisi (dB)','depassements':'Dépassements','frequence_pct':'Fréquence (%)'}),hide_index=True,alt='Statistiques descriptives par groupe')
        if len(groups)>=2:
            opts=groups[by].tolist();a,b=st.columns(2)
            left=a.selectbox('Groupe A',opts,index=0);right=b.selectbox('Groupe B',opts,index=1)
            av=float(groups.loc[groups[by]==left,'moyenne_db'].iloc[0]);bv=float(groups.loc[groups[by]==right,'moyenne_db'].iloc[0])
            st.metric('Écart des moyennes arithmétiques B − A',fmt(bv-av)+' dB');st.caption('Moyennes des valeurs en dB ; elles ne sont pas des LAeq regroupés. Vérifier point, durées, appareil et conditions avant de comparer.')
        with st.expander('Regroupement énergétique du carnet'):
            result,reason=energy_status(d);st.markdown(reason)
            if result is not None:st.metric('LAeq des périodes observées',fmt(result)+' dB(A)');st.caption('Durée totale intégrée : '+fmt(d.duree_min.sum(),1)+' min. Ce résultat n’est pas automatiquement celui d’une journée entière.')
    else:note('Comparaison en attente','Charger ou saisir des observations pour comparer les groupes. Les calculateurs ci-dessous fonctionnent indépendamment du carnet.')
    calculators()


def probabilities():
    heading('03 / PROBABILITÉS','Du constat au modèle, en explicitant les hypothèses.','La fréquence observée décrit les lignes du carnet ; une projection suppose un échantillonnage adapté et des observations comparables.')
    st.session_state.pop('al_risk',None)
    if no_data():return
    ctx=current();d=ctx['df'];target=ctx['settings']['target'];s=describe(d.niveau_db,target)
    kpis([('Dépassements stricts',f"{s['k']} / {s['n']}",'Niveau > '+fmt(target,1)+' dB(A)'),('Fréquence observée',fmt(100*s['p'],1)+' %','Sur la sélection entière'),('Moyenne arithmétique',fmt(s['mean'])+' dB','Statistique des valeurs, pas LAeq'),('Écart-type',fmt(s['std'])+' dB','Diviseur n−1 ; au moins 2 valeurs')])
    chart(alt.Chart(d).mark_bar(color=VIOLET).encode(x=alt.X('niveau_db:Q',bin=alt.Bin(maxbins=14),title='Niveau en dB(A)'),y=alt.Y('count():Q',title='Observations'))+alt.Chart(pd.DataFrame({'objectif':[target]})).mark_rule(color=LIME,strokeDash=[5,4]).encode(x='objectif:Q'),'Distribution des observations et objectif exploratoire',250)
    st.subheader('Étudier un groupe comparable')
    groups=cohorts(d)
    if not groups:note('Un protocole est nécessaire','Renseigner le protocole ou la campagne pour former un groupe. Aucun modèle n’est ajusté sur des références inconnues.');return
    labels=[g[0] for g in groups];chosen=st.selectbox('Point, activité, protocole et durée identiques',labels,index=0 if len(labels)==1 else None,placeholder='Choisir un groupe à étudier')
    if chosen is None:return
    group=groups[labels.index(chosen)][1];st.caption(str(len(group))+' observation(s). Les mêmes étiquettes ne garantissent pas des conditions identiques. Des valeurs proches dans le temps peuvent être dépendantes.')
    approved=st.checkbox('Supposer ces observations comparables, indépendantes et représentatives',key='al_assumptions_'+hashlib.sha256((signature(ctx)+chosen).encode()).hexdigest()[:12],wrap=True)
    if not approved:return
    if len(group)<2:note('Échantillon insuffisant pour cette démonstration','Ajouter au moins une seconde observation comparable. Un petit échantillon laisse une forte incertitude.');return
    r=infer(group.niveau_db,target)
    kpis([('Fréquence du groupe',fmt(100*r['p'],1)+' %',f"{r['k']} / {r['n']} observations"),('Intervalle de Wilson à 95 %',fmt(100*r['low'],1)+' – '+fmt(100*r['high'],1)+' %','Incertitude d’estimation de la proportion')])
    note('Lire l’intervalle','Sous les hypothèses retenues, la méthode vise une couverture de 95 % au fil d’échantillonnages répétés. Ce n’est ni une plage contenant 95 % des niveaux sonores, ni une garantie pour les prochaines observations.')
    st.subheader('Projection binomiale et simulation')
    st.markdown('On suppose que la probabilité de dépassement reste **p = '+fmt(r['p'],3)+'**, estimée par la fréquence du groupe. On compte les dépassements dans une nouvelle série indépendante. L’incertitude sur p n’est pas propagée dans cette projection.')
    with st.form('al_sim_form'):
        a,b=st.columns(2)
        n=a.number_input('Nombre de futures observations',min_value=1,max_value=500,value=10,step=1)
        minimum=b.number_input('Au moins combien de dépassements ?',min_value=0,max_value=500,value=6,step=1)
        repeats=st.select_slider('Nombre de séries simulées',[1000,10000,50000],value=10000)
        seed=st.number_input('Graine aléatoire',min_value=0,max_value=2147483647,value=42,step=1)
        run=st.form_submit_button('Calculer et simuler',type='primary',icon=':material/casino:')
    ssig=signature(ctx)+chosen
    if run:
        if minimum>n:st.error('Le nombre minimal de dépassements doit être inférieur ou égal au nombre d’observations.')
        else:
            r=infer(group.niveau_db,target,int(n),int(minimum));sim=simulate(r['p'],int(n),int(minimum),int(repeats),int(seed));st.session_state.al_sim=dict(signature=ssig,result=r,sim=sim)
    saved=st.session_state.get('al_sim')
    if saved and saved['signature']==ssig:
        r=saved['result'];sim=saved['sim']
        kpis([('Calcul binomial',fmt(100*r['tail'],2)+' %',f"P(K ≥ {r['minimum']}) pour {r['future']} observations"),('Estimation Monte-Carlo',fmt(100*sim['rate'],2)+' %',f"{sim['repeats']} séries · graine {sim['seed']}"),('Erreur de simulation',fmt(100*sim['se'],2)+' point(s)','Écart-type estimé de l’estimation Monte-Carlo')])
        counts=np.bincount(sim['values'],minlength=r['future']+1);frame=pd.DataFrame({'Dépassements':np.arange(r['future']+1),'Fréquence simulée':counts/sim['repeats']});frame['Probabilité binomiale']=stats.binom.pmf(frame['Dépassements'],r['future'],r['p'])
        chart(alt.Chart(frame).mark_bar(color=VIOLET,opacity=.75).encode(x=alt.X('Dépassements:Q',axis=alt.Axis(tickMinStep=1)),y=alt.Y('Fréquence simulée:Q',axis=alt.Axis(format='%')))+alt.Chart(frame).mark_line(color=LIME,point=True).encode(x='Dépassements:Q',y=alt.Y('Probabilité binomiale:Q',title='Fréquence / probabilité',axis=alt.Axis(format='%'))),'Nombre de dépassements : fréquences simulées et loi binomiale',260)
        st.caption('Barres violettes : simulation. Ligne citron : calcul binomial. Une estimation aléatoire peut varier ; l’erreur Monte-Carlo ne mesure pas le biais d’échantillonnage ni l’erreur de mesure.')
    st.session_state.al_risk=dict(signature=signature(ctx),group=chosen,result=r)


def pdf_bytes(ctx,title,student,conclusion,risk=None):
    from reportlab import rl_config
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate,Paragraph,Spacer,Table,TableStyle,PageBreak,KeepTogether
    from reportlab.graphics.shapes import Drawing,Rect,String,Line
    fonts=Path(rl_config.__file__).parent/'fonts'
    for name,file in [('AL','Vera.ttf'),('AL-Bold','VeraBd.ttf')]:
        if name not in pdfmetrics.getRegisteredFontNames():pdfmetrics.registerFont(TTFont(name,str(fonts/file)))
    pdfmetrics.registerFontFamily('AL',normal='AL',bold='AL-Bold',italic='AL',boldItalic='AL-Bold')
    styles=getSampleStyleSheet();styles.add(ParagraphStyle(name='ALBody',fontName='AL',fontSize=9,leading=13,spaceAfter=9));styles.add(ParagraphStyle(name='ALHead',fontName='AL-Bold',fontSize=14,leading=18,textColor=colors.HexColor('#5E3AA4'),spaceBefore=9,spaceAfter=12,keepWithNext=True));styles.add(ParagraphStyle(name='ALTitle',fontName='AL-Bold',fontSize=29,leading=34,spaceAfter=10));styles.add(ParagraphStyle(name='ALSmall',fontName='AL',fontSize=7.5,leading=11,spaceAfter=7))
    out=io.BytesIO();items=[]
    def p(text,style='ALBody'):return Paragraph(escape(str(text)).replace('\n','<br/>'),styles[style])
    def para(text,style='ALBody'):items.append(p(text,style))
    def table(rows,widths):
        t=Table([[p(v,'ALSmall') for v in row] for row in rows],colWidths=widths,repeatRows=1,hAlign='LEFT');t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#E6DDF8')),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.HexColor('#F8F6FC'),colors.white]),('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),7),('RIGHTPADDING',(0,0),(-1,-1),7),('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),4)]));items.append(t);items.append(Spacer(1,10))
    def figure(labels,values,target):
        drawing=Drawing(520,225);drawing.add(Rect(0,0,520,225,fillColor=colors.HexColor('#F1EDF7'),strokeColor=None))
        if not len(values):drawing.add(String(25,115,'Aucune observation.',fontName='AL',fontSize=11));return drawing
        low=min(0,min(values)-5,target-5);high=max(max(values),target)+8;scale=155/(high-low);step=460/len(values)
        for i in range(5):
            v=low+(high-low)*i/4;y=38+(v-low)*scale;drawing.add(Line(45,y,510,y,strokeColor=colors.HexColor('#CEC7DC'),strokeWidth=.4));drawing.add(String(4,y-3,fmt(v,1),fontName='AL',fontSize=7))
        for i,(label,v) in enumerate(zip(labels,values)):
            y=38+(v-low)*scale;drawing.add(Rect(46+i*step,38,max(1,step*.65),max(.5,y-38),fillColor=colors.HexColor('#A84A66' if v>target else '#7651BA'),strokeColor=None));text=str(label)
            while text and pdfmetrics.stringWidth(text,'AL',6)>step-3:text=text[:-1]
            drawing.add(String(46+i*step,24,text,fontName='AL',fontSize=6))
        y=38+(target-low)*scale;drawing.add(Line(45,y,510,y,strokeColor=colors.HexColor('#505F12'),strokeWidth=1,strokeDashArray=[4,3]));drawing.add(String(45,207,'Niveau en dB(A) · ligne pointillée = objectif',fontName='AL',fontSize=8));return drawing
    d=ctx['df'];target=ctx['settings']['target'];s=describe(d.niveau_db,target)
    para('AcoustiLab','ALTitle');para(title,'ALHead');para('Étudiant / groupe : '+(student or 'Non renseigné'));para('Édité le '+datetime.now(timezone.utc).strftime('%Y-%m-%d à %H:%M UTC')+' · Version '+VERSION,'ALSmall')
    para('1. Provenance et périmètre','ALHead');para(ctx['meta']['name']+' · '+ctx['meta']['kind']);para(f"Lignes brutes : {len(ctx['raw'])} ; analysables avant filtres : {len(ctx['clean'])} ; sélection : {len(d)}. Exclusions : {ctx['quality']['excluded']}.")
    para('Points : '+(', '.join(ctx['settings']['points']) or 'tous')+' ; activités : '+(', '.join(ctx['settings']['activities']) or 'toutes')+' ; doublons retirés : '+('oui' if ctx['settings']['dedup'] else 'non')+'.')
    para('Objectif exploratoire : '+fmt(target)+' dB(A). Dépassement strict : niveau > objectif ; l’égalité n’est pas comptée. Cet objectif n’est pas une limite réglementaire.')
    table([['n','Médiane dB','Moyenne arith. dB','Écart-type dB','Dépassements'],[str(s['n']),fmt(s['median']),fmt(s['mean']),fmt(s['std']),str(s['k'])+'/'+str(s['n'])]],[45,95,135,110,135])
    para('Fréquence observée : '+fmt(None if s['p'] is None else 100*s['p'])+' %. Elle décrit les observations et non une proportion de temps d’exposition.')
    para('2. Bien identifier la mesure','ALHead');para('Moyenne, médiane et écart-type décrivent les nombres en dB. La moyenne arithmétique n’est pas un LAeq. Le maximum saisi n’est pas nécessairement un LAmax mesuré. Aucune mesure n’est créée par l’application.')
    para('Pour des LAeq successifs et comparables : LAeq = 10 log10[ somme(ti x 10^(Li/10)) / somme(ti) ]. Pour des contributions simultanées non corrélées au même récepteur : Ltotal = 10 log10[ somme(10^(Li/10)) ].')
    result,reason=energy_status(d);para(reason+(' Résultat : '+fmt(result)+' dB(A).' if result is not None else ''))
    items.append(PageBreak());para('3. Comparaison des activités','ALHead');g=group_table(d,target)
    if len(g):table([['Activité','n','Médiane dB','Moyenne dB','Au-dessus']]+[[r.activite,str(r.n),fmt(r.mediane_db),fmt(r.moyenne_db),str(r.depassements)] for r in g.head(30).itertuples()],[175,35,100,110,100])
    else:para('Aucune activité dans la sélection.')
    if len(g)>30:para('Tableau limité aux 30 premières activités ; les exports contiennent toute la sélection.')
    para('Moyennes arithmétiques par activité · 12 premières','ALHead');items.append(figure(g.activite.head(12).tolist(),g.moyenne_db.head(12).tolist(),target));para('Les différences sont descriptives. Point de mesure, durée, distance, appareil et conditions peuvent expliquer une partie des écarts. Aucune relation causale n’est démontrée.')
    items.append(PageBreak());para('4. Observations et interprétation','ALHead');items.append(figure(d.id.head(25).tolist(),d.niveau_db.head(25).tolist(),target));para('Les 25 premières observations sont dessinées dans l’ordre du carnet. Les indicateurs portent sur toute la sélection. Le dessin ne constitue pas une mesure continue.','ALSmall')
    if risk:
        r=risk['result'];para('Estimation probabiliste retenue','ALHead');para('Groupe : '+risk['group']);para(f"{r['n']} observations ; {r['k']} dépassements. Fréquence : {fmt(100*r['p'])} %. Intervalle de Wilson à 95 % : {fmt(100*r['low'])} à {fmt(100*r['high'])} %.")
        para('Hypothèses activées : comparabilité, indépendance, représentativité. L’intervalle concerne une proportion et non une plage de niveaux sonores. Les simulations et projections conditionnelles sont consultables dans l’application ; elles ne sont pas incluses dans ce rapport.','ALSmall')
    para('Conclusion relue par l’utilisateur','ALHead');para(conclusion)
    source_start=len(items);para('Sources et limites','ALHead')
    for label,url in SOURCES.items():para(label+' : '+url,'ALSmall')
    para('Les huit observations de démonstration sont fictives. L’outil ne certifie ni une conformité acoustique ni un risque sanitaire. Les données réelles demandent un appareil et un protocole adaptés. Aucun calcul de propagation ni d’exposition réglementaire n’est effectué.','ALSmall')
    items[source_start:]=[KeepTogether(items[source_start:])]
    def footer(canvas,doc):
        canvas.setFont('AL',8);canvas.setFillColor(colors.HexColor('#5E3AA4'));canvas.drawString(36,22,'AcoustiLab · Analyse pédagogique');canvas.drawRightString(A4[0]-36,22,str(doc.page))
    SimpleDocTemplate(out,pagesize=A4,leftMargin=36,rightMargin=36,topMargin=35,bottomMargin=45,title=title,author=student or 'AcoustiLab').build(items,onFirstPage=footer,onLaterPages=footer)
    return out.getvalue()


def methods():
    with st.expander('Test guidé : résultats attendus',expanded=True):
        st.markdown('1. Charger les **8 observations fictives** ; objectif **75 dB(A)**.\n2. Retrouver **4 / 8 dépassements**, soit **50 %**, et une médiane de **75 dB(A)**.\n3. Moyenne arithmétique : **75 dB** ; préparation **74 dB**, terrassement **76 dB** ; écart **2 dB**.\n4. Dans Comparer : périodes de 60 et 70 dB(A), 60 min chacune → **67,40 dB(A)**. Sources simultanées de 60 et 60 → **63,01 dB(A)**.\n5. Dans Probabilités : choisir une activité et activer les hypothèses. Fréquence du groupe **2 / 4 = 50 %** ; Wilson ≈ **15,00–85,00 %**.\n6. Projeter **10** observations, au moins **6** dépassements : probabilité binomiale **37,6953 %**. Une simulation de 10 000 séries doit être voisine, avec environ 0,5 point d’erreur-type.\n7. Préparer un rapport et sauvegarder le carnet JSON. À objectif **76 dB(A)** : **3 / 8** dépassements, soit **37,5 %**.')
        st.button('Charger ce test fictif',on_click=load_demo,type='primary',icon=':material/science:')
    with st.expander('Formules, hypothèses et limites'):
        st.markdown('**Fréquence** = nombre de valeurs strictement au-dessus de l’objectif / nombre de valeurs. **Variance d’échantillon** : somme des écarts à la moyenne au carré, divisée par n−1. **Écart-type** : racine de la variance. Le CV des valeurs en dB n’est pas utilisé : zéro dB n’est pas une absence d’énergie.')
        st.latex(r'\widehat p=\frac{k}{n},\qquad K\sim\mathrm{Binomiale}(m,\widehat p)')
        st.markdown('La projection fixe p à la fréquence observée : elle ne propage pas l’incertitude sur p. Aucune loi normale n’est imposée aux niveaux. Les groupes probabilistes utilisent même point, activité, protocole, type et durée de LAeq ; cela ne garantit pas l’indépendance. Les répétitions proches peuvent être dépendantes. Les créneaux horaires sont des regroupements descriptifs choisis pour l’exercice, pas des périodes réglementaires.')
        st.markdown('Ne pas confondre une lecture ponctuelle, un LAeq et un LAmax. Les calculs énergétiques ne donnent pas automatiquement un niveau journalier. Aucune analyse fréquentielle, mesure par microphone ou carte acoustique de propagation n’est réalisée.')
        for label,url in SOURCES.items():st.markdown(f'[{label}]({url})')


def reports():
    heading('04 / RAPPORTS & MÉTHODES','Des résultats que vous pouvez présenter.','Exporter une analyse lisible et conserver le carnet pour reprendre le travail plus tard.')
    ctx=current();sig=signature(ctx)
    if len(ctx['raw']):
        with st.container(horizontal=True):
            st.download_button('Sauvegarder le carnet JSON',backup_bytes(ctx),'acoustilab-carnet.json','application/json',on_click='ignore',icon=':material/save:',width='stretch',wrap=True)
            st.download_button('Sélection CSV',safe_csv(ctx['df'].reindex(columns=COLS)),'acoustilab-selection.csv','text/csv',on_click='ignore',icon=':material/download:',width='stretch',wrap=True)
        title=st.text_input('Titre du rapport',value='Analyse des observations acoustiques',max_chars=180,key='al_title',persist_state='session')
        student=st.text_input('Étudiant / groupe (facultatif)',max_chars=120,key='al_student',persist_state='session')
        s=describe(ctx['df'].niveau_db,ctx['settings']['target'])
        default=f"La sélection comprend {s['n']} observation(s). La médiane est de {fmt(s['median'])} dB(A). {s['k']} observation(s) dépassent strictement l’objectif exploratoire de {fmt(ctx['settings']['target'])} dB(A), soit {fmt(None if s['p'] is None else 100*s['p'])} %. Cette fréquence décrit le carnet ; elle n’est ni une durée d’exposition ni une validation de conformité."
        conclusion=st.text_area('Conclusion modifiable',value=default,height=150,max_chars=3000,key='al_conclusion_'+sig[:12],persist_state='session')
        risk=st.session_state.get('al_risk');risk=risk if risk and risk['signature']==sig else None
        include=st.checkbox('Inclure l’estimation probabiliste étudiée',disabled=risk is None,key='al_include',persist_state='session')
        options_sig=hashlib.sha256(json.dumps([sig,title,student,conclusion,risk if include else None],sort_keys=True).encode()).hexdigest()
        if st.button('Préparer le PDF et Excel',type='primary',icon=':material/picture_as_pdf:'):
            if not title.strip():st.error('Donner un titre au rapport.')
            else:
                with st.spinner('Préparation des documents…'):st.session_state.al_report=(options_sig,pdf_bytes(ctx,title,student,conclusion,risk if include else None),excel_bytes(ctx))
        prepared=st.session_state.get('al_report')
        if prepared and prepared[0]==options_sig:
            with st.container(horizontal=True):
                st.download_button('Télécharger le PDF',prepared[1],'acoustilab-rapport.pdf','application/pdf',on_click='ignore',width='stretch')
                st.download_button('Télécharger Excel',prepared[2],'acoustilab-analyse.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',on_click='ignore',width='stretch')
        st.caption('Excel : Selection, Activites, Nettoyage, Brut, Contexte. PDF : résultats, graphiques, hypothèses et sources. JSON : carnet et filtres, sans simulations ni brouillon de rapport. Les données restent dans la session.')
    with st.expander('Restaurer un carnet sauvegardé'):
        upload=st.file_uploader('Sauvegarde AcoustiLab JSON',type=['json'],key='al_json')
        if upload:
            try:
                raw,meta,settings=restore(upload.getvalue());st.caption(f"{len(raw)} lignes ; source déclarée : {meta['name']}.")
                yes=st.checkbox('Remplacer le carnet et ses réglages par cette sauvegarde',key='al_restore_'+hashlib.sha256(upload.getvalue()).hexdigest()[:12])
                if st.button('Restaurer le carnet',disabled=not yes,type='primary'):
                    activate(raw,meta,settings);st.session_state.al_flash='Carnet restauré.';st.rerun()
            except ValueError as e:st.error(str(e))
    methods()


def style():
    st.html('''<style>
    .stApp{background:#14151D;color:#F5F3FF}[data-testid="stMainBlockContainer"]{padding:4.7rem 2.6rem 2.7rem;max-width:1600px}
    [data-testid="stHeader"]{background:#14151DF2;color:#F5F3FF}[data-testid="stHeader"] button{color:#F5F3FF}
    [data-testid="stSidebar"]{background:#1B1B27;color:#F5F3FF;border-right:1px solid #373546}[data-testid="stSidebar"] *{scrollbar-color:#6F5B9A #1B1B27}
    [data-testid="stSidebarNav"] a{color:#DCD8EE;border-radius:10px;padding:.65rem .8rem;margin:.16rem 0}[data-testid="stSidebarNav"] a[aria-current="page"]{background:#3B2D62;color:white;border-left:3px solid #D8F36A}
    [data-testid="stMarkdownContainer"], [data-testid="stWidgetLabel"], [data-testid="stCaptionContainer"], [data-testid="stMetricLabel"], [data-testid="stMetricValue"]{color:inherit}
    .stApp h1,.stApp h2,.stApp h3,.stApp label{color:#F5F3FF}.stApp h2{font-size:25px}.stApp h3{font-size:21px}.stApp [data-testid="stCaptionContainer"]{color:#C2BFD3}.stApp a{color:#C8AEFF}
    .stApp [data-testid="stExpander"]{border-color:#484052;background:#1B1C28;border-radius:12px}.stApp [data-testid="stExpander"] summary{color:#F5F3FF}.stApp [data-testid="stForm"],.stApp [data-testid="stVerticalBlockBorderWrapper"]{border-color:#484052}
    .stApp input,.stApp textarea{color:#1A1427!important;-webkit-text-fill-color:#1A1427!important;background:#F4F1FC!important}
    .stApp [data-baseweb="input"],.stApp [data-baseweb="textarea"],.stApp [data-baseweb="select"]>div{background:#F4F1FC!important;color:#1A1427!important;border-radius:8px}
    .stApp [data-baseweb="select"] span{color:#1A1427}.stApp input::placeholder,.stApp textarea::placeholder{color:#70657D!important;-webkit-text-fill-color:#70657D!important}
    .stApp [data-testid="stBaseButton-secondary"],.stApp [data-testid="stBaseButton-minimal"],.stApp [data-testid="stBaseButton-pills"]{background:#29243A;border-color:#736387;color:#F5F3FF}
    .stApp [data-testid="stBaseButton-primary"]{background:#7044CB;color:#FFFFFF;border:1px solid #B996FF;border-radius:9px;font-weight:650}.stApp button:disabled{opacity:.5}
    .stApp [data-testid="stFileUploaderDropzone"]{background:#29243A;color:#F5F3FF}.stApp [data-testid="stFileUploaderDropzone"] small{color:#D3CDE1}
    .stApp [data-testid="stMetric"]{background:#252331;border-radius:10px;padding:13px}.stApp [data-testid="stMetric"] label{color:#D8CEE8}
    .al-brand{padding:15px 0 23px;display:flex;align-items:center;gap:12px}.al-brand>span{display:grid;place-items:center;border:1px solid #B392FF;color:#D8F36A;font-size:26px;letter-spacing:-5px;font-weight:700;background:#302344;width:48px;height:48px;border-radius:14px}.al-brand strong{font-size:25px;letter-spacing:-1px;color:#FFFFFF}.al-brand small{display:block;font-size:10px;letter-spacing:1.5px;color:#D1BFEE;margin-top:3px}
    .al-topline{display:flex;justify-content:space-between;align-items:center;gap:15px;margin-bottom:23px;font-size:11px;letter-spacing:1.4px;color:#C9C2DB}.al-chip{font-size:11px;color:#D8F36A;background:#2D3421;border:1px solid #566137;border-radius:30px;padding:7px 12px;letter-spacing:0;overflow-wrap:anywhere}
    .al-hero{background:linear-gradient(115deg,#342245,#211D34 60%,#292238);border:1px solid #685079;border-radius:23px;padding:33px 36px;display:grid;grid-template-columns:1.4fr 1fr;gap:25px;align-items:center;margin:8px 0 30px;position:relative;overflow:hidden}.al-hero:after{content:'';position:absolute;width:220px;height:220px;border:1px solid #9072BC30;border-radius:50%;right:-55px;top:-60px;pointer-events:none}
    .al-eyebrow{font-size:10px;letter-spacing:2px;color:#DECEEE;font-weight:650}.al-hero h1{font-family:'Segoe UI',Arial,sans-serif;font-size:48px;font-weight:750;letter-spacing:-2px;line-height:1.08;margin:19px 0 13px;padding:0;color:#FFFFFF}.al-hero p{font-size:15px;color:#E0D4ED;margin:0}.al-tag{display:inline-block;font-size:11px;color:#D8F36A;padding:7px 11px;border:1px solid #938E5B70;border-radius:30px;margin-top:22px}.al-wave{display:flex;align-items:center;justify-content:center;gap:5px;height:180px;border-bottom:1px solid #AC8CFF55;background:repeating-linear-gradient(0deg,transparent,transparent 29px,#8C67BA20 30px),repeating-linear-gradient(90deg,transparent,transparent 29px,#8C67BA20 30px);padding:10px 12px;z-index:1}.al-wave i{display:block;flex:1;height:var(--h);min-width:2px;max-width:7px;border-radius:4px;background:linear-gradient(#D8F36A,#AB7DE2)}
    .al-kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin:20px 0 25px}.al-kpis article{background:#222230;border:1px solid #434055;border-radius:15px;padding:19px 20px;min-width:0}.al-kpis article:nth-child(2){background:#D8F36A;border-color:#D8F36A}.al-kpis article:nth-child(2) span,.al-kpis article:nth-child(2) b,.al-kpis article:nth-child(2) small{color:#222A12}.al-kpis article:nth-child(3){background:#352A47;border-color:#8D6AB0}.al-kpis span{display:block;color:#D2CBE0;font-size:12px}.al-kpis b{display:block;font-size:28px;line-height:1.2;color:#FFFFFF;letter-spacing:-.7px;margin:12px 0;overflow-wrap:anywhere}.al-kpis small{display:block;font-size:11px;line-height:1.6;color:#D2CBE0}
    .al-objective{border:1px solid #5D565F;border-radius:16px;padding:22px;background:#252829}.al-objective>span{font-size:10px;letter-spacing:2px;color:#D8F36A}.al-objective b{display:block;color:#D8F36A;font-size:55px;font-weight:600;letter-spacing:-2px;margin:9px 0}.al-objective b small{font-size:17px;letter-spacing:0}.al-objective p{font-size:13px;line-height:1.7;color:#DCE0D2}
    .al-heading>span{font-size:11px;letter-spacing:2px;color:#D8F36A}.al-heading h1{font-size:36px;line-height:1.12;letter-spacing:-1px;margin:14px 0}.al-heading p{font-size:15px;color:#CFC7DD;line-height:1.7;margin:0 0 22px;max-width:900px}.al-note{border:1px solid #625274;border-left:3px solid #AC8CFF;border-radius:11px;background:#272331;padding:17px 20px;margin:15px 0}.al-note strong{font-size:14px;color:#EBDEFF}.al-note p{font-size:13px;line-height:1.75;color:#D0C8DC;margin:8px 0 0}
    @media(max-width:1100px){.al-kpis{grid-template-columns:repeat(2,minmax(0,1fr))}.al-hero{grid-template-columns:1.3fr .8fr;padding:27px}.al-hero h1{font-size:39px}.al-wave{gap:3px}}
    @media(max-width:640px){[data-testid="stMainBlockContainer"]{padding:4.8rem 1rem 2rem}.al-topline{align-items:flex-start;flex-direction:column;gap:10px}.al-hero{grid-template-columns:1fr;padding:24px 21px;gap:20px;border-radius:17px}.al-hero h1{font-size:38px;letter-spacing:-1.6px}.al-wave{height:85px}.al-eyebrow{font-size:9px;line-height:1.7;display:block}.al-kpis article{padding:16px 12px}.al-kpis b{font-size:23px}.al-heading h1{font-size:29px}.al-objective b{font-size:44px}.al-note{padding:14px 15px}}
    @media(max-width:350px){.al-kpis{grid-template-columns:1fr}.al-hero h1{font-size:33px}}
    [data-testid="stSidebarNav"] a p,[data-testid="stSidebarNav"] a span,[data-testid="stSidebarNav"] a svg,[data-testid="stSidebarNav"] a{color:#E2DCEC!important}
    [data-testid="stSidebarNav"] a[aria-current="page"] p,[data-testid="stSidebarNav"] a[aria-current="page"] span{color:#FFFFFF!important}
    .stApp [data-testid="stCaptionContainer"] p{color:#C2BFD3!important}
    .stApp [data-testid="stExpander"] summary{background:#30273E!important;color:#F5F3FF!important}
    .stApp [data-testid="stExpander"] summary p,.stApp [data-testid="stExpander"] summary span{color:#F5F3FF!important}
    .stApp [data-testid="stFormSubmitButton"] button{background:#7044CB!important;border:1px solid #B996FF!important;color:#FFFFFF!important}
    .stApp [data-testid="stButtonGroup"] button{background:#29243A;color:#F5F3FF;border-color:#736387}
    .stApp [data-testid="stButtonGroup"] button[aria-pressed="true"]{background:#7044CB;color:#FFFFFF;border-color:#B996FF}
    [data-testid="stExpandSidebarButton"] svg,[data-testid="stSidebarCollapseButton"] svg{color:#DFD7EF!important;fill:currentColor!important}
    .stApp button[data-variant="primary"]{background:#7044CB!important;color:#FFFFFF!important;border-color:#B996FF!important}
    .stApp button[data-variant="secondary"]{background:#29243A!important;color:#F5F3FF!important;border-color:#736387!important}
    .stApp button[data-variant="segmented_control"][aria-checked="true"]{background:#7044CB!important;color:#FFFFFF!important;border-color:#B996FF!important}
    .stApp button[data-variant="segmented_control"][aria-checked="true"] p{color:#FFFFFF!important}
    </style>''')


def main():
    st.set_page_config(page_title='AcoustiLab · Observatoire du bruit',page_icon=':material/graphic_eq:',layout='wide',initial_sidebar_state='auto')
    for k,v in [('al_raw',pd.DataFrame(columns=COLS)),('al_meta',metadata('Aucun carnet','Aucune source')),('al_revision',0),('al_defaults',dict(target=75.,points=[],activities=[],dedup=False))]:
        if k not in st.session_state:st.session_state[k]=v
    style()
    with st.sidebar:st.html('<div class="al-brand"><span>)))</span><div><strong>AcoustiLab</strong><small>OBSERVER LE BRUIT</small></div></div>')
    pages=[st.Page(dashboard,title='Tableau de bord',icon=':material/space_dashboard:',default=True,url_path='accueil'),st.Page(observations,title='Observations',icon=':material/edit_note:',url_path='observations'),st.Page(compare,title='Comparer et calculer',icon=':material/graphic_eq:',url_path='comparer'),st.Page(probabilities,title='Probabilités',icon=':material/casino:',url_path='probabilites'),st.Page(reports,title='Rapports et méthodes',icon=':material/description:',url_path='rapports')]
    st.session_state.al_pages={p.url_path:p for p in pages};nav=st.navigation(pages,expanded=True)
    raw=st.session_state.al_raw;rev=str(st.session_state.al_revision);defaults=st.session_state.al_defaults;clean,audit,q=clean_data(raw)
    with st.sidebar:
        st.subheader('Périmètre d’analyse')
        target=st.number_input('Objectif exploratoire en dB(A)',min_value=-50.,max_value=200.,value=float(defaults['target']),step=1.,key='al_target_'+rev,help='75 dB(A) est un objectif choisi pour l’exercice, sans valeur réglementaire.')
        points=st.multiselect('Points de mesure',sorted(clean.point.unique()) if len(clean) else [],default=[v for v in defaults['points'] if v in clean.point.values],placeholder='Tous les points',key='al_points_'+rev)
        activities=st.multiselect('Activités',sorted(clean.activite.unique()) if len(clean) else [],default=[v for v in defaults['activities'] if v in clean.activite.values],placeholder='Toutes les activités',key='al_activities_'+rev)
        dedup=st.checkbox('Retirer les doublons complets',value=defaults['dedup'],key='al_dedup_'+rev,wrap=True)
        st.caption('Une sélection vide signifie tous les points ou toutes les activités. Objectif commun aux cinq rubriques.')
        st.caption('Version '+VERSION+' · Projet pédagogique')
    if dedup:clean,audit,q=clean_data(raw,True)
    d=clean.loc[clean.point.isin(points)].copy() if points else clean.copy()
    if activities:d=d.loc[d.activite.isin(activities)].copy()
    settings=dict(target=target,points=points,activities=activities,dedup=dedup);ctx=dict(raw=raw,clean=clean,df=d,audit=audit,quality=q,meta=st.session_state.al_meta,settings=settings);st.session_state.al_context=ctx
    st.html('<div class="al-topline"><span>LABORATOIRE / ANALYSE ACOUSTIQUE</span><div class="al-chip">'+escape(ctx['meta']['kind'])+' · '+str(len(d))+' observation(s)</div></div>')
    if 'al_flash' in st.session_state:st.success(st.session_state.pop('al_flash'))
    if len(raw):st.caption(ctx['meta']['name']+' · Sauvegarder le JSON pour conserver le carnet entre les sessions.')
    nav.run()


if __name__=='__main__':main()
