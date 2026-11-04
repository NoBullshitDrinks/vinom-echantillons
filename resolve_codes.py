# -*- coding: utf-8 -*-
"""
Résolveur de codes articles hors-tarif — version "fournisseur d'abord".
Règle d'or : ne JAMAIS attribuer un code à un fournisseur qui ne le possède pas.
Pipeline :
  1. Identifie le FOURNISSEUR via son nom complet + raison sociale (fournisseurs_vinistoria.csv).
  2. Cherche le code UNIQUEMENT dans les articles de ce fournisseur (catalogue_vinistoria.csv).
  3. Si la cuvée n'est pas chez ce fournisseur, signale qu'un article similaire existe ailleurs
     (sans l'imposer) et laisse le code vide.
Deux fichiers de référence attendus à côté du script :
  catalogue_vinistoria.csv   (nf, code, des1, des2, compl, app, mill, coul)
  fournisseurs_vinistoria.csv (nf, nom)
"""
import pandas as pd, re, unicodedata, os
from collections import Counter

# Overrides domaine -> N° fournisseur, pour les cas absents de la base fournisseurs.
ALIAS_FOURN = {'dufouleur':'727'}

def _norm(s):
    s=unicodedata.normalize('NFKD',str(s or '')).encode('ascii','ignore').decode().lower()
    return re.sub(r'[^a-z0-9]+',' ',s).strip()
STOP={'rouge','blanc','blc','rose','ros','rge','75cl','cl','vin','vins','igp','aop','aoc','ac','doc','docg',
 'pays','oc','cuvee','domaine','domaines','chateau','ch','maison','mas','clos','de','du','des','la','le','les',
 'et','d','l','aux','sur','en','vignobles','cave','caves','sainte','saint','st','sas','sarl','scea','sa','earl',
 'gaec','eurl','sc','monsieur','madame','mr','mme','m','societe','fils','freres','pere'}
MILL={str(y) for y in range(1990,2031)}
COUL={'rouge':'RGE','rge':'RGE','blanc':'BLC','blc':'BLC','blanche':'BLC','rose':'ROS','ros':'ROS'}
def _toks(s): return [t for t in _norm(s).split() if len(t)>2 and t not in STOP and t not in MILL]

_C={}
def _load(cat,four):
    k=(cat,four)
    if k in _C: return _C[k]
    df=pd.read_csv(cat,dtype=str).fillna('')
    df['nf']=df['nf'].str.strip(); df['code']=df['code'].str.strip()
    df['cuvee']=(df['des1']+' '+df['compl']+' '+df['app']).map(_norm)
    df['coul']=df['coul'].str.upper().str.strip()
    df['des1s']=df['des1']
    fr=pd.read_csv(four,dtype=str).fillna('')
    fr['nf']=fr['nf'].str.strip()
    idf={r['nf']:set(_toks(r['nom'])) for _,r in fr.iterrows()}
    dfreq=Counter()
    for h in df['cuvee']:
        for t in set(h.split()):
            if len(t)>2 and t not in STOP and t not in MILL: dfreq[t]+=1
    _C[k]=(df,idf,dfreq); return _C[k]

def resolve(domaine, article, catalogue_path='catalogue_vinistoria.csv', fournisseurs_path='fournisseurs_vinistoria.csv'):
    """-> (nf, code, note). code='' si non attribuable avec certitude."""
    if not (os.path.exists(catalogue_path) and os.path.exists(fournisseurs_path)): return '','',''
    df,idf,dfreq=_load(catalogue_path,fournisseurs_path)
    at=set(_toks(article)); dnorm=_norm(domaine); dt=set(_toks(domaine))
    coul=next((v for k,v in COUL.items() if re.search(r'\b'+k+r'\b',_norm(article))),'')
    mm=re.search(r'\b(20\d{2})\b',str(article)); mill=mm.group(1) if mm else ''

    # 1) code à 8 chiffres écrit en clair -> autoritaire
    m=re.search(r'\b(\d{8}[A-Z]{0,3})\b',str(article))
    if m:
        h=df[df['code']==m.group(1)]
        if len(h): return h.iloc[0]['nf'],h.iloc[0]['code'],'code lu dans le libellé'

    distinct={t for t in at if 0<dfreq.get(t,0)<=40}
    def best_in_nf(nf):
        sub=df[df['nf']==nf].copy()
        if not len(sub): return ('none',None)
        sub['sc']=sub['cuvee'].map(lambda h:len(at & set(h.split())))
        c=sub[sub['sc']>=1]
        if coul and len(c[c['coul']==coul]): c=c[c['coul']==coul]
        # si la cuvée a un mot distinctif, le candidat DOIT le contenir
        if distinct:
            c=c[c['cuvee'].map(lambda h:bool(distinct & set(h.split())))]
        if not len(c): return ('nocuvee',None)
        c=c.sort_values('sc',ascending=False); c=c[c['sc']==c['sc'].max()]
        if mill:
            ex=c[c['mill']==mill]
            if len(ex)==1: return ('code',ex.iloc[0]['code'])
        u=c.drop_duplicates('code')
        if len(u)==1: return ('code',u.iloc[0]['code'])
        lst='; '.join(f"{r['code']} = {r['des1s']} {r['mill']} {r['coul']}".strip() for _,r in u.head(8).iterrows())
        return ('many',lst)

    # 2) fournisseur : alias, puis base fournisseurs (nom complet + raison sociale)
    nf=''
    for k,v in ALIAS_FOURN.items():
        if k in dnorm: nf=v; break
    if not nf and dt:
        best,bs=None,0
        for n,bag in idf.items():
            s=len(dt & bag)
            if s>bs: bs,best=s,n
        if bs>=1: nf=best
    # indice alternatif : où se trouve la cuvée dans TOUT le catalogue (token distinctif)
    alt=''
    if distinct:
        pool=df.copy()
        pool['ds']=pool['cuvee'].map(lambda h:len(distinct & set(h.split())))
        pool=pool[pool['ds']>=1]
        if coul and len(pool[pool['coul']==coul]): pool=pool[pool['coul']==coul]
        if len(pool):
            pool=pool[pool['ds']==pool['ds'].max()]
            if mill and len(pool[pool['mill']==mill]): pool=pool[pool['mill']==mill]
            au=pool.drop_duplicates('code').head(4)
            alt='; '.join(f"{r['code']} (fourn {r['nf']}, {r['des1s']} {r['mill']} {r['coul']})".strip() for _,r in au.iterrows())

    if not nf:
        if alt: return '','',f"fournisseur non identifié pour « {domaine} ». Article similaire au catalogue : {alt} (à vérifier)"
        return '','',f"pas assez d'infos : domaine « {domaine} » introuvable en base fournisseurs"

    kind,val=best_in_nf(nf)
    if kind=='code': return nf,val,''
    if kind=='many': return nf,'',f"millésime à préciser — candidats fourn {nf} : {val}"
    # cuvée pas chez ce fournisseur
    if alt: return nf,'',f"fournisseur {nf} identifié, mais « {article} » absent de son catalogue. Ailleurs : {alt} (à vérifier)"
    return nf,'',f"fournisseur {nf} identifié, « {article} » introuvable au catalogue (à créer / préciser)"
