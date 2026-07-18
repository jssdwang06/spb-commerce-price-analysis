# -*- coding: utf-8 -*-
from __future__ import annotations
import re, warnings, json
from pathlib import Path
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib import font_manager
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.patches import FancyBboxPatch
from sklearn.compose import ColumnTransformer
from sklearn.ensemble import RandomForestRegressor
from sklearn.impute import SimpleImputer
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder
from sklearn.model_selection import cross_val_score
warnings.filterwarnings('ignore')
ROOT = Path(__file__).resolve().parent.parent  # project root
OUT = ROOT / 'figure'
OUT.mkdir(exist_ok=True)
CACHE = OUT / '_cache'
CACHE.mkdir(exist_ok=True)
PALETTE = {'navy':'#0B3C5D','teal':'#1F8A70','gold':'#F2A104','slate':'#4A5568','soft':'#E8F1F5','grid':'#D0D7DE','buy':'#0B3C5D','rent':'#1F8A70'}
for fp in [Path('C:/Windows/Fonts/msyh.ttc'), Path('C:/Windows/Fonts/simhei.ttf'), Path('C:/Windows/Fonts/arial.ttf')]:
    if fp.exists():
        font_manager.fontManager.addfont(str(fp))
        plt.rcParams['font.family'] = 'Microsoft YaHei' if fp.suffix.lower()=='.ttc' else ('SimHei' if 'simhei' in fp.name.lower() else 'Arial')
        break
plt.rcParams.update({'axes.unicode_minus':False,'figure.dpi':120,'savefig.dpi':300,'axes.spines.top':False,'axes.spines.right':False,'axes.edgecolor':'#8A95A1','axes.labelcolor':'#0B3C5D','text.color':'#0B3C5D','xtick.color':'#4A5568','ytick.color':'#4A5568','grid.color':'#D0D7DE','grid.alpha':0.75})
FEATURE_LABELS = {'area':'面积','metro_min':'到地铁时间','metro_km':'到地铁距离','has_metro':'有地铁信息','entrances':'入口数量','year':'建成年份','floor':'楼层','floors':'总层数','is_first_floor':'是否一层','center_km':'到市中心距离','ceiling':'层高','power_kw':'电力功率','type':'物业类型','layout':'布局','repair':'装修','location':'位置类型','walls':'墙体结构','parking':'停车','line':'临街线','hood':'有无排烟','tenant':'是否有租客','heating':'供暖'}

def num_from_text(value):
    if pd.isna(value): return np.nan
    text = str(value).replace('\xa0',' ').replace(' ','').replace(',','.')
    m = re.search(r'-?\d+(?:\.\d+)?', text)
    return float(m.group()) if m else np.nan

def nearest_distance_m_buy(value):
    if pd.isna(value): return np.nan
    ds=[]
    for part in str(value).split(';'):
        nums=re.findall(r'\d+(?:[.,]\d+)?', part)
        if nums: ds.append(float(nums[-1].replace(',','.')))
    return min(ds) if ds else np.nan

def first_metro_km_rent(value):
    if pd.isna(value): return np.nan
    m=re.search(r'\((\d+(?:[,.]\d+)?)\s*км\)', str(value))
    if m: return float(m.group(1).replace(',','.'))
    m=re.search(r'(\d+(?:[,.]\d+)?)\s*км', str(value))
    return float(m.group(1).replace(',','.')) if m else np.nan

def first_metro_min_rent(value):
    if pd.isna(value): return np.nan
    m=re.search(r'(\d+)\s*мин', str(value))
    return float(m.group(1)) if m else np.nan

def floor_pair(value):
    nums=re.findall(r'-?\d+', str(value))
    return (float(nums[0]), float(nums[1])) if len(nums)>=2 else (np.nan, np.nan)

def clean_cat(s):
    return s.astype('string').fillna('нет данных').replace({'':'нет данных','nan':'нет данных'}).str.strip()

def winsor_mask(df, cols):
    mask=pd.Series(True,index=df.index)
    for c in cols:
        lo,hi=df[c].quantile([0.01,0.99]); mask &= df[c].between(lo,hi)
    return mask

def short_label(t,n=28):
    t=str(t); return t if len(t)<=n else t[:n-1]+'…'

def load_buy():
    raw=pd.read_excel(ROOT/'data'/'Купить.xlsx'); d=pd.DataFrame(index=raw.index)
    d['price']=raw['Цена'].map(num_from_text); d['price_m2']=raw['Цена за м²'].map(num_from_text)
    d['area']=raw['Общая площадь'].map(num_from_text); d['type']=clean_cat(raw['Тип'])
    d['metro_min']=pd.to_numeric(raw['Метро время, мин'], errors='coerce')
    d['metro_km']=raw['Метро расстояние'].map(nearest_distance_m_buy)/1000
    d['has_metro']=raw['Метро'].notna().astype(float)
    d['entrances']=pd.to_numeric(raw['Количество входов'], errors='coerce')
    d['year']=pd.to_numeric(raw['Год постройки'], errors='coerce')
    fl=raw['Этаж / Этажность'].map(floor_pair); d['floor']=[x[0] for x in fl]; d['floors']=[x[1] for x in fl]
    d['is_first_floor']=(d['floor']==1).astype(float); d.loc[d['floor'].isna(),'is_first_floor']=np.nan
    d['layout']=clean_cat(raw['Планировка']); d['repair']=clean_cat(raw['Ремонт']); d['location']=clean_cat(raw['Расположение'])
    d['walls']=clean_cat(raw['Стены']); d['parking']=clean_cat(raw['Парковка']); d['line']=clean_cat(raw['Линия'])
    d['hood']=clean_cat(raw['Наличие вытяжки']); d['tenant']=clean_cat(raw['Арендатор']); d['heating']=clean_cat(raw['Отопление'])
    return d.replace([np.inf,-np.inf],np.nan).dropna(subset=['price','price_m2','area'])

def load_rent():
    raw=pd.read_excel(ROOT/'data'/'Снять.xlsx'); d=pd.DataFrame(index=raw.index)
    d['price']=raw['Цена'].map(num_from_text); d['price_m2']=raw['Цена за м²'].map(num_from_text)
    d['area']=raw['Общая площадь'].map(num_from_text); d['type']=clean_cat(raw['Тип'])
    d['metro_min']=raw['Метро'].map(first_metro_min_rent); d['metro_km']=raw['Метро'].map(first_metro_km_rent)
    d['has_metro']=raw['Метро'].notna().astype(float)
    d['entrances']=pd.to_numeric(raw['Количество входов'], errors='coerce')
    d['year']=pd.to_numeric(raw['Год постройки'], errors='coerce')
    fl=raw['Этаж / Этажность'].map(floor_pair); d['floor']=[x[0] for x in fl]; d['floors']=[x[1] for x in fl]
    d['is_first_floor']=(d['floor']==1).astype(float); d.loc[d['floor'].isna(),'is_first_floor']=np.nan
    d['center_km']=raw['До центра'].map(num_from_text)
    d['layout']=clean_cat(raw['Планировка']); d['repair']=clean_cat(raw['Ремонт']); d['location']=clean_cat(raw['Расположение'])
    d['walls']=clean_cat(raw['Стены']); d['parking']=clean_cat(raw['Парковка']); d['hood']=clean_cat(raw['Наличие вытяжки'])
    d['ceiling']=raw['Высота потолков'].map(num_from_text); d['power_kw']=raw['Мощность электричества'].map(num_from_text)
    return d.replace([np.inf,-np.inf],np.nan).dropna(subset=['price','price_m2','area'])

def importance_fast(df, target, num_cols, cat_cols):
    use_num=[c for c in num_cols if c in df.columns and df[c].notna().sum()>=15]
    use_cat=[c for c in cat_cols if c in df.columns and df[c].nunique(dropna=True)>=2]
    x=df[use_num+use_cat].copy(); y=np.log1p(df[target].astype(float).values)
    pre=ColumnTransformer([
        ('num', SimpleImputer(strategy='median'), use_num),
        ('cat', Pipeline([('imp', SimpleImputer(strategy='constant', fill_value='нет данных')),
                          ('oh', OneHotEncoder(handle_unknown='ignore', min_frequency=5, sparse_output=False))]), use_cat),
    ])
    rf=RandomForestRegressor(n_estimators=120, min_samples_leaf=3, random_state=42, n_jobs=1)
    pipe=Pipeline([('pre', pre), ('rf', rf)])
    scores=cross_val_score(pipe, x, y, cv=3, scoring='r2', n_jobs=1)
    pipe.fit(x,y)
    # map one-hot importances back to original features by permutation once on train
    rng=np.random.default_rng(42); base=pipe.score(x,y); drops={}
    for col in use_num+use_cat:
        xs=x.copy(); vals=xs[col].to_numpy().copy(); rng.shuffle(vals); xs[col]=vals
        drops[col]=max(0.0, base-pipe.score(xs,y))
    imp=pd.Series(drops).sort_values(ascending=False)
    return imp, float(np.mean(scores))

def metro_group(minutes):
    g=pd.cut(minutes, bins=[-np.inf,5,10,15,30,np.inf], labels=['≤5 分钟','6–10 分钟','11–15 分钟','16–30 分钟','>30 分钟'])
    return g.astype('string').fillna('无数据')

def center_group(km):
    g=pd.cut(km, bins=[-np.inf,3,6,10,15,np.inf], labels=['≤3 km','3–6 km','6–10 km','10–15 km','>15 km'])
    return g.astype('string').fillna('无数据')

def area_group(area):
    return pd.cut(area, bins=[0,50,100,200,400,np.inf], labels=['≤50㎡','50–100㎡','100–200㎡','200–400㎡','>400㎡'])

def savefig(fig,name):
    fig.savefig(OUT/name, bbox_inches='tight', facecolor='white'); plt.close(fig); print('saved', name)

def step_data():
    buy=load_buy(); rent=load_rent()
    buy_p=buy.loc[winsor_mask(buy,['price','price_m2','area'])].copy()
    rent_p=rent.loc[winsor_mask(rent,['price','price_m2','area'])].copy()
    buy.to_pickle(CACHE/'buy.pkl'); rent.to_pickle(CACHE/'rent.pkl')
    buy_p.to_pickle(CACHE/'buy_p.pkl'); rent_p.to_pickle(CACHE/'rent_p.pkl')
    print(f'buy={len(buy)} rent={len(rent)} plot_buy={len(buy_p)} plot_rent={len(rent_p)}')

def step_model():
    buy=pd.read_pickle(CACHE/'buy.pkl'); rent=pd.read_pickle(CACHE/'rent.pkl')
    imp_buy,r2_buy=importance_fast(buy,'price_m2',
        ['area','metro_min','metro_km','has_metro','entrances','year','floor','floors','is_first_floor'],
        ['type','layout','repair','location','walls','parking','line','hood','tenant','heating'])
    imp_rent,r2_rent=importance_fast(rent,'price_m2',
        ['area','metro_min','metro_km','has_metro','entrances','year','floor','floors','is_first_floor','center_km','ceiling','power_kw'],
        ['type','layout','repair','location','walls','parking','hood'])
    imp_buy.to_pickle(CACHE/'imp_buy.pkl'); imp_rent.to_pickle(CACHE/'imp_rent.pkl')
    (CACHE/'metrics.json').write_text(json.dumps({'r2_buy':r2_buy,'r2_rent':r2_rent}, ensure_ascii=False, indent=2), encoding='utf-8')
    print('r2_buy', r2_buy, 'r2_rent', r2_rent)
    print(imp_buy.head(8)); print(imp_rent.head(8))

if __name__=='__main__':
    import sys
    cmd=sys.argv[1] if len(sys.argv)>1 else 'data'
    if cmd=='data': step_data()
    elif cmd=='model': step_model()
    else: raise SystemExit('unknown')
