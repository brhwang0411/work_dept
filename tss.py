'''
TSS(T-Significant Signaling) Market Timing Model by Eunseok Yang

- Version  : 1.4.0
- 업데이트 : 2020.10.07
- 시차조정 (2020.10.05 부터 적용되는 기준가 컷오프제 반영)


'''

# 필요 module & data import
import numpy as np
import pandas as pd
import os, sqlite3
from scipy.stats import t as tdis
from datetime import datetime

# 현재 절대경로 확인
dbpath = 'D:/flaskapps/rasite/db/gqsdb.sqlite3'

# 데이터베이스 읽기
erbtr = {
'80':pd.DataFrame, '65':pd.DataFrame, '50':pd.DataFrame, '35':pd.DataFrame, '20':pd.DataFrame, '10':pd.DataFrame
}

# # 최초 DB 생성
# init_NAV = 1000.0
# ticker_name = ['국내주식','미국주식','유럽주식','일본주식','중국주식','신흥국주식','원자재','글로벌리츠','하이일드채권','신흥국채권','선진국채권','국내채권','단기자금']

# conn = sqlite3.connect(/dbpath/)
# df_daily = pd.read_sql_query('SELECT * FROM "idx_daily"', conn, index_col='date')
# conn.close()

# for t in erbtr:
#         erbtr[t] = pd.DataFrame(data=None, index=df_daily.index, columns=['NAV', 'REB', 'SIG']+ticker_name)
#         erbtr[t][0:3] = 0
#         erbtr[t].iloc[:3, 0] = init_NAV
#         erbtr[t].REB[0] = 1
#         erbtr[t].dropna(inplace=True)

conn = sqlite3.connect(dbpath)
df_daily = pd.read_sql_query('SELECT * FROM "idx_daily"', conn, index_col='date')
erbtr['80'] = pd.read_sql_query('SELECT * FROM "erb80_tss" ORDER BY "date" DESC LIMIT 42', conn, index_col='date').sort_values('date').iloc[:-1]
erbtr['65'] = pd.read_sql_query('SELECT * FROM "erb65_tss" ORDER BY "date" DESC LIMIT 42', conn, index_col='date').sort_values('date').iloc[:-1]
erbtr['50'] = pd.read_sql_query('SELECT * FROM "erb50_tss" ORDER BY "date" DESC LIMIT 42', conn, index_col='date').sort_values('date').iloc[:-1]
erbtr['35'] = pd.read_sql_query('SELECT * FROM "erb35_tss" ORDER BY "date" DESC LIMIT 42', conn, index_col='date').sort_values('date').iloc[:-1]
erbtr['20'] = pd.read_sql_query('SELECT * FROM "erb20_tss" ORDER BY "date" DESC LIMIT 42', conn, index_col='date').sort_values('date').iloc[:-1]
erbtr['10'] = pd.read_sql_query('SELECT * FROM "erb10_tss" ORDER BY "date" DESC LIMIT 42', conn, index_col='date').sort_values('date').iloc[:-1]
conn.close()

# 운용개시일로부터 시계열 재정의
dt_start = datetime(2018,11,5).date().isoformat()   # 운용개시일
df_daily = df_daily.loc[dt_start:]

# 리밸런싱 신호 유형별 유의수준 임계치(Machine Learning 결과)
# threshold = {'80':0.24, '65':0.27, '50':0.45, '35':0.42, '20':0.41, '10':0.40}    # 2020.06.08 이전
threshold = {'80':0.22, '65':0.20, '50':0.20, '35':0.21, '20':0.21, '10':0.54}

# 특정 weight에 대한 과거 평균 롤링수익률 계산
def avg_rollret_weight(dt_when, w_array, ts_ret, obs_n):
    w_array = np.array(w_array)
    df_ret = ts_ret[ts_ret.index.get_loc(dt_when)-obs_n+1:ts_ret.index.get_loc(dt_when)+1]
    avgret = np.sum(w_array*df_ret).sum()/obs_n
    return avgret

# 특정 NAV(Net Asset Value)에 대한 과거 평균 롤링수익률 계산
def avg_rollret_nav(dt_when, ts_nav, obs_n):
    df_ts = ts_nav[ts_nav.index.get_loc(dt_when)-obs_n:ts_nav.index.get_loc(dt_when)+1]
    rollret = df_ts.pct_change(1)[1:]
    avgret = np.average(rollret)
    return avgret

# 특정 weight에 대한 과거 분산 계산
def rollvar_weight(dt_when, w_array, ts_ret, obs_n):
    w_array = np.array(w_array)
    df_ret = ts_ret[ts_ret.index.get_loc(dt_when)-obs_n+1:ts_ret.index.get_loc(dt_when)+1]
    rollvar = np.square(np.sum(w_array*df_ret, axis=1)).sum()/obs_n
    return rollvar

# 특정 NAV(Net Asset Value)에 대한 과거 분산 계산
def rollvar_nav(dt_when, ts_nav, obs_n):
    df_ts = ts_nav[ts_nav.index.get_loc(dt_when)-obs_n:ts_nav.index.get_loc(dt_when)+1]
    rollret = df_ts.pct_change(1)[1:]
    rollvar = np.square(rollret).sum()/obs_n    
    return rollvar

# 기준일에 적용되는 ERB 비중 호출
def read_erb(dt_when, erb_type):
    conn = sqlite3.connect(dbpath)
    erbdt_list = pd.read_sql_query('SELECT "date" FROM erb{0}_hist'.format(erb_type), conn, index_col='date').index.tolist()
    erbdt_list.sort(reverse=True)
    erb = pd.DataFrame
    for i in erbdt_list:
        if dt_when >= i:
            erb = pd.read_sql_query("""SELECT * FROM erb{0}_hist WHERE date='{1}'""".format(erb_type, i), conn, index_col='date')
            return erb.iloc[0].T
            break
    if erb.empty:
        print('tracking에 참고할 기준일 이전의 유효한 자산배분안이 없습니다.''\n'
            '기준일 직전 일자의 ERB 자산배분 솔루션을 먼저 도출하십시오.')
    conn.close()    

# TSS 모델 사용하는 t-statistics 계산
def tstat(ret_w, ret_n, var_w, var_n, obs_n):
    tstat = (ret_w - ret_n)/np.sqrt(var_w/obs_n + var_n/obs_n)
    return tstat

# TSS 모델 Market Timing 신호 산출
def signaling(tstat, var_w, var_n, erb_type, obs_n):
    model = -tdis.ppf(threshold[erb_type]/2,
                  (var_w/obs_n+var_n/obs_n)**2/(((var_w/obs_n)**2)/(obs_n-1)+((var_n/obs_n)**2)/(obs_n-1)))
    if np.abs(tstat) > model:
        return 1.0
    else:
        return 0.0

# rebalancing 여부 도출
def isrebal(dt_when, erb_type):
    when_ix = erbtr[erb_type].index.get_loc(dt_when)
    if dt_when != dt_start and when_ix <= 5:
        return 0
    elif when_ix < 41 and np.sum(erbtr[erb_type].REB[when_ix-5:when_ix]) == 0:
        return signaling(tstat, var_w, var_n, erb_type, obs_n)
    else:
        if np.sum(erbtr[erb_type].REB[when_ix-40:when_ix]) == 0:
            return 1
        elif np.sum(erbtr[erb_type].REB[when_ix-5:when_ix]) == 0:
            return signaling(tstat, var_w, var_n, erb_type, obs_n)
        else:
            return 0

# Daily NAV 계산
def calnav(dt_when, ts_ret, erb_type):
    when_ix = erbtr[erb_type].index.get_loc(dt_when)
    if erbtr[erb_type].REB[when_ix-3]:
        erbtr[erb_type].iloc[when_ix][3:] = erbtr[erb_type].NAV.iloc[when_ix-1]*read_erb(erbtr[erb_type].index[when_ix-3], erb_type)*(1+ts_ret.loc[dt_when])
    else:
        erbtr[erb_type].iloc[when_ix][3:] = erbtr[erb_type].iloc[when_ix-1][3:]*(1+ts_ret.loc[dt_when])
    erbtr[erb_type].loc[dt_when, 'NAV'] = np.sum(erbtr[erb_type].iloc[when_ix][3:]) # python3.12 2025-02-28 수정

# daily return 구하기
df_dret = df_daily.pct_change(1)

# 시장별 시차 반영(shift)한 daily return 구하기
df_dretsft = pd.DataFrame([
    df_dret.국내주식.shift(1),
    df_dret.미국주식.shift(2),
    df_dret.유럽주식.shift(2),
    df_dret.일본주식.shift(1),
    df_dret.중국주식.shift(1),
    df_dret.신흥국주식.shift(2),
    df_dret.원자재.shift(2),
    df_dret.글로벌리츠.shift(2),
    df_dret.하이일드채권.shift(2),
    df_dret.신흥국채권.shift(2),
    df_dret.선진국채권.shift(2),
    df_dret.국내채권.shift(1),
    df_dret.단기자금.shift(1)
    ]).T

# 기준가 컷오프 반영 (2020-10-05 기준가부터 적용)
df_dretsft.loc['2020-10-05':, '일본주식'] = df_dret.일본주식.shift(2)['2020-10-05':] #python3.12
df_dretsft.loc['2020-10-05':, '일본주식'] = 0.0 #python3.12

df_dretsft.loc['2020-10-05':, '중국주식'] = df_dret.중국주식.shift(2)['2020-10-05':] #python3.12
df_dretsft.loc['2020-10-05':, '중국주식'] = 0.0 #python3.12

# Daily Position 계산
dt_when = '' #python3.12
idxfill = {'80':[], '65':[], '50':[], '35':[], '20':[], '10':[]}
for t in erbtr:
    erbtr[t] = erbtr[t].reindex(df_daily.index[df_daily.index.get_loc(erbtr[t].index[0]):], fill_value='NA')
    idxfill[t] = erbtr[t][erbtr[t].NAV == 'NA'].index.tolist()
    for i in idxfill[t]:
        # NAV 계산
        dt_when = i
        ts_ret = df_dretsft
        erb_type = t
        calnav(dt_when, ts_ret, erb_type)
        
        # SIG 계산
        if erbtr[t].index.get_loc(i) >= 6:
            w_array = read_erb(dt_when, erb_type)
            obs_n = 4
            ret_w = avg_rollret_weight(dt_when, w_array, ts_ret, obs_n)
            ts_nav = erbtr[t].NAV
            ret_n = avg_rollret_nav(dt_when, ts_nav, obs_n)
            var_w = rollvar_weight(dt_when, w_array, ts_ret, obs_n)
            var_n = rollvar_nav(dt_when, ts_nav, obs_n)
            tstat = (ret_w - ret_n)/np.sqrt(var_w/obs_n + var_n/obs_n)
            erbtr[t].loc[i, 'SIG'] = signaling(tstat, var_w, var_n, erb_type, obs_n) #python3.12
        else:
            erbtr[t].loc[i, 'SIG'] = 0 #python3.12
        
        # REB 계산
        erbtr[t].loc[i, 'REB'] = isrebal(dt_when, erb_type) #python3.12
        if erbtr[t].loc[i, 'REB'] == 1: #python3.12
            print('ERB',t, '유형에서', i, '에 리밸런싱 시그널이 발생했습니다.')

# TSS 모니터링 결과
df_monitor = pd.DataFrame(data=None, index=['80', '65', '50', '35', '20', '10'], columns=['NAV', 'REB', 'SIG'])
for t in erbtr:
    df_monitor.loc[t, 'NAV'] = erbtr[t].loc[dt_when, 'NAV'] #python3.12
    df_monitor.loc[t, 'REB'] = erbtr[t].loc[dt_when, 'REB'] #python3.12
    df_monitor.loc[t, 'SIG'] = erbtr[t].loc[dt_when, 'SIG'] #python3.12
print(
    dt_when,'\n\n',
    df_monitor,'\n\n'
    )

# Tracking 결과값 저장하기
conn = sqlite3.connect(dbpath)
cur = conn.cursor()
for t in erbtr:
    for i in idxfill[t]:
        cur.execute("""DELETE FROM erb{}_tss WHERE date='{}'""".format(t, i))
    for j in idxfill[t]:
        temp = [j]+erbtr[t].iloc[erbtr[t].index.get_loc(j)].tolist()
        cur.execute("""INSERT INTO erb{}_tss VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""".format(t), temp)
conn.commit()
conn.close()
print('결과가 저장되었습니다.')