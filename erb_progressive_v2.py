#!/usr/bin/env python
# coding: utf-8

# In[10]:


import numpy as np
import pandas as pd
import sqlite3, copy, os, json, warnings
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.ticker as mticker
from matplotlib import font_manager
from matplotlib.gridspec import GridSpec
from scipy.optimize import minimize
from datetime import datetime
from tqdm.notebook import tqdm
from IPython.display import display, HTML
from scipy.stats import t as tdis

warnings.filterwarnings('ignore')
pd.set_option('display.float_format', '{:.4f}'.format)

import platform
if platform.system() == 'Windows':
    font_path = 'C:/Windows/Fonts/malgun.ttf'
    font_name = font_manager.FontProperties(fname=font_path).get_name()
    matplotlib.rc('font', family=font_name)
elif platform.system() == 'Darwin':
    matplotlib.rc('font', family='AppleGothic')
else:
    matplotlib.rc('font', family='NanumGothic')
matplotlib.rcParams['axes.unicode_minus'] = False

print('✅ 라이브러리 로드 완료')

# ================================================================
#  ⚙️  파라미터 설정 — 테스트마다 여기만 수정
# ================================================================

# DB 경로
dbpath           = '/shared/HQ_Robo/DATA/Common/gqsdb.sqlite3'
dbpath_bm_weight = '/shared/HQ_Robo/DATA/Common/gqsoffdb.sqlite3'

# CSV 저장 폴더
csv_dir = './trials/ExcessReturn'


# In[11]:


# 이번 trial 설명 메모
trial_note = '2회차 재테스트'

# 백테스트 날짜 범위
start_date = '2018-09-30'
end_date   = None

# TSS 신호 없어도 강제 저장 여부 (테스트용)
FORCE_SAVE = False

# 모델 파라미터
numasset    = 12
window      = 26
decayfactor = 0.9
assetfloor        = 0.01
assetcap          = 0.8  
alt_asset_penalty = 0.5 

REGIME_SR_THRESHOLD   = 0.0
REGIME_LAMBDA_MIN     = 0.0
REGIME_RECOVERY_WEEKS = 4 
use_regime_switch = True

erb_profile2_bull = {
    '80': [3.0, 0.0, 0.05, 0.9],  
    '65': [2.5, 0.0, 0.05, 0.75],
    '50': [2.0, 0.0, 0.05, 0.56],
    '35': [1.0, 0.0, 0.05, 0.52],
    '20': [1.0, 0.0, 0.15, 0.22],
    '10': [0.0, 'MDD', 0.15, 0.37],
}
erb_profile2_bear = {
    "80": [1.0, 0.5, 0.5, 0.6], 
    "65": [1.0, 0.5, 0.5, 0.5], 
    "50": [0.5, 1.0, 0.8, 0.4], 
    "35": [0.5, 1.0, 0.8, 0.35], 
    "20": [0.0, 1.0, 1.0, 0.15], 
    "10": [0.0, "MDD", 1.0, 0.2]
}

# 현재 활성 프로파일 (regime switch 꺼져 있을 때 사용)
erb_profile2 = erb_profile2_bull

PROFILES = list(erb_profile2.keys())


# In[12]:


# ── Trial 번호 자동 부여 ──────────────────────────────────────────
os.makedirs(csv_dir, exist_ok=True)
trial_index_path = os.path.join(csv_dir, 'trial_index.csv')

if os.path.exists(trial_index_path):
    _idx_df  = pd.read_csv(trial_index_path)
    TRIAL_ID = int(_idx_df['trial_id'].max()) + 1
else:
    TRIAL_ID = 1

TRIAL_PREFIX = os.path.join(csv_dir, f'trial_{TRIAL_ID:03d}')

# ── 메타 출력 ────────────────────────────────────────────────────
print(f'{"="*55}')
print(f'  Trial {TRIAL_ID:03d}  |  {trial_note}')
print(f'{"="*55}')
print(f'  저장 경로   : {TRIAL_PREFIX}_*.csv')
print(f'  기간        : {start_date or "(DB 최초)"} ~ {end_date or "(DB 최신)"}')
print()
print(f'  [모델 파라미터]')
print(f'  window={window}주  decayfactor={decayfactor}  '
      f'cap={assetcap}  floor={assetfloor}  alt_penalty={alt_asset_penalty}')
print()
print(f'  [Regime Switching] {"ON" if use_regime_switch else "OFF"}')
if use_regime_switch:
    print(f'  SR 임계치={REGIME_SR_THRESHOLD}  '
          f'Lambda 임계치={REGIME_LAMBDA_MIN}  '
          f'복귀 유예={REGIME_RECOVERY_WEEKS}주')
    print()
    print(f'  [Bull 프로파일]  α    β    γ    δ')
    for p, v in erb_profile2_bull.items():
        print(f'    ERB {p}       '
              f'{str(v[0]):<5}{str(v[1]):<5}{str(v[2]):<5}{v[3]:.2f}')
    print()
    print(f'  [Bear 프로파일]  α    β    γ    δ')
    for p, v in erb_profile2_bear.items():
        print(f'    ERB {p}       '
              f'{str(v[0]):<5}{str(v[1]):<5}{str(v[2]):<5}{v[3]:.2f}')
else:
    print(f'  [단일 프로파일]  α    β    γ    δ')
    for p, v in erb_profile2.items():
        print(f'    ERB {p}       '
              f'{str(v[0]):<5}{str(v[1]):<5}{str(v[2]):<5}{v[3]:.2f}')
print(f'{"="*55}')


# In[13]:


ASSET_NAMES      = ['국내주식','미국주식','유럽주식','일본주식',
                    '중국주식','신흥국주식','원자재','글로벌리츠',
                    '하이일드채권','신흥국채권','선진국채권','국내채권']
ASSET_NAMES_CASH = ASSET_NAMES + ['단기자금']

# ── DB 유틸 ─────────────────────────────────────────────────────

def read_table(table_name):
    with sqlite3.connect(dbpath) as conn:
        return pd.read_sql_query(
            f'SELECT * FROM "{table_name}" ORDER BY "date"',
            conn, index_col='date'
        )

def get_bm_weight(dt_when):
    with sqlite3.connect(dbpath_bm_weight) as conn:
        dt_list = pd.read_sql_query(
            'SELECT "date" FROM bm_weight ORDER BY date',
            conn, index_col='date'
        ).index.tolist()
    for i in sorted(dt_list, reverse=True):
        if dt_when >= i:
            with sqlite3.connect(dbpath_bm_weight) as conn:
                row = pd.read_sql_query(
                    f"SELECT * FROM bm_weight WHERE date='{i}'",
                    conn, index_col='date'
                ).values.ravel().tolist()
            ms_norm = ([row[-1]] + [row[10]-row[3]] + row[5:8]
                       + [row[8]-row[-1]] + [row[9]] + [row[3]]
                       + [row[11]+row[13]/4] + [row[12]+row[13]/4])
            ms_norm += [(100-sum(ms_norm[8:10]))/2]*2
            ms_roff  = (ms_norm[:6] + [0.0] + [ms_norm[7]]
                        + [0.0, 0.0]
                        + [50+sum(row[11:13])] + [50-sum(row[11:13])])
            return ms_norm, ms_roff
    return None, None

# ── 입력 생성 ───────────────────────────────────────────────────

def make_input(df_ts, dt_when):
    idx = np.where(df_ts.index == dt_when)[0][0]
    return df_ts[idx - window: idx + 1]

def make_array_decay():
    arr = np.ones(1) * (1 - decayfactor)
    while len(arr) < window:
        arr = np.append(arr, arr[-1] * decayfactor)
    return np.sort(arr)

# ── 포트폴리오 리스크 ────────────────────────────────────────────

def portrisk(weight, covmat):
    w = np.array(weight)
    return float((w.T @ covmat @ w) ** 0.5)

def riskcontribution(weight, covmat):
    w  = np.array(weight)
    pr = portrisk(w, covmat)
    rc = w * ((1/pr) * (covmat @ w))
    return rc / rc.sum()

# ── 최적화 ──────────────────────────────────────────────────────

def objective_rb(weight, args):
    return float(np.sum(np.square(
        riskcontribution(np.array(weight), args[0]) - args[1]
    )))

# def rbweight(covmat, rc_target, x0):
#     res = minimize(
#         fun=objective_rb, args=[covmat, rc_target], x0=x0,
#         bounds=((0.,1.),)*numasset, method='SLSQP',
#         constraints=(
#             {'type':'eq',   'fun': lambda x: np.sum(x)-1.},
#             {'type':'ineq', 'fun': lambda x: x},
#         ),
#         options={'ftol':1e-12,'maxiter':500,'disp':False}
#     )
#     return res.x

def _regularize_cov(covmat, alpha=1e-5):
    """
    공분산 행렬 정규화 — ill-conditioned 상태 완화.
    대각선에 작은 값을 더해서 행렬을 양정치(positive definite)로 만듦.
    alpha가 클수록 강하게 정규화 (변동성 과소추정 방지).
    """
    cov = np.array(covmat, dtype=float)
    # 조건수 확인
    try:
        cond = np.linalg.cond(cov)
    except Exception:
        cond = np.inf

    if cond > 1e10:   # ill-conditioned 기준
        # 대각 원소 평균의 alpha 배를 더함
        diag_mean = np.diag(cov).mean()
        cov = cov + np.eye(len(cov)) * diag_mean * alpha
    return cov


def rbweight(covmat, rc_target, x0, tol=1e-11):
    """
    SLSQP 기반 리스크 버짓팅.
    - 공분산 정규화로 ill-conditioned 대응
    - 다중 초기값으로 수렴 실패 보완
    - 단계적 허용 오차 완화로 최후 폴백
    """
    cov = _regularize_cov(covmat)
    rc  = np.array(rc_target, dtype=float)
    n   = len(rc)

    # ── 초기값 후보군 ────────────────────────────────────────────
    x0_arr = np.array(x0, dtype=float).ravel()
    x0_arr = np.maximum(x0_arr, 0.)
    if x0_arr.sum() > 0:
        x0_arr /= x0_arr.sum()

    # IVP: 역변동성 비중
    diag_sd  = np.sqrt(np.diag(cov))
    ivp_w    = (1. / diag_sd) / (1. / diag_sd).sum()

    # rc_target 비중: 목표 리스크 예산 자체를 초기값으로
    rc_w     = rc / rc.sum()

    # 균등 비중
    eq_w     = np.full(n, 1. / n)

    x0_candidates = [x0_arr, ivp_w, rc_w, eq_w]

    # ── SLSQP 공통 설정 ──────────────────────────────────────────
    constraints = (
        {'type': 'eq',   'fun': lambda x: x.sum() - 1.},
        {'type': 'ineq', 'fun': lambda x: x},
    )
    bounds = ((0., 1.),) * n

    best_w, best_obj = None, np.inf

    # ── 1차: 표준 정밀도로 다중 초기값 시도 ─────────────────────
    for _x0 in x0_candidates:
        try:
            res = minimize(
                fun=objective_rb, args=[cov, rc], x0=_x0,
                bounds=bounds, method='SLSQP',
                constraints=constraints,
                options={'ftol': 1e-14, 'maxiter': 2000, 'disp': False}
            )
            obj = objective_rb(res.x, (cov, rc))
            if obj < best_obj:
                best_obj, best_w = obj, res.x.copy()
            if best_obj < tol:
                return best_w
        except Exception:
            continue

    if best_obj < 1e-6:
        return best_w

    # ── 2차: 정규화 강도 높여서 재시도 ──────────────────────────
    cov_strong = _regularize_cov(covmat, alpha=1e-3)
    for _x0 in [best_w if best_w is not None else eq_w, ivp_w]:
        try:
            res = minimize(
                fun=objective_rb, args=[cov_strong, rc], x0=_x0,
                bounds=bounds, method='SLSQP',
                constraints=constraints,
                options={'ftol': 1e-12, 'maxiter': 3000, 'disp': False}
            )
            obj = objective_rb(res.x, (cov, rc))  # 원래 cov로 검증
            if obj < best_obj:
                best_obj, best_w = obj, res.x.copy()
            if best_obj < tol:
                return best_w
        except Exception:
            continue

    if best_obj < 1e-4:
        return best_w

    # ── 3차: rc_target 평탄화 후 재시도 (ERC에 가깝게) ──────────
    # 목표 리스크 예산을 균등과 원래 값의 중간으로 완화
    rc_smooth = 0.5 * rc + 0.5 * eq_w
    rc_smooth /= rc_smooth.sum()
    for _x0 in [best_w if best_w is not None else eq_w, ivp_w]:
        try:
            res = minimize(
                fun=objective_rb, args=[cov, rc_smooth], x0=_x0,
                bounds=bounds, method='SLSQP',
                constraints=constraints,
                options={'ftol': 1e-12, 'maxiter': 3000, 'disp': False}
            )
            obj = objective_rb(res.x, (cov, rc))  # 원래 rc로 검증
            if obj < best_obj:
                best_obj, best_w = obj, res.x.copy()
            if best_obj < tol:
                return best_w
        except Exception:
            continue

    # ── 최후 폴백: IVP 반환 ──────────────────────────────────────
    if best_w is None or best_obj > 0.1:
        return ivp_w

    return best_w

def rbweight_alt(covmat, rc_target, x0):
    res = minimize(
        fun=lambda w: np.sum(np.square(w)), x0=x0,
        bounds=((0.,1.),)*numasset, method='SLSQP',
        constraints=(
            {'type':'eq',   'fun': lambda x: np.sum(x)-1.},
            {'type':'ineq', 'fun': lambda x: x},
            {'type':'eq',   'fun': lambda x: np.sum(np.square(
                riskcontribution(x, covmat) - rc_target))},
        ),
        options={'ftol':1e-12,'maxiter':500,'disp':False}
    )
    return res.x

def get_regime_profile(erc_lambda, erc_sr, regime_history):
    """
    현재 시장 국면 판단 → 적용할 erb_profile2 반환
    
    regime_history: 최근 N주의 regime 기록 (1=bull, 0=bear)
    """
    if not use_regime_switch:
        return erb_profile2, 'bull'

    # 위기 판단
    is_bear = (erc_lambda < REGIME_LAMBDA_MIN or
               erc_sr     < REGIME_SR_THRESHOLD)

    if is_bear:
        return erb_profile2_bear, 'bear'

    # 위기 해제 후 유예 기간 — 최근 N주 중 bear가 있으면 대기
    recent_bear_count = sum(1 for r in regime_history[-REGIME_RECOVERY_WEEKS:]
                            if r == 'bear')
    if recent_bear_count > 0:
        return erb_profile2_bear, 'recovery'   # 유예 중

    return erb_profile2_bull, 'bull'

# ── ERB 4단계 폴백 검증 ─────────────────────────────────────────

def validate_erb(profile, erb_w, cov, rc, erc_w, ivp_w, prev_cache):
    tol = 1e-11
    def ok(w): return objective_rb(w, (cov, rc)) < tol
    if ok(erb_w):                                     return erb_w, 'OK-init'
    w2 = rbweight_alt(cov, rc, erc_w)
    if ok(w2):                                        return w2,    'OK-alt'
    w3 = rbweight(cov, rc, ivp_w)
    if ok(w3):                                        return w3,    'OK-ivp'
    prev = prev_cache.get(profile)
    if prev is not None:
        w4 = rbweight(cov, rc, prev)
        if ok(w4):                                    return w4,    'OK-prev'
    return erb_w,                                            'FAIL'

# ── Cap/Floor 후처리 ─────────────────────────────────────────────

def apply_cap_floor(erb_dict):
    bef = pd.DataFrame(erb_dict, index=ASSET_NAMES)
    aft = pd.DataFrame(index=ASSET_NAMES, columns=PROFILES, dtype=float)

    simple = ['국내주식','유럽주식','일본주식','중국주식','원자재',
              '글로벌리츠','하이일드채권','신흥국채권','국내채권']
    for a in simple:
        aft.loc[a] = np.where(bef.loc[a] < assetfloor, 0.,
                              np.minimum(bef.loc[a], assetcap))

    em_pool     = bef.loc[['국내주식','중국주식','원자재','신흥국주식']].sum()
    em_alloc    = aft.loc[['국내주식','원자재','신흥국주식']].sum()
    em_residual = em_pool - em_alloc
    aft.loc['신흥국주식'] = np.where(em_residual < assetfloor, 0.,
                                    np.minimum(em_residual, assetcap))

    eq_pool  = bef.iloc[:8].sum()
    eq_alloc = pd.concat([aft.iloc[:1], aft.iloc[2:8]]).sum()
    aft.loc['미국주식'] = np.minimum(eq_pool - eq_alloc, assetcap)

    bond_pool  = bef.iloc[8:11].sum()
    bond_alloc = aft.iloc[8:10].sum()
    aft.loc['선진국채권'] = np.where(
        bef.loc['선진국채권'] < assetfloor, 0.,
        np.minimum(bond_pool - bond_alloc, assetcap)
    )

    cash = (1. - aft.sum()).clip(lower=0.)
    aft  = pd.concat([aft, pd.DataFrame(cash).T.rename(index={0:'단기자금'})])
    return aft

# ── 단조성 보정 ──────────────────────────────────────────────────

def monotone_fix(erb_mod, adj_cov):
    erb_sd    = {p: portrisk(erb_mod[p].values[:-1], adj_cov) for p in PROFILES}
    erb_final = {}
    for i, p in enumerate(PROFILES):
        if i == 0:
            erb_final[p] = erb_mod[p].values.tolist()
        else:
            prev_p = PROFILES[i-1]
            erb_final[p] = (erb_final[prev_p]
                            if erb_sd[p] >= erb_sd[prev_p]
                            else erb_mod[p].values.tolist())
    return erb_final, erb_sd

# ── CSV append 유틸 ──────────────────────────────────────────────

def append_or_create(path, df):
    if os.path.exists(path):
        df.to_csv(path, mode='a', header=False, index=False)
    else:
        df.to_csv(path, index=False)

print('✅ 함수 정의 완료')


# In[14]:


# ── 데이터 로드 ─────────────────────────────────────────────────
df_weekly = read_table('idx_weekly')
print(f'주간 DB 로드: {len(df_weekly)}행  {df_weekly.index[0]} ~ {df_weekly.index[-1]}')

rf_col_candidates = [c for c in df_weekly.columns if '단기자금' in c]
rf_col = rf_col_candidates[0] if rf_col_candidates else None

# ── 백테스트 날짜 범위 ───────────────────────────────────────────
all_dates   = df_weekly.index.tolist()
first_valid = all_dates[window]
_sd = start_date if start_date else first_valid
_ed = end_date   if end_date   else all_dates[-1]
_sd = max(_sd, first_valid)
bt_dates = [d for d in all_dates if _sd <= d <= _ed]
print(f'백테스트: {bt_dates[0]} ~ {bt_dates[-1]}  총 {len(bt_dates)}개 기준일')

# ── 루프 전 초기화 (루프 밖!) ────────────────────────────────────
buf_weights    = []
buf_risk       = []
buf_erc        = []
prev_erb_cache = {p: None for p in PROFILES}
prev_erc_cache = None
regime_history = []
array_decay    = make_array_decay()
run_start      = datetime.now()

# ── 메인 루프 ────────────────────────────────────────────────────
for dt_when in tqdm(bt_dates, desc=f'Trial {TRIAL_ID:03d}'):

    # ── Step 1. BM Weight ────────────────────────────────────────
    try:
        bm_raw = get_bm_weight(dt_when)
        if bm_raw[0] is None:
            print(f'  ⚠️ {dt_when} BM weight 없음')
            continue
        ms_norm = pd.Series(bm_raw[0], index=ASSET_NAMES)
        ms_roff = pd.Series(bm_raw[1], index=ASSET_NAMES)
    except Exception as e:
        print(f'  ⚠️ {dt_when} BM 실패: {e}')
        continue

    # ── Step 2. 리스크 지표 ──────────────────────────────────────
    try:
        df_input         = make_input(df_weekly, dt_when).iloc[:, :numasset].copy()
        df_input.columns = ASSET_NAMES
        df_input_ret     = df_input.pct_change(1)[1:]
        rf_rate          = df_weekly['기준금리'][dt_when] / 100

        hist_ret = df_input.iloc[-1] / df_input.iloc[0] - 1
        avg_ret  = df_input_ret.mean()
        mom_ret  = (df_input.iloc[-1] / df_input.iloc[-21]) ** (52 / 21) - 1
        hist_sd  = np.sqrt(52) * np.std(df_input_ret)
        ewma_sd  = np.sqrt(
            np.sum((df_input_ret ** 2) *
                   np.tile(array_decay, numasset).reshape(numasset, window).T,
                   axis=0) * 52
        )
        mdd = np.min((df_input / np.maximum.accumulate(df_input)) - 1)

        up_count = np.sum(np.ceil(np.maximum(df_input_ret - avg_ret, 0)))
        up_avg   = np.sum(np.ceil(np.maximum(df_input_ret - avg_ret, 0))
                          * df_input_ret) / up_count
        up_vol   = np.sqrt(
            (np.sum((np.ceil(np.maximum(df_input_ret - avg_ret, 0))
                     * df_input_ret - up_avg) ** 2)
             - (window - up_count) * up_avg ** 2) * window / up_count
        )
        down_count = np.sum(np.ceil(np.maximum(avg_ret - df_input_ret, 0)))
        down_avg   = np.sum(np.ceil(np.maximum(avg_ret - df_input_ret, 0))
                            * df_input_ret) / down_count
        down_vol   = np.sqrt(
            (np.sum((np.ceil(np.maximum(avg_ret - df_input_ret, 0))
                     * df_input_ret - down_avg) ** 2)
             - (window - down_count) * down_avg ** 2) * window / down_count
        )

        sd_adj_delta = ewma_sd - hist_sd
        sd_adj_delta[sd_adj_delta >= 0] = (
            sd_adj_delta[sd_adj_delta >= 0] * (down_vol / up_vol)
        )
        sd_adj_delta[sd_adj_delta < 0] = (
            sd_adj_delta[sd_adj_delta < 0] * (up_vol / down_vol)
        )
        sd_adj_min = (
            np.sqrt(52) * np.std(
                make_input(df_weekly, dt_when)[rf_col].pct_change(1)[1:]
            ) if rf_col else 0.001
        )
        adj_sd   = np.maximum(hist_sd + sd_adj_delta, sd_adj_min)
        hist_cor = df_input_ret.corr()
        adj_cov  = np.matrix(adj_sd).T * np.matrix(adj_sd) * hist_cor

    except Exception as e:
        print(f'  ⚠️ {dt_when} 리스크 계산 실패: {e}')
        continue

    # ── Step 3. ERC ──────────────────────────────────────────────
    try:
        ivp     = 1 / adj_sd / np.sum(1 / adj_sd)
        erc_tgt = np.repeat(1 / numasset, numasset)

        x0_candidates = [ivp, np.repeat(1 / numasset, numasset)]
        if prev_erc_cache is not None:
            x0_candidates.insert(1, prev_erc_cache)

        best_erc, best_obj = None, np.inf
        for x0 in x0_candidates:
            w   = rbweight(adj_cov, erc_tgt, x0)
            obj = objective_rb(w, (adj_cov, erc_tgt))
            if obj < best_obj:
                best_obj, best_erc = obj, w
            if obj < 1e-6:
                break

        erc = best_erc.tolist()

        # ── 수렴 불완전 시 전주 ERC 블렌딩 ──────────────────────
        if best_obj >= 1e-3:
            if prev_erc_cache is not None:
                blend = min((best_obj - 1e-3) / (2e-2 - 1e-3), 1.0)
                
                # prev_erc_cache가 list일 수도 numpy array일 수도 있음
                prev_list = (prev_erc_cache.tolist() 
                             if hasattr(prev_erc_cache, 'tolist')
                             else list(prev_erc_cache))
                
                erc_blended = [
                    (1 - blend) * c + blend * p
                    for c, p in zip(erc, prev_list)
                ]
                erc_sum = sum(erc_blended)
                erc = [w / erc_sum for w in erc_blended]
                print(f'  ℹ️  {dt_when} ERC 블렌딩 '
                      f'(obj={best_obj:.2e}, blend={blend:.2f}: '
                      f'현재{(1-blend)*100:.0f}% + 전주{blend*100:.0f}%)')
            else:
                print(f'  ⚠️ {dt_when} ERC 수렴 불완전 (obj={best_obj:.2e}), '
                      f'전주 없어 best 사용')
        # ─────────────────────────────────────────────────────────

        erc_risk   = portrisk(erc, adj_cov)
        erc_sr     = (np.sum(hist_ret * erc) - rf_rate) / erc_risk
        erc_lambda = (erc_sr / erc_risk if erc_sr >= 0
                      else (np.sum(mom_ret * erc) - rf_rate) / erc_risk ** 2)
        erc_lambda = float(np.clip(erc_lambda, -500.0, 500.0))  # 폭발 방지
        apc_werc   = ((hist_cor.sum() - 1) / (numasset - 1) * erc).sum()
        mu_hat     = np.asarray(erc_lambda * adj_cov @ erc).ravel()
        prev_erc_cache = erc

        # ── Regime 판단 (ERC 완료 직후) ──────────────────────────
        current_profile, current_regime = get_regime_profile(
            erc_lambda, float(erc_sr), regime_history
        )
        regime_history.append(current_regime)
        if len(regime_history) > REGIME_RECOVERY_WEEKS * 2:
            regime_history.pop(0)

    except Exception as e:
        print(f'  ⚠️ {dt_when} ERC 실패: {e}')
        continue

    # ── Step 4. ERB ──────────────────────────────────────────────
    try:
        mksize = ms_norm if erc_lambda > 0 else ms_roff
        wp = {p: pd.concat([
            mksize[:-4] / mksize[:-4].sum() * current_profile[p][3] * 100,
            mksize[-4:] / mksize[-4:].sum() * (1 - current_profile[p][3]) * 100
        ]) for p in PROFILES}
        adj_cor2 = {p: np.maximum(hist_cor, hist_cor * current_profile[p][2])
                    for p in PROFILES}
        adj_cov2 = {p: np.matrix(adj_sd).T * np.matrix(adj_sd) * adj_cor2[p]
                    for p in PROFILES}
        apc_  = {p: (adj_cor2[p].sum() - 1) / (numasset - 1) for p in PROFILES}
        wapc  = {p: (apc_[p] * wp[p]).sum() / 100             for p in PROFILES}
        sumpos = np.sum([r for r in mu_hat if r > 0])
        abgt   = pd.Series(np.ones(numasset), index=ASSET_NAMES)
        abgt['원자재':'글로벌리츠'] = alt_asset_penalty
        bgt = {p: (1 / adj_sd * abgt if sumpos == 0
                   else mu_hat ** current_profile[p][0]
                        / adj_sd ** current_profile[p][1] * abgt)
               for p in PROFILES if p != '10'}
        bgt['10'] = mu_hat ** current_profile['10'][0] / -mdd * abgt
        bgt = {p: bgt[p] * wp[p] ** (0.5 + wapc[p]) for p in PROFILES}
        bgt_adj = copy.deepcopy(bgt)
        for p in PROFILES:
            for c in bgt[p].index:
                bgt_adj[p][c] = (0. if bgt[p][c] <= 0.1
                                 else round(bgt[p][c], 6))
        rc_target = {}
        for p in PROFILES:
            s_raw = bgt[p].sum()
            s_adj = bgt_adj[p].sum()
            rc_target[p] = (bgt[p] / s_raw
                            if (s_adj == 0 or s_adj / s_raw < 0.5)
                            else bgt_adj[p] / bgt_adj[p].sum())
    
        erb = {p: rbweight(adj_cov2[p], rc_target[p], erc) for p in PROFILES}
        for p in PROFILES:
            erb[p], _ = validate_erb(
                p, erb[p], adj_cov2[p], rc_target[p],
                erc, ivp.values, prev_erb_cache
            )
        erb_mod               = apply_cap_floor({p: erb[p] for p in PROFILES})
        erb_final, erb_sd_map = monotone_fix(erb_mod, adj_cov)
        for p in PROFILES:
            prev_erb_cache[p] = np.array(erb_final[p][:-1])
    except Exception as e:
        print(f'  ⚠️ {dt_when} ERB 실패: {e}')
        continue

    # ── Step 5. 버퍼 적재 ────────────────────────────────────────
    for p in PROFILES:
        for asset, w in zip(ASSET_NAMES_CASH, erb_final[p]):
            buf_weights.append({
                'date':    dt_when,
                'profile': p,
                'asset':   asset,
                'weight':  round(w, 6),
            })

    for asset in ASSET_NAMES:
        ai = ASSET_NAMES.index(asset)
        buf_risk.append({
            'date':     dt_when,
            'asset':    asset,
            'hist_ret': round(float(hist_ret[asset]), 6),
            'mom_ret':  round(float(mom_ret[asset]),  6),
            'hist_sd':  round(float(hist_sd[asset]),  6),
            'ewma_sd':  round(float(ewma_sd[asset]),  6),
            'adj_sd':   round(float(adj_sd[asset]),   6),
            'mu_hat':   round(float(mu_hat[ai]),      6),
        })

    # regime은 erc_row에 한 번만 기록 (for 루프 밖)
    erc_row = {
        'date':       dt_when,
        'regime':     current_regime,
        'erc_lambda': round(float(erc_lambda), 6),
        'erc_sr':     round(float(erc_sr),     6),
        'erc_risk':   round(float(erc_risk),   6),
        'rf_rate':    round(float(rf_rate),     6),
        'apc_werc':   round(float(apc_werc),   6),
    }
    for asset, w in zip(ASSET_NAMES, erc):
        erc_row[f'erc_{asset}'] = round(float(w), 6)
    buf_erc.append(erc_row)

run_end = datetime.now()

# ── CSV 저장 ─────────────────────────────────────────────────────
df_weights = pd.DataFrame(buf_weights)
df_risk    = pd.DataFrame(buf_risk)
df_erc     = pd.DataFrame(buf_erc)

df_weights.to_csv(f'{TRIAL_PREFIX}_weights.csv', index=False)
df_risk.to_csv(   f'{TRIAL_PREFIX}_risk.csv',    index=False)
df_erc.to_csv(    f'{TRIAL_PREFIX}_erc.csv',     index=False)

# ── 메타 저장 ────────────────────────────────────────────────────
meta = {
    'trial_id':              TRIAL_ID,
    'note':                  trial_note,
    'run_at':                run_start.strftime('%Y-%m-%d %H:%M:%S'),
    'elapsed_sec':           round((run_end - run_start).total_seconds(), 1),
    'bt_start':              bt_dates[0],
    'bt_end':                bt_dates[-1],
    'n_dates':               len(bt_dates),
    'window':                window,
    'decayfactor':           decayfactor,
    'assetcap':              assetcap,
    'assetfloor':            assetfloor,
    'alt_asset_penalty':     alt_asset_penalty,
    'use_regime_switch':     use_regime_switch,
    'regime_sr_threshold':   REGIME_SR_THRESHOLD,
    'regime_lambda_min':     REGIME_LAMBDA_MIN,
    'regime_recovery_weeks': REGIME_RECOVERY_WEEKS,
    'erb_profile2_bull':     json.dumps(erb_profile2_bull),
    'erb_profile2_bear':     json.dumps(erb_profile2_bear),
}
pd.DataFrame([meta]).to_csv(f'{TRIAL_PREFIX}_meta.csv', index=False)

# ── trial_index.csv 업데이트 ─────────────────────────────────────
meta_slim = {k: meta[k] for k in [
    'trial_id', 'note', 'run_at', 'elapsed_sec',
    'bt_start', 'bt_end', 'n_dates',
    'window', 'decayfactor', 'assetcap', 'assetfloor',
    'alt_asset_penalty', 'use_regime_switch',
    'regime_sr_threshold', 'regime_lambda_min', 'regime_recovery_weeks',
]}
append_or_create(trial_index_path, pd.DataFrame([meta_slim]))

print(f'\n✅ Trial {TRIAL_ID} 완료  (소요: {run_end - run_start})')
print(f'   weights : {TRIAL_PREFIX}_weights.csv  ({len(df_weights):,}행)')
print(f'   risk    : {TRIAL_PREFIX}_risk.csv     ({len(df_risk):,}행)')
print(f'   erc     : {TRIAL_PREFIX}_erc.csv      ({len(df_erc):,}행)')


# In[21]:


from scipy.stats import t as tdis

# ── 일간 데이터 로드 ─────────────────────────────────────────────
with sqlite3.connect(dbpath) as conn:
    df_daily = pd.read_sql_query(
        'SELECT * FROM "idx_daily"', conn, index_col='date'
    )
print(f'일간 DB 로드: {len(df_daily)}행  {df_daily.index[0]} ~ {df_daily.index[-1]}')

TICKER_COLS = [
    '국내주식', '미국주식', '유럽주식', '일본주식', '중국주식',
    '신흥국주식', '원자재', '글로벌리츠', '하이일드채권',
    '신흥국채권', '선진국채권', '국내채권', '단기자금'
]

# ── TSS 파라미터 ─────────────────────────────────────────────────
INIT_NAV  = 1000.0
OBS_N     = 4
COOL_DOWN = 5
FORCE_REB = 40
LAG_REB   = 3
nav_start_date = "2018-11-05"  # 예: '2020-01-05'

TSS_THRESHOLD = {
    '80': 0.22, '65': 0.20, '50': 0.20,
    '35': 0.21, '20': 0.21, '10': 0.54,
}

# ── 일간 수익률 시차 반영 ────────────────────────────────────────
df_dret = df_daily[TICKER_COLS].pct_change(1)
SHIFT_MAP = {
    '국내주식':1, '미국주식':2, '유럽주식':2, '일본주식':1,
    '중국주식':1, '신흥국주식':2, '원자재':2, '글로벌리츠':2,
    '하이일드채권':2, '신흥국채권':2, '선진국채권':2,
    '국내채권':1, '단기자금':1,
}
df_dretsft = pd.DataFrame(
    {col: df_dret[col].shift(s) for col, s in SHIFT_MAP.items()}
).fillna(0.0)

# 기준가 컷오프 보정
for col in ['일본주식', '중국주식']:
    df_dretsft.loc['2020-10-05':, col] = df_dret[col].shift(2).loc['2020-10-05':]
    df_dretsft.loc['2020-10-05', col]  = 0.0

# ── 주간 ERB 비중 → 날짜별 딕셔너리 ─────────────────────────────
# buf_weights(메모리) 또는 CSV에서 로드
if len(buf_weights) > 0:
    _wdf = pd.DataFrame(buf_weights)
else:
    _wdf = pd.read_csv(f'{TRIAL_PREFIX}_weights.csv')

weekly_erb = {}   # {profile: {date_str: np.array(13,)}}
for p in PROFILES:
    _sub = _wdf[_wdf['profile'] == p].copy()
    _sub = _sub.pivot(index='date', columns='asset', values='weight')
    _sub = _sub.reindex(columns=TICKER_COLS).fillna(0.0)
    weekly_erb[p] = _sub   # DataFrame, index=날짜, cols=TICKER_COLS 순서 보장

print(f'ERB 비중 로드: {_wdf["date"].min()} ~ {_wdf["date"].max()}')

def _get_erb_w(p, date):
    """date 이전 가장 최근 주간 ERB 비중 → np.array(13,)"""
    wdf   = weekly_erb[p]
    valid = wdf.index[wdf.index <= date]
    if len(valid) == 0:
        return None
    return wdf.loc[valid[-1]].values.astype(float)

# ── 일간 날짜 범위 ───────────────────────────────────────────────
# ERB 비중이 존재하는 첫 날짜 확인
_first_erb_date = _wdf['date'].min()

if nav_start_date is not None:
    # 지정한 날짜가 ERB 비중 범위 안에 있는지 검증
    if nav_start_date < _first_erb_date:
        print(f'⚠️  nav_start_date({nav_start_date})가 ERB 시작일({_first_erb_date})보다 이릅니다.')
        print(f'   → ERB 시작일({_first_erb_date})로 자동 조정합니다.')
        _nav_start = _first_erb_date
    else:
        _nav_start = nav_start_date
else:
    _nav_start = _first_erb_date

# ERB 종료일
_nav_end = _wdf['date'].max()

# 일간 날짜: nav_start_date 이후이면서 df_dretsft에 존재하는 날짜만
daily_dates = [d for d in df_daily.index
               if _nav_start <= d <= _nav_end
               and d in df_dretsft.index]

print(f'ERB 비중 범위  : {_first_erb_date} ~ {_nav_end}')
print(f'NAV 트래킹 시작: {daily_dates[0]}')
print(f'NAV 트래킹 종료: {daily_dates[-1]}  총 {len(daily_dates)}일')

# ── NAV / REB / SIG / 포지션 배열 초기화 ────────────────────────
# NAV: float 배열,  REB/SIG: float 배열,  pos: (N_days × 13) 배열
N = len(daily_dates)
nav_arr = {p: np.full(N, np.nan) for p in PROFILES}
reb_arr = {p: np.zeros(N)        for p in PROFILES}
sig_arr = {p: np.zeros(N)        for p in PROFILES}
pos_arr = {p: np.zeros((N, 13))  for p in PROFILES}  # 13자산 포지션 금액

# ── 첫날 초기화: 첫 ERB 비중으로 포지션 구성 ─────────────────────
for p in PROFILES:
    w0 = _get_erb_w(p, daily_dates[0])
    if w0 is None:
        nav_arr[p][0] = INIT_NAV
        continue
    # 첫날 포지션 = INIT_NAV × 비중 (비중 합 = 1)
    pos_arr[p][0]  = INIT_NAV * w0
    nav_arr[p][0]  = pos_arr[p][0].sum()
    reb_arr[p][0]  = 1.0  # 첫날 리밸런싱으로 기록

print('첫날 NAV:')
for p in PROFILES:
    print(f'  ERB {p}: {nav_arr[p][0]:.2f}  포지션합={pos_arr[p][0].sum():.2f}')

# ── TSS 함수 ─────────────────────────────────────────────────────

def _ret_row(date):
    """해당 날짜 시차 반영 수익률 배열 (13,)"""
    return df_dretsft.loc[date, TICKER_COLS].values.astype(float)

def _avg_rollret_weight(ix, w13, obs_n):
    """ERB 비중 기준 obs_n일 평균수익률"""
    dates = daily_dates[ix - obs_n + 1: ix + 1]
    rets  = np.array([df_dretsft.loc[d, TICKER_COLS].values for d in dates], dtype=float)
    return float(np.sum(w13 * rets).sum() / obs_n)

def _avg_rollret_nav(p, ix, obs_n):
    """NAV 기준 obs_n일 평균수익률"""
    nav = nav_arr[p][ix - obs_n: ix + 1]
    rr  = np.diff(nav) / nav[:-1]
    return float(rr.mean()) if len(rr) > 0 else 0.0

def _rollvar_weight(ix, w13, obs_n):
    dates = daily_dates[ix - obs_n + 1: ix + 1]
    rets  = np.array([df_dretsft.loc[d, TICKER_COLS].values for d in dates], dtype=float)
    port_rets = np.sum(w13 * rets, axis=1)
    return float(np.square(port_rets).sum() / obs_n)

def _rollvar_nav(p, ix, obs_n):
    nav = nav_arr[p][ix - obs_n: ix + 1]
    rr  = np.diff(nav) / nav[:-1]
    return float(np.square(rr).sum() / obs_n) if len(rr) > 0 else 0.0

def _signaling(ts, var_w, var_n, p, obs_n):
    num   = (var_w / obs_n + var_n / obs_n) ** 2
    denom = ((var_w / obs_n) ** 2 + (var_n / obs_n) ** 2) / max(obs_n - 1, 1)
    df_   = num / denom if denom > 0 else obs_n - 1
    crit  = -tdis.ppf(TSS_THRESHOLD[p] / 2, max(df_, 1))
    return 1.0 if abs(ts) > crit else 0.0

def _isrebal(p, ix, ts, var_w, var_n, obs_n):
    if ix <= COOL_DOWN:
        return 0.0
    reb5  = reb_arr[p][max(0, ix - COOL_DOWN): ix]
    reb40 = reb_arr[p][max(0, ix - FORCE_REB): ix]
    if ix < FORCE_REB and reb5.sum() == 0:
        return _signaling(ts, var_w, var_n, p, obs_n)
    else:
        if reb40.sum() == 0:
            return 1.0          # 강제 리밸런싱
        elif reb5.sum() == 0:
            return _signaling(ts, var_w, var_n, p, obs_n)
        else:
            return 0.0

# ── 메인 NAV 루프 ────────────────────────────────────────────────
print('\n🔄 NAV 트래킹 시작...')

for ix in tqdm(range(1, N), desc='NAV Tracking'):
    dt = daily_dates[ix]
    ret = _ret_row(dt)       # (13,) 오늘 수익률

    for p in PROFILES:
        try:
            erb_w = _get_erb_w(p, dt)
            if erb_w is None:
                nav_arr[p][ix] = nav_arr[p][ix - 1]
                pos_arr[p][ix] = pos_arr[p][ix - 1]
                continue

            # ── NAV 계산 ─────────────────────────────────────────
            lag_ix = ix - LAG_REB
            if lag_ix >= 0 and reb_arr[p][lag_ix] == 1.0:
                # 리밸런싱: LAG_REB 전 시점 NAV × 현재 ERB 비중
                reb_erb_w = _get_erb_w(p, daily_dates[lag_ix])
                if reb_erb_w is None:
                    reb_erb_w = erb_w
                prev_nav       = nav_arr[p][ix - 1]
                pos_arr[p][ix] = prev_nav * reb_erb_w * (1 + ret)
            else:
                # 드리프트: 전일 포지션 × (1 + 수익률)
                pos_arr[p][ix] = pos_arr[p][ix - 1] * (1 + ret)

            nav_arr[p][ix] = pos_arr[p][ix].sum()

            # ── SIG / REB 계산 ───────────────────────────────────
            if ix >= OBS_N + LAG_REB:
                ret_w = _avg_rollret_weight(ix, erb_w, OBS_N)
                ret_n = _avg_rollret_nav(p, ix, OBS_N)
                var_w = _rollvar_weight(ix, erb_w, OBS_N)
                var_n = _rollvar_nav(p, ix, OBS_N)
                denom = np.sqrt(var_w / OBS_N + var_n / OBS_N)
                ts    = float((ret_w - ret_n) / denom) if denom > 0 else 0.0
                sig_arr[p][ix] = _signaling(ts, var_w, var_n, p, OBS_N)
                reb_arr[p][ix] = _isrebal(p, ix, ts, var_w, var_n, OBS_N)

        except Exception as e:
            # 에러 시 전일값 유지
            nav_arr[p][ix] = nav_arr[p][ix - 1]
            pos_arr[p][ix] = pos_arr[p][ix - 1]

# ── DataFrame 변환 + CSV 저장 ────────────────────────────────────
nav_df = {}
for p in PROFILES:
    nav_df[p] = pd.DataFrame({
        'NAV': nav_arr[p],
        'REB': reb_arr[p],
        'SIG': sig_arr[p],
    }, index=daily_dates)
    nav_df[p].to_csv(f'{TRIAL_PREFIX}_nav_{p}.csv')

print('\n✅ NAV 트래킹 완료')
for p in PROFILES:
    valid_nav = nav_arr[p][~np.isnan(nav_arr[p])]
    n_reb     = int(reb_arr[p].sum())
    print(f'  ERB {p}: 최종 NAV={valid_nav[-1]:,.1f}  리밸런싱={n_reb}회')


# In[22]:


# ================================================================
#  Cell FAR-FULL | ERB-only vs 상품매칭 백테스트 비교
#  - 랭킹: DB cpfund_score 우선, 없으면 실시간 계산 (farfmp 1.4.0 기반)
#  - 입력: df_weights (Cell 4), nav_df (Cell 5)
# ================================================================
from scipy.stats import norm

# ── 파라미터 ─────────────────────────────────────────────────────
INIT_NAV_FAR = INIT_NAV
LAG_REB_FAR  = LAG_REB

# ── DB 전체 로드 (한 번만) ───────────────────────────────────────
print('📥 DB 로드 중...')
with sqlite3.connect(dbpath) as conn:
    cpfund_score_all = pd.read_sql_query(
        'SELECT date, category, "운용코드", SCORE, RANK, IS_HOLD '
        'FROM cpfund_score WHERE RANK IS NOT NULL ORDER BY date',
        conn
    )
    cpfund_codes = pd.read_sql_query(
        'SELECT * FROM "cpfund_codes"', conn, index_col='운용코드'
    )
    fund_codes = pd.read_sql_query(
        'SELECT * FROM "fund_codes"', conn, index_col='협회코드'
    )
    bm_px = pd.read_sql_query(
        'SELECT * FROM "bm_px"', conn, index_col='date'
    )
    cpfund_info = pd.read_sql_query(
        'SELECT * FROM "cpfund_info"', conn, index_col='펀드코드'
    )
    erc_hist = pd.read_sql_query(
        'SELECT date, apc_werc FROM "erc_hist" ORDER BY date', conn
    )

# 운용코드 → 펀드명 사전
print('   종목명 사전 구성 중...')
fund_name_map = {}
for op_code in cpfund_codes.index:
    try:
        assoc_code = cpfund_codes.loc[op_code, '협회_퇴직C']
        fund_name_map[op_code] = (
            cpfund_info.loc[assoc_code, '펀드명(Full)']
            if assoc_code in cpfund_info.index else op_code
        )
    except Exception:
        fund_name_map[op_code] = op_code
print(f'   fund_name_map: {len(fund_name_map)}개')

FALLBACK_SCORE_DATE   = '2020-01-04'
FALLBACK_APPLY_BEFORE = '2020-07-17'

# 펀드 가격 전체 로드
all_eval_codes = cpfund_codes.index.tolist()
placeholders   = ','.join(['?'] * len(all_eval_codes))
with sqlite3.connect(dbpath) as conn:
    fund_px_raw = pd.read_sql_query(
        f'SELECT date, code, price FROM fund_px_intg '
        f'WHERE code IN ({placeholders}) '
        f'AND date >= ? AND date <= ?',
        conn, index_col='date',
        params=all_eval_codes + [daily_dates[0], daily_dates[-1]]
    )
fund_px_raw = (fund_px_raw.reset_index()
               .drop_duplicates(['date', 'code'], keep='last')
               .pivot(index='date', columns='code', values='price'))

print(f'   cpfund_score: {len(cpfund_score_all)}행  '
      f'날짜 {cpfund_score_all.date.min()} ~ {cpfund_score_all.date.max()}')
print(f'   fund_px: {fund_px_raw.shape}')

# ── 상수 정의 ────────────────────────────────────────────────────
CATEGORY_NAMES = [
    '글로벌주식','국내주식','미국주식','유럽주식',
    '일본주식','중국본토주식','중국역외주식','신흥국주식',
    '원자재','글로벌리츠','하이일드채권','신흥국채권',
    '선진국채권','국내채권','단기자금','TDF'
]
SLOTLIST = [
    '글로벌주식','국내주식','미국주식','유럽주식',
    '일본주식','중국본토주식','중국역외주식','신흥국주식',
    '원자재','글로벌리츠','하이일드채권','신흥국채권',
    '선진국채권1','선진국채권2','국내채권1','국내채권2',
    '단기자금1','단기자금2'
]
SLOT_TO_CAT = {
    '글로벌주식':   ('글로벌주식',   1),
    '국내주식':     ('국내주식',     1),
    '미국주식':     ('미국주식',     1),
    '유럽주식':     ('유럽주식',     1),
    '일본주식':     ('일본주식',     1),
    '중국본토주식': ('중국본토주식', 1),
    '중국역외주식': ('중국역외주식', 1),
    '신흥국주식':   ('신흥국주식',   1),
    '원자재':       ('원자재',       1),
    '글로벌리츠':   ('글로벌리츠',   1),
    '하이일드채권': ('하이일드채권', 1),
    '신흥국채권':   ('신흥국채권',   1),
    '선진국채권1':  ('선진국채권',   1),
    '선진국채권2':  ('선진국채권',   2),
    '국내채권1':    ('국내채권',     1),
    '국내채권2':    ('국내채권',     2),
    '단기자금1':    ('단기자금',     1),
    '단기자금2':    ('단기자금',     2),
}

# ── Top2 캐시 구성 ───────────────────────────────────────────────
print('🔄 Top2 캐시 구성 중...')
cpfund_score_all['SCORE'] = pd.to_numeric(cpfund_score_all['SCORE'], errors='coerce')
cpfund_score_all['RANK']  = pd.to_numeric(cpfund_score_all['RANK'],  errors='coerce')

score_dates_sorted = sorted(cpfund_score_all['date'].unique())
top2_cache = {}

for dt_s, grp in cpfund_score_all.groupby('date'):
    top2_cache[dt_s] = {}
    for cat, cgrp in grp.groupby('category'):
        if cat not in CATEGORY_NAMES:
            continue
        cgrp_valid = cgrp.dropna(subset=['RANK', 'SCORE'])
        if len(cgrp_valid) == 0:
            continue
        r1 = cgrp_valid[cgrp_valid['RANK'] == 1]['운용코드'].values
        r2 = cgrp_valid[cgrp_valid['RANK'] == 2]['운용코드'].values
        top2_cache[dt_s][cat] = {
            'Rank1':  r1[0] if len(r1) > 0 else None,
            'Rank2':  r2[0] if len(r2) > 0 else (r1[0] if len(r1) > 0 else None),
            'Mean':   float(cgrp_valid['SCORE'].mean()),
            'SD':     float(cgrp_valid['SCORE'].std(ddof=1)) if len(cgrp_valid) > 1 else 0.0,
            'scores': dict(zip(cgrp_valid['운용코드'], cgrp_valid['SCORE'])),
        }

print(f'   Top2 캐시: {len(top2_cache)}개 날짜')

_fallback = top2_cache.get(FALLBACK_SCORE_DATE, {})
print(f'   Fallback 랭킹: {FALLBACK_SCORE_DATE}  카테고리 수: {len(_fallback)}')

# ── 실시간 랭킹 계산 관련 함수 (farfmp 1.4.0 기반) ──────────────

def get_category_repr(dt_ref):
    """날짜별 카테고리-벤치마크 매핑 반환"""
    cols = bm_px.columns
    if dt_ref >= '2020-10-05':
        return {
            '글로벌주식':   (cols[13], 2),
            '국내주식':     (cols[0],  1),
            '미국주식':     (cols[1],  2),
            '유럽주식':     (cols[2],  2),
            '일본주식':     (cols[3],  2),
            '중국본토주식': (cols[14], 2),
            '중국역외주식': (cols[15], 2),
            '신흥국주식':   (cols[5],  2),
            '원자재':       (cols[6],  2),
            '글로벌리츠':   (cols[7],  2),
            '하이일드채권': (cols[8],  2),
            '신흥국채권':   (cols[9],  2),
            '선진국채권':   (cols[10], 2),
            '국내채권':     (cols[11], 1),
            '단기자금':     (cols[12], 1),
            'TDF':          (cols[13], 2),
        }
    else:
        return {
            '글로벌주식':   (cols[13], 2),
            '국내주식':     (cols[0],  1),
            '미국주식':     (cols[1],  2),
            '유럽주식':     (cols[2],  2),
            '일본주식':     (cols[3],  1),
            '중국본토주식': (cols[14], 1),
            '중국역외주식': (cols[15], 1),
            '신흥국주식':   (cols[5],  2),
            '원자재':       (cols[6],  2),
            '글로벌리츠':   (cols[7],  2),
            '하이일드채권': (cols[8],  2),
            '신흥국채권':   (cols[9],  2),
            '선진국채권':   (cols[10], 2),
            '국내채권':     (cols[11], 1),
            '단기자금':     (cols[12], 1),
        }

def get_apc_werc_at(dt_ref):
    """dt_ref 이전 가장 최근 apc_werc 반환"""
    valid = erc_hist[erc_hist['date'] <= dt_ref]
    return float(valid.iloc[-1]['apc_werc']) if not valid.empty else 0.3

def get_ts_bm_for_compute(cat, bm_col, from_date, to_date):
    """일본/중국 기준가 컷오프 보정"""
    ts_bm = bm_px.loc[from_date:to_date, bm_col].dropna()
    if cat in ['일본주식', '중국본토주식', '중국역외주식']:
        adj_head = ts_bm[:'2020-09-28'].iloc[1:]
        adj_tail  = ts_bm['2020-09-28':].iloc[:-1]
        ts_bm = pd.Series(
            index=ts_bm.index[1:],
            data=adj_head.tolist() + adj_tail.tolist()
        )
    return ts_bm

def compute_ranking_at(dt_ref):
    """
    farfmp 1.4.0 로직으로 dt_ref 날짜 랭킹 실시간 계산.
    반환: top2_cache 형식과 동일한 dict
    """
    print(f'   ⚡ {dt_ref} 랭킹 실시간 계산 중...')

    px_dates = pd.Index(sorted(bm_px.index))

    def get_lag_date(dt, lag):
        valid = px_dates[px_dates <= dt]
        return str(valid[-1 - lag]) if len(valid) > lag else str(valid[0])

    def get_from_dates(to_date):
        dt = pd.Timestamp(to_date)
        def shift_month(m):
            target = dt - pd.DateOffset(months=m)
            valid  = px_dates[px_dates >= str(target)[:10]]
            return str(valid[0]) if len(valid) > 0 else str(px_dates[0])
        return {'to': to_date, '1M': shift_month(1),
                '3M': shift_month(3), '6M': shift_month(6), '1Y': shift_month(12)}

    periods  = get_from_dates(dt_ref)
    dt_table = {
        period: {
            0: periods[period],
            1: get_lag_date(periods[period], 1),
            2: get_lag_date(periods[period], 2),
        }
        for period in periods
    }

    # 가중치
    apc_werc  = get_apc_werc_at(dt_ref)
    w_er      = min(max(0.6 - apc_werc, 0.25), 0.8)
    weighting = [w_er, 0.85 - w_er, 0.10, 0.05]  # ER, TE, SIZE, EXP

    category_repr = get_category_repr(dt_ref)

    # 유효 유니버스 필터 (farfmp 1.4.0 동일)
    valid_codes = cpfund_codes[
        cpfund_codes['협회_퇴직C'].isin(cpfund_info.index) &
        cpfund_codes.index.isin(cpfund_info.index)
    ]
    if not fund_codes.empty:
        idx_tgt = fund_codes[
            fund_codes.index.isin(valid_codes['협회_퇴직C']) &
            (fund_codes['매도결제일'] < 10) &
            (fund_codes['펀드위험등급'] != 11) &
            (~fund_codes['퇴직한도'].isna())
        ].index
        valid_codes = valid_codes[valid_codes['협회_퇴직C'].isin(idx_tgt)]

    # 펀드 가격 로드 (1Y ~ dt_ref)
    from_date_1y = dt_table['1Y'][0]
    eval_codes   = valid_codes.index.tolist()
    if not eval_codes:
        return {}
    ph = ','.join(['?'] * len(eval_codes))
    with sqlite3.connect(dbpath) as conn:
        px_tmp = pd.read_sql_query(
            f'SELECT date, code, price FROM fund_px_intg '
            f'WHERE code IN ({ph}) AND date >= ? AND date <= ?',
            conn, params=eval_codes + [from_date_1y, dt_ref]
        )
    px_tmp = (px_tmp.drop_duplicates(['date', 'code'], keep='last')
              .pivot(index='date', columns='code', values='price'))

    result = {}

    for cat in CATEGORY_NAMES:
        if cat not in category_repr:
            continue
        bm_col, lag = category_repr[cat]
        if bm_col not in bm_px.columns:
            continue

        # 카테고리 유니버스
        if cat == '글로벌리츠':
            excluded = ['미래에셋자산운용', '멀티에셋자산운용', '미래에셋운용']
            cat_codes = [
                c for c in valid_codes.index[valid_codes['자산군'] == cat]
                if c in cpfund_info.index and
                   cpfund_info.loc[c, '운용회사명'] not in excluded
            ]
        else:
            cat_codes = valid_codes.index[valid_codes['자산군'] == cat].tolist()

        cat_codes = [c for c in cat_codes if c in px_tmp.columns]
        if not cat_codes:
            continue

        records = []
        for code in cat_codes:
            try:
                er_list, te_list = [], []

                for period_key in ['1M', '3M', '6M', '1Y']:
                    # 펀드 수익률 (lag=0 고정)
                    f_date = dt_table[period_key][0]
                    t_date = dt_table['to'][0]
                    if f_date not in px_tmp.index or t_date not in px_tmp.index:
                        continue
                    f_px = px_tmp.loc[f_date, code]
                    t_px = px_tmp.loc[t_date, code]
                    if pd.isna(f_px) or pd.isna(t_px) or f_px <= 0:
                        continue
                    fund_ret = t_px / f_px - 1

                    # 벤치마크 수익률 (lag 적용)
                    bm_f = dt_table[period_key][lag]
                    bm_t = dt_table['to'][lag]
                    if bm_f not in bm_px.index or bm_t not in bm_px.index:
                        continue
                    bm_ret = bm_px.loc[bm_t, bm_col] / bm_px.loc[bm_f, bm_col] - 1
                    er_list.append(fund_ret - bm_ret)

                    # 추적오차
                    ts_fund = px_tmp.loc[f_date:t_date, code].dropna()
                    ts_bm   = get_ts_bm_for_compute(cat, bm_col, bm_f, bm_t)
                    min_len = min(len(ts_fund), len(ts_bm))
                    if min_len > 1:
                        ret_f = ts_fund.pct_change().dropna().values[-min_len+1:]
                        ret_b = ts_bm.pct_change().dropna().values[-min_len+1:]
                        te_list.append(np.std(ret_f - ret_b))

                # 1Y 포함 4개 기간 모두 있어야 유효 (설정 1년 미만 제외)
                if len(er_list) < 4:
                    continue

                wa_er = float(np.mean(er_list))
                wa_te = float(np.mean(te_list)) if te_list else 0.0

                # AuM, 보수
                assoc_code = valid_codes.loc[code, '협회_퇴직C']
                aum = float(cpfund_info.loc[assoc_code, '패밀리운용규모'])                       if assoc_code in cpfund_info.index else 0.0
                exp = float(cpfund_info.loc[assoc_code, '총보수율'])                       if assoc_code in cpfund_info.index else 2.0
                if assoc_code in fund_codes.index:
                    if fund_codes.loc[assoc_code, '속성구분'] == 8:
                        exp += 0.85

                records.append({'code': code, 'wa_er': wa_er,
                                'wa_te': wa_te, 'aum': aum, 'exp': exp})
            except Exception:
                continue

        if not records:
            continue

        df_cat = pd.DataFrame(records).set_index('code')

        def norm_cdf_score(series, higher_better=True):
            arr = series.astype(float).values
            if len(arr) < 2 or np.nanstd(arr) < 1e-10:
                return pd.Series(0.5, index=series.index)
            mu, sd = np.nanmean(arr), np.nanstd(arr, ddof=1)
            z = (arr - mu) / sd
            return pd.Series(norm.cdf(z) if higher_better else norm.cdf(-z),
                             index=series.index)

        df_cat['S_ER']   = norm_cdf_score(df_cat['wa_er'], True)
        df_cat['S_TE']   = norm_cdf_score(df_cat['wa_te'], False)
        df_cat['S_SIZE'] = norm_cdf_score(df_cat['aum'],   True)
        df_cat['S_EXP']  = norm_cdf_score(df_cat['exp'],   False)
        df_cat['SCORE']  = (df_cat['S_ER']   * weighting[0] +
                            df_cat['S_TE']   * weighting[1] +
                            df_cat['S_SIZE'] * weighting[2] +
                            df_cat['S_EXP']  * weighting[3])

        sorted_codes = df_cat['SCORE'].sort_values(ascending=False).index.tolist()
        result[cat] = {
            'Rank1':  sorted_codes[0] if len(sorted_codes) > 0 else None,
            'Rank2':  sorted_codes[1] if len(sorted_codes) > 1 else
                      sorted_codes[0] if sorted_codes else None,
            'Mean':   float(df_cat['SCORE'].mean()),
            'SD':     float(df_cat['SCORE'].std(ddof=1)) if len(df_cat) > 1 else 0.0,
            'scores': df_cat['SCORE'].to_dict(),
        }

    print(f'   ✅ {dt_ref} 실시간 랭킹 완료 ({len(result)}개 카테고리)')
    return result

# ── get_top2_at (DB 우선, 없으면 실시간 계산) ────────────────────
def get_top2_at(dt_ref):
    if dt_ref < FALLBACK_APPLY_BEFORE:
        return _fallback
    valid = [d for d in score_dates_sorted if d <= dt_ref]
    if valid:
        result = top2_cache.get(valid[-1], {})
        if result:
            return result
    # DB 랭킹 없음 → 실시간 계산
    return compute_ranking_at(dt_ref) or _fallback

# ── 헬퍼 함수 ────────────────────────────────────────────────────
def split_erb_to_slots(erb_w_dict):
    e = erb_w_dict
    us  = e.get('미국주식', 0) + e.get('원자재', 0) / 2
    em  = e.get('신흥국주식', 0) + e.get('원자재', 0) / 2
    eq_sum = us + em + sum(e.get(k, 0) for k in
                           ['국내주식','유럽주식','일본주식','중국주식','글로벌리츠'])
    has_eq = eq_sum > 0.05
    def cap30(v):
        return 0.30 if v > 0.35 else (v - 0.05 if v > 0.30 else v)
    cn = e.get('중국주식', 0)
    d  = {s: 0.0 for s in SLOTLIST}
    d['국내주식']     = e.get('국내주식', 0)  if has_eq else 0.0
    d['미국주식']     = cap30(us)             if has_eq else eq_sum
    d['글로벌주식']   = us - d['미국주식']    if has_eq else 0.0
    d['유럽주식']     = e.get('유럽주식', 0)  if has_eq else 0.0
    d['일본주식']     = e.get('일본주식', 0)  if has_eq else 0.0
    d['중국본토주식'] = (cn/2 if cn > 0.04 else cn)  if has_eq else 0.0
    d['중국역외주식'] = (cn/2 if cn > 0.04 else 0.0) if has_eq else 0.0
    d['신흥국주식']   = em                    if has_eq else 0.0
    d['원자재']       = 0.0
    d['글로벌리츠']   = e.get('글로벌리츠', 0) if has_eq else 0.0
    d['하이일드채권'] = e.get('하이일드채권', 0)
    d['신흥국채권']   = e.get('신흥국채권', 0)
    bd = e.get('선진국채권', 0)
    d['선진국채권1']  = cap30(bd); d['선진국채권2'] = bd - d['선진국채권1']
    kd = e.get('국내채권', 0)
    d['국내채권1']    = cap30(kd); d['국내채권2']   = kd - d['국내채권1']
    cd = e.get('단기자금', 0)
    d['단기자금1']    = min(cd, 0.30); d['단기자금2'] = cd - d['단기자금1']
    return d

def make_fmp(erbdiv, top2):
    fmp = {}
    for slot in SLOTLIST:
        cat, rank = SLOT_TO_CAT[slot]
        if cat not in top2:
            fmp[slot] = {'Code': None, 'Weight': erbdiv[slot]}
            continue
        fmp[slot] = {'Code': top2[cat][f'Rank{rank}'], 'Weight': erbdiv[slot]}
    return fmp

def apply_far(fmp_prv, fmp_cur, top2):
    fmp = {}
    for slot in SLOTLIST:
        cat, rank = SLOT_TO_CAT[slot]
        cur_code  = fmp_cur[slot]['Code']
        cur_wt    = fmp_cur[slot]['Weight']
        prv_code  = fmp_prv.get(slot, {}).get('Code') if fmp_prv else None
        if cur_wt < 1e-7 or prv_code is None or cat not in top2:
            fmp[slot] = {'Code': cur_code, 'Weight': cur_wt}
            continue
        if prv_code == cur_code:
            fmp[slot] = {'Code': cur_code, 'Weight': cur_wt}
            continue
        scores    = top2[cat]['scores']
        prv_score = scores.get(prv_code)
        cur_score = scores.get(cur_code, 0)
        sd_score  = top2[cat]['SD']
        if prv_score is None:
            fmp[slot] = {'Code': cur_code, 'Weight': cur_wt}
        else:
            buf       = norm.ppf(0.7) * sd_score
            threshold = max(top2[cat]['Mean'] + buf, cur_score - buf)
            keep      = float(prv_score) >= threshold
            fmp[slot] = {'Code': (prv_code if keep else cur_code), 'Weight': cur_wt}
    return fmp

# ── 주간 ERB 비중 → 날짜별 dict ─────────────────────────────────
erb_weekly = {p: {} for p in PROFILES}
for _, row in df_weights.iterrows():
    d, p, a, w = row['date'], row['profile'], row['asset'], float(row['weight'])
    if d not in erb_weekly[p]:
        erb_weekly[p][d] = {}
    erb_weekly[p][d][a] = w
erb_weekly_dates = {p: sorted(erb_weekly[p].keys()) for p in PROFILES}

def get_erb_at(p, dt_ref):
    valid = [d for d in erb_weekly_dates[p] if d <= dt_ref]
    return erb_weekly[p][valid[-1]] if valid else None

# ── 메인 백테스트 루프 ───────────────────────────────────────────
print('\n🔄 FAR NAV 백테스트 시작...')

far_nav_results  = {}
far_history_rows = []

for p in tqdm(PROFILES, desc='Profile'):
    N          = len(daily_dates)
    nav_arr    = np.full(N, np.nan)
    nav_arr[0] = INIT_NAV_FAR
    pos_dict   = {}
    fmp_prv    = None
    reb_no     = 0
    reb_ser    = nav_df[p]['REB'].fillna(0).astype(float)

    for ix in range(1, N):
        dt = daily_dates[ix]

        # 드리프트
        if pos_dict:
            try:
                dt_prev = daily_dates[ix - 1]
                px_t    = fund_px_raw.loc[dt,      list(pos_dict.keys())]
                px_t1   = fund_px_raw.loc[dt_prev, list(pos_dict.keys())]
                rets    = ((px_t / px_t1) - 1).fillna(0)
                for code in list(pos_dict.keys()):
                    pos_dict[code] *= (1 + float(rets.get(code, 0)))
            except Exception:
                pass

        # 리밸런싱
        lag_ix = ix - LAG_REB_FAR
        if lag_ix >= 0 and reb_ser.iloc[lag_ix] == 1.0:
            reb_dt = daily_dates[lag_ix]
            erb_w  = get_erb_at(p, reb_dt)
            if erb_w is not None:
                reb_no  += 1
                prev_nav = float(nav_arr[ix - 1])
                erbdiv   = split_erb_to_slots(erb_w)
                top2     = get_top2_at(reb_dt)
                fmp_cur  = make_fmp(erbdiv, top2)
                fmp_new  = apply_far(fmp_prv, fmp_cur, top2)

                pos_dict = {}
                for slot in SLOTLIST:
                    code = fmp_new[slot]['Code']
                    wt   = float(fmp_new[slot]['Weight'])
                    if code is None or wt < 1e-7:
                        continue
                    pos_dict[code] = pos_dict.get(code, 0) + prev_nav * wt

                fmp_prv = fmp_new

                for slot in SLOTLIST:
                    code = fmp_new[slot]['Code']
                    wt   = float(fmp_new[slot]['Weight'])
                    far_history_rows.append({
                        'profile':    p,
                        'reb_no':     reb_no,
                        'date':       dt,
                        'reb_date':   reb_dt,
                        'slot':       slot,
                        'code':       code,
                        'fund_name':  fund_name_map.get(code, code) if code else None,
                        'weight':     round(wt, 6),
                        'nav_at_reb': round(prev_nav, 2),
                    })

        nav_arr[ix] = sum(pos_dict.values()) if pos_dict else nav_arr[ix - 1]

    far_nav_results[p] = pd.Series(nav_arr, index=daily_dates, name='FAR_NAV')
    far_nav_results[p].to_csv(f'{TRIAL_PREFIX}_far_nav_{p}.csv', header=True)

# ── 히스토리 저장 ────────────────────────────────────────────────
df_history = pd.DataFrame(far_history_rows)
if not df_history.empty:
    df_history.to_csv(f'{TRIAL_PREFIX}_far_history.csv', index=False)
    print(f'   history: {len(df_history)}행 저장')

# ── 기여수익률 계산 ──────────────────────────────────────────────
print('🔄 회차별 기여수익률 계산 중...')
contrib_rows = []

for p in PROFILES:
    sub_h = df_history[df_history['profile'] == p].copy()
    if sub_h.empty:
        continue
    reb_dates_p = sorted(sub_h['date'].unique())

    for i, reb_dt in enumerate(reb_dates_p):
        reb_no      = i + 1
        next_reb    = reb_dates_p[i + 1] if i + 1 < len(reb_dates_p) else daily_dates[-1]
        period_days = [d for d in daily_dates if reb_dt <= d <= next_reb]
        if len(period_days) < 2:
            continue

        nav_s     = far_nav_results[p]
        nav_start = float(nav_s.loc[period_days[0]])  if period_days[0]  in nav_s.index else np.nan
        nav_end   = float(nav_s.loc[period_days[-1]]) if period_days[-1] in nav_s.index else np.nan
        fmp_this  = sub_h[sub_h['date'] == reb_dt].set_index('slot')

        for slot in SLOTLIST:
            if slot not in fmp_this.index:
                continue
            row  = fmp_this.loc[slot]
            code = row['code']
            wt   = float(row['weight'])
            name = row['fund_name'] if 'fund_name' in row.index else fund_name_map.get(code, code)
            if code is None or wt < 1e-7:
                continue
            try:
                px_sub   = fund_px_raw.loc[
                    [d for d in period_days if d in fund_px_raw.index], code
                ].dropna()
                fund_ret = float(px_sub.iloc[-1] / px_sub.iloc[0] - 1) if len(px_sub) >= 2 else 0.0
            except Exception:
                fund_ret = 0.0

            contrib_rows.append({
                'profile':      p,
                'reb_no':       reb_no,
                'period_start': period_days[0],
                'period_end':   period_days[-1],
                'slot':         slot,
                'code':         code,
                'fund_name':    name,
                'weight':       round(wt, 4),
                'fund_ret_pct': round(fund_ret * 100, 4),
                'contrib_pct':  round(wt * fund_ret * 100, 4),
                'nav_start':    round(nav_start, 2) if not np.isnan(nav_start) else None,
                'nav_end':      round(nav_end,   2) if not np.isnan(nav_end)   else None,
            })

df_contrib = pd.DataFrame(contrib_rows)
if not df_contrib.empty:
    df_contrib.to_csv(f'{TRIAL_PREFIX}_far_contrib.csv', index=False)
    print(f'   contrib: {len(df_contrib)}행 저장')

# ── 완료 출력 ────────────────────────────────────────────────────
print(f'\n✅ FAR 백테스트 완료')
for p in PROFILES:
    final = float(far_nav_results[p].dropna().iloc[-1])
    n_reb = int(nav_df[p]['REB'].fillna(0).sum())
    print(f'   ERB {p}: 최종 NAV={final:,.1f}  리밸런싱={n_reb}회')
print(f'\n[저장 파일]')
print(f'   {TRIAL_PREFIX}_far_nav_{{p}}.csv')
print(f'   {TRIAL_PREFIX}_far_history.csv')
print(f'   {TRIAL_PREFIX}_far_contrib.csv')


# In[23]:


import sqlite3
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import pandas as pd
import numpy as np

# ── 설정 ──────────────────────────────────────────
TARGET_PROFILE = 80
TOP_N = 20
DB_PATH = dbpath
TRIAL_ID_STR = str(TRIAL_ID).zfill(3)
CSV_PATH = f'trials/ExcessReturn/trial_{TRIAL_ID_STR}_far_contrib.csv'
# ──────────────────────────────────────────────────

# ── Step 1: CSV 로드 + fund_name DB 업데이트 ──────
df_contrib = pd.read_csv(CSV_PATH)
print(f'CSV 로드 완료: {len(df_contrib)}행')

con = sqlite3.connect(DB_PATH)
codes = df_contrib['code'].dropna().unique().tolist()
placeholders = ','.join(['?' for _ in codes])
df_names = pd.read_sql(f"""
    SELECT 펀드코드 AS code, "펀드명(Full)" AS fund_name
    FROM cpfund_info
    WHERE 펀드코드 IN ({placeholders})
""", con, params=codes)
con.close()

name_map = df_names.set_index('code')['fund_name'].to_dict()
df_contrib['fund_name'] = df_contrib['code'].map(name_map)

missing = df_contrib[df_contrib['fund_name'].isna()]['code'].unique()
if len(missing) > 0:
    print(f'⚠️ DB에 없는 코드 {len(missing)}개: {missing}')
else:
    print('✅ 모든 코드 매핑 완료')

df_contrib.to_csv(CSV_PATH, index=False)
print('✅ CSV 저장 완료')

# ── Step 2: 분석 데이터 준비 ──────────────────────
df = df_contrib[df_contrib['profile'] == TARGET_PROFILE].copy()

# 기여율: weight × fund_ret × 100 (%)
df['nav_contrib'] = df['contrib_pct'] * 100

period_map = df.groupby('reb_no')['period_start'].first()

# ── 회차별 실제 NAV 변화율 (nav_start → nav_end) ──
reb_nav_stats = df.groupby('reb_no').agg(
    nav_start=('nav_start', 'first'),
    nav_end=('nav_end', 'first')
).dropna()
reb_nav_stats['ret_pct'] = (
    reb_nav_stats['nav_end'] / reb_nav_stats['nav_start'] - 1
) * 100

top_funds = (
    df.groupby('fund_name')['nav_contrib']
    .sum()
    .nlargest(TOP_N)
    .index.tolist()
)

df_top = df[df['fund_name'].isin(top_funds)]

pivot = (
    df_top.groupby(['reb_no', 'fund_name'])['nav_contrib']
    .sum()
    .unstack(fill_value=0)
)[top_funds]

x_labels = [f"#{r}\n{period_map[r][:7]}" for r in pivot.index]
cmap   = plt.cm.get_cmap('tab20', TOP_N)
colors = [cmap(i) for i in range(TOP_N)]

# ── FAR NAV 회차별 시작값 추출 ────────────────────
far_nav = far_nav_results[str(TARGET_PROFILE)].dropna()
far_nav.index = pd.to_datetime(far_nav.index)

reb_nav_vals = []
for r in pivot.index:
    dt = pd.to_datetime(period_map[r])
    if dt in far_nav.index:
        reb_nav_vals.append(float(far_nav.loc[dt]))
    else:
        nearest = far_nav.index[far_nav.index.get_indexer([dt], method='nearest')[0]]
        reb_nav_vals.append(float(far_nav.loc[nearest]))

last_nav  = float(far_nav.iloc[-1])
nav_x_ext = list(range(len(pivot))) + [len(pivot) - 0.2]
nav_y_ext = reb_nav_vals + [last_nav]

# ── Step 3: 회차별 스택 바 + NAV 변화율 레이블 + NAV 꺾은선 ──
fig1, ax1 = plt.subplots(figsize=(22, 8))
ax1_nav = ax1.twinx()

ax1.set_title(
    f'ERB {TARGET_PROFILE} — 회차별 펀드 기여율 (상위 {TOP_N}) + FAR NAV',
    fontsize=13, fontweight='bold'
)

bottom_pos = np.zeros(len(pivot))
bottom_neg = np.zeros(len(pivot))

for i, col in enumerate(pivot.columns):
    vals     = pivot[col].values
    pos_vals = np.where(vals > 0, vals, 0)
    neg_vals = np.where(vals < 0, vals, 0)
    ax1.bar(range(len(pivot)), pos_vals, bottom=bottom_pos,
            label=col, color=colors[i], alpha=0.80, width=0.7)
    ax1.bar(range(len(pivot)), neg_vals, bottom=bottom_neg,
            color=colors[i], alpha=0.80, width=0.7)
    bottom_pos += pos_vals
    bottom_neg += neg_vals

# ── 회차별 NAV 변화율 레이블 (막대 위/아래) ────────
for xi, r in enumerate(pivot.index):
    if r not in reb_nav_stats.index:
        continue
    ret = reb_nav_stats.loc[r, 'ret_pct']
    # 양수면 막대 위, 음수면 막대 아래
    y_pos  = bottom_pos[xi] if ret >= 0 else bottom_neg[xi]
    offset = 6 if ret >= 0 else -6
    va     = 'bottom' if ret >= 0 else 'top'
    color  = '#2A9D8F' if ret >= 0 else '#E63946'
    ax1.annotate(
        f'{ret:+.1f}%',
        xy=(xi, y_pos),
        xytext=(0, offset),
        textcoords='offset points',
        ha='center', va=va,
        fontsize=9, fontweight='bold', color=color,
        bbox=dict(boxstyle='round,pad=0.2', fc='white', alpha=0.7, ec=color)
    )

ax1.axhline(0, color='black', linewidth=0.8)
ax1.set_xticks(range(len(pivot)))
ax1.set_xticklabels(x_labels, fontsize=7)
ax1.set_ylabel('펀드별 기여율 (%)', fontsize=10)
ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{x:+.1f}%'))
ax1.legend(loc='upper left', fontsize=7, ncol=2,
           bbox_to_anchor=(1.08, 1), borderaxespad=0)

# NAV 꺾은선
ax1_nav.plot(nav_x_ext, nav_y_ext,
             color='#1D3557', linewidth=2.2, marker='D',
             markersize=6, markerfacecolor='white',
             markeredgecolor='#1D3557', markeredgewidth=1.8,
             label='FAR NAV', zorder=5)
for xi, yi in zip(nav_x_ext, nav_y_ext):
    ax1_nav.annotate(f'{yi:,.0f}',
                     xy=(xi, yi), xytext=(0, 10),
                     textcoords='offset points',
                     ha='center', fontsize=7,
                     color='#1D3557', fontweight='bold')
ax1_nav.set_ylabel('FAR NAV', fontsize=10, color='#1D3557')
ax1_nav.tick_params(axis='y', labelcolor='#1D3557')
ax1_nav.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{x:,.0f}'))
ax1_nav.legend(loc='upper right', fontsize=9)

plt.tight_layout()
plt.show()

# ── Step 4: 누적 기여 추이 + NAV 꺾은선 ──────────
fig2, ax2 = plt.subplots(figsize=(22, 7))
ax2_nav = ax2.twinx()

ax2.set_title(
    f'ERB {TARGET_PROFILE} — 누적 기여율 추이 (상위 {TOP_N}) + FAR NAV',
    fontsize=13, fontweight='bold'
)

cumulative = pivot.cumsum()
for i, col in enumerate(cumulative.columns):
    ax2.plot(range(len(cumulative)), cumulative[col].values,
             label=col, color=colors[i], linewidth=1.6, marker='o', markersize=3)

ax2.set_xticks(range(len(cumulative)))
ax2.set_xticklabels(x_labels, fontsize=7)
ax2.set_ylabel('누적 기여율 (%)', fontsize=10)
ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{x:+.1f}%'))
ax2.axhline(0, color='black', linewidth=0.8, linestyle='--')
ax2.legend(loc='upper left', fontsize=7, ncol=2,
           bbox_to_anchor=(1.08, 1), borderaxespad=0)

ax2_nav.plot(nav_x_ext, nav_y_ext,
             color='#1D3557', linewidth=2.5, marker='D',
             markersize=6, markerfacecolor='white',
             markeredgecolor='#1D3557', markeredgewidth=1.8,
             label='FAR NAV', zorder=5)
ax2_nav.set_ylabel('FAR NAV', fontsize=10, color='#1D3557')
ax2_nav.tick_params(axis='y', labelcolor='#1D3557')
ax2_nav.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{x:,.0f}'))
ax2_nav.legend(loc='lower right', fontsize=9)

plt.tight_layout()
plt.show()

# ── Step 5: 히트맵 테이블 ────────────────────────
print(f'\n[ ERB {TARGET_PROFILE} | 기여 상위 {TOP_N} 펀드 × 회차 히트맵 ]')
display_pivot = pivot.T.copy()
display_pivot.columns = [f"#{r} {period_map[r][:7]}" for r in pivot.index]
display_pivot['합계'] = display_pivot.sum(axis=1)
display_pivot = display_pivot.sort_values('합계', ascending=False)

# 마지막 행에 회차별 실제 NAV 변화율 추가
total_row = pd.DataFrame(
    [[reb_nav_stats.loc[r, 'ret_pct'] if r in reb_nav_stats.index else 0.0
      for r in pivot.index] + [reb_nav_stats['ret_pct'].sum()]],
    columns=display_pivot.columns,
    index=['▶ 회차 NAV 수익률']
)
display_pivot = pd.concat([display_pivot, total_row])

display(display_pivot.style
    .format('{:+.2f}%')
    .background_gradient(cmap='RdYlGn', axis=None)
    .set_caption(f'ERB {TARGET_PROFILE} | 행=펀드, 열=회차 | 단위: %')
    .apply(lambda x: ['font-weight: bold; border-top: 2px solid black'] * len(x)
           if x.name == '▶ 회차 NAV 수익률' else [''] * len(x), axis=1)
)


# In[24]:


# ================================================================
#  Cell FAR-VIZ | ERB vs FAR 성과 비교 시각화
#  - 입력: far_nav_results, nav_df, df_history, df_contrib
# ================================================================
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import matplotlib.gridspec as gridspec
import pandas as pd
import numpy as np

PROFILE_COLORS = {
    '80': '#E63946', '65': '#F4A261', '50': '#E9C46A',
    '35': '#2A9D8F', '20': '#457B9D', '10': '#1D3557'
}

# ── 공통 유틸 ────────────────────────────────────────────────────
def calc_stats(nav_series):
    nav = nav_series.dropna()
    if len(nav) < 2:
        return dict(tot=0, ann=0, vol=0, sr=0, mdd=0)
    ret = nav.pct_change().dropna()
    n   = len(nav)
    tot = nav.iloc[-1] / nav.iloc[0] - 1
    ann = (1 + tot) ** (252 / n) - 1
    vol = ret.std() * np.sqrt(252)
    sr  = ann / vol if vol > 0 else 0
    mdd = (nav / nav.cummax() - 1).min()
    return dict(tot=tot, ann=ann, vol=vol, sr=sr, mdd=mdd)

# ================================================================
#  시각화 1 | ERB vs FAR NAV 수익률 차이 (유형별)
# ================================================================
fig1, axes = plt.subplots(3, 2, figsize=(18, 14))
fig1.suptitle('ERB vs FAR — 유형별 NAV 및 수익률 차이',
              fontsize=14, fontweight='bold')
axes = axes.ravel()

for idx, p in enumerate(PROFILES):
    ax  = axes[idx]
    ax2 = ax.twinx()

    erb_nav = nav_df[p]['NAV'].astype(float).dropna()
    erb_nav.index = pd.to_datetime(erb_nav.index)
    far_nav = far_nav_results[p].dropna()
    far_nav.index = pd.to_datetime(far_nav.index)

    # 공통 기간
    common_idx = erb_nav.index.intersection(far_nav.index)
    erb_c = erb_nav.loc[common_idx]
    far_c = far_nav.loc[common_idx]

    # 누적 수익률 차이 (FAR - ERB, %p)
    erb_ret = erb_c / erb_c.iloc[0] - 1
    far_ret = far_c / far_c.iloc[0] - 1
    diff    = (far_ret - erb_ret) * 100

    # NAV 라인
    ax.plot(erb_c.index, erb_c.values,
            color=PROFILE_COLORS[p], linewidth=1.8, label='ERB-only', zorder=3)
    ax.plot(far_c.index, far_c.values,
            color='#222222', linewidth=1.4, linestyle='--', label='FAR', zorder=3)
    ax.set_ylabel('NAV', fontsize=9)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{x:,.0f}'))

    # 수익률 차이 면적
    ax2.fill_between(common_idx, diff.values, 0,
                     where=(diff.values >= 0),
                     color='#2A9D8F', alpha=0.25, label='FAR 우위')
    ax2.fill_between(common_idx, diff.values, 0,
                     where=(diff.values < 0),
                     color='#E63946', alpha=0.25, label='ERB 우위')
    ax2.plot(common_idx, diff.values,
             color='gray', linewidth=0.8, linestyle=':')
    ax2.axhline(0, color='black', linewidth=0.6, linestyle='--', alpha=0.4)
    ax2.set_ylabel('차이 (%p)', fontsize=9)
    ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{x:+.1f}%p'))

    e_s = calc_stats(erb_c)
    f_s = calc_stats(far_c)
    ax.set_title(
        f'ERB {p}  |  ERB {e_s["tot"]*100:.1f}%  FAR {f_s["tot"]*100:.1f}%  '
        f'차이 {(f_s["tot"]-e_s["tot"])*100:+.2f}%p',
        fontsize=10, fontweight='bold'
    )
    ax.legend(loc='upper left', fontsize=8)
    ax2.legend(loc='upper right', fontsize=8)
    ax.spines[['top']].set_visible(False)
    ax.tick_params(axis='x', rotation=20, labelsize=8)
    ax.yaxis.grid(True, linestyle='--', alpha=0.3)

plt.tight_layout()
plt.show()

# ================================================================
#  시각화 2 | 유형별 편입 자산군 비중 시계열
# ================================================================

# df_history에서 자산군별 비중 시계열 추출
SLOT_TO_ASSET = {
    '글로벌주식':   '글로벌주식',
    '국내주식':     '국내주식',
    '미국주식':     '미국주식',
    '유럽주식':     '유럽주식',
    '일본주식':     '일본주식',
    '중국본토주식': '중국주식',
    '중국역외주식': '중국주식',
    '신흥국주식':   '신흥국주식',
    '원자재':       '원자재',
    '글로벌리츠':   '글로벌리츠',
    '하이일드채권': '하이일드채권',
    '신흥국채권':   '신흥국채권',
    '선진국채권1':  '선진국채권',
    '선진국채권2':  '선진국채권',
    '국내채권1':    '국내채권',
    '국내채권2':    '국내채권',
    '단기자금1':    '단기자금',
    '단기자금2':    '단기자금',
}
ASSET_GROUPS = [
    '미국주식','글로벌주식','국내주식','유럽주식','일본주식',
    '중국주식','신흥국주식','글로벌리츠','원자재',
    '하이일드채권','신흥국채권','선진국채권','국내채권','단기자금'
]
ASSET_COLORS = [
    '#E63946','#FF6B6B','#F4A261','#E9C46A','#2A9D8F',
    '#E76F51','#264653','#457B9D','#A8DADC',
    '#6D6875','#B5838D','#1D3557','#48CAE4','#ADE8F4'
]

fig2, axes2 = plt.subplots(3, 2, figsize=(18, 14))
fig2.suptitle('FAR 유형별 편입 자산군 비중 시계열',
              fontsize=14, fontweight='bold')
axes2 = axes2.ravel()

for idx, p in enumerate(PROFILES):
    ax = axes2[idx]
    sub = df_history[df_history['profile'] == p].copy()
    if sub.empty:
        continue

    sub['asset'] = sub['slot'].map(SLOT_TO_ASSET)
    # 회차별 자산군 비중 합산
    wt_pivot = (
        sub.groupby(['date', 'asset'])['weight']
        .sum()
        .unstack(fill_value=0)
        .reindex(columns=ASSET_GROUPS, fill_value=0)
    )
    wt_pivot.index = pd.to_datetime(wt_pivot.index)

    # 스택 영역 차트
    x = np.arange(len(wt_pivot))
    bottom = np.zeros(len(wt_pivot))
    for ai, asset in enumerate(ASSET_GROUPS):
        if asset not in wt_pivot.columns:
            continue
        vals = wt_pivot[asset].values
        if vals.sum() < 1e-7:
            continue
        color = ASSET_COLORS[ai % len(ASSET_COLORS)]
        ax.bar(x, vals, bottom=bottom, label=asset,
               color=color, alpha=0.85, width=0.8)
        bottom += vals

    ax.set_xticks(x)
    ax.set_xticklabels(
        [str(d)[:7] for d in wt_pivot.index],
        rotation=45, ha='right', fontsize=7
    )
    ax.set_ylim(0, 1.05)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{x*100:.0f}%'))
    ax.set_title(f'ERB {p} — 리밸런싱 회차별 자산군 비중', fontsize=10, fontweight='bold')
    ax.set_ylabel('비중')
    ax.legend(loc='upper left', fontsize=6, ncol=2,
              bbox_to_anchor=(1.01, 1), borderaxespad=0)
    ax.yaxis.grid(True, linestyle='--', alpha=0.3)
    ax.spines[['top', 'right']].set_visible(False)

plt.tight_layout()
plt.show()

# ================================================================
#  시각화 3 | 성과 요약 테이블
# ================================================================
perf_rows = []
for p in PROFILES:
    erb_nav = nav_df[p]['NAV'].astype(float).dropna()
    far_nav = far_nav_results[p].dropna()
    e = calc_stats(erb_nav)
    f = calc_stats(far_nav)
    n_reb = len(df_history[df_history['profile'] == p]['reb_no'].unique())             if not df_history.empty else 0

    perf_rows.append({
        '유형':        f'ERB {p}',
        'ERB 누적':    f'{e["tot"]*100:.1f}%',
        'FAR 누적':    f'{f["tot"]*100:.1f}%',
        '누적 차이':   f'{(f["tot"]-e["tot"])*100:+.2f}%p',
        'ERB 연수익':  f'{e["ann"]*100:.2f}%',
        'FAR 연수익':  f'{f["ann"]*100:.2f}%',
        'ERB 변동성':  f'{e["vol"]*100:.2f}%',
        'FAR 변동성':  f'{f["vol"]*100:.2f}%',
        'ERB SR':      f'{e["sr"]:.3f}',
        'FAR SR':      f'{f["sr"]:.3f}',
        'ERB MDD':     f'{e["mdd"]*100:.2f}%',
        'FAR MDD':     f'{f["mdd"]*100:.2f}%',
        'REB 횟수':    n_reb,
    })

perf_df = pd.DataFrame(perf_rows).set_index('유형')

print(f'\n{"="*100}')
print(f'  Trial {TRIAL_ID:03d}  |  성과 요약')
print(f'{"="*100}')

def color_diff(val):
    """누적 차이 양수=초록, 음수=빨강"""
    try:
        v = float(val.replace('%p', '').replace('+', ''))
        color = '#2A9D8F' if v >= 0 else '#E63946'
        return f'color: {color}; font-weight: bold'
    except Exception:
        return ''

display(
    perf_df.style
    .set_properties(**{'text-align': 'center'})
    .applymap(color_diff, subset=['누적 차이'])
    .set_table_styles([
        {'selector': 'th',
         'props': [('background-color', '#1D3557'),
                   ('color', 'white'),
                   ('text-align', 'center'),
                   ('font-size', '11px')]},
        {'selector': 'td',
         'props': [('font-size', '11px')]},
    ])
)


# In[26]:


import sqlite3
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import pandas as pd
import numpy as np

# ── 설정 ──────────────────────────────────────────
TARGET_PROFILE = 80
TARGET_REB_NO  = 45       # 보고 싶은 리밸런싱 회차
DB_PATH = dbpath
TRIAL_ID_STR = str(TRIAL_ID).zfill(3)
CSV_PATH = f'trials/ExcessReturn/trial_{TRIAL_ID_STR}_far_contrib.csv'
# ──────────────────────────────────────────────────

# ── 데이터 준비 ───────────────────────────────────
df_contrib = pd.read_csv(CSV_PATH)
df = df_contrib[
    (df_contrib['profile'] == TARGET_PROFILE) &
    (df_contrib['reb_no']  == TARGET_REB_NO)
].copy()

if df.empty:
    print(f'⚠️ ERB {TARGET_PROFILE} / {TARGET_REB_NO}회차 데이터 없음')
else:
    period_start = df['period_start'].iloc[0]
    period_end   = df['period_end'].iloc[0]
    nav_start    = df['nav_start'].iloc[0]
    nav_end      = df['nav_end'].iloc[0]
    nav_ret      = (nav_end / nav_start - 1) * 100

    print(f'ERB {TARGET_PROFILE} | #{TARGET_REB_NO}회차')
    print(f'기간: {period_start} ~ {period_end}')
    print(f'NAV: {nav_start:,.1f} → {nav_end:,.1f}  ({nav_ret:+.2f}%)')

    # 이 회차에 편입된 종목 (비중 > 0)
    df_slots = df[df['weight'] > 1e-4].copy()
    codes     = df_slots['code'].dropna().unique().tolist()

    # 종목별 가격 시계열 로드
    placeholders = ','.join(['?'] * len(codes))
    with sqlite3.connect(DB_PATH) as con:
        px_tmp = pd.read_sql_query(
            f'SELECT date, code, price FROM fund_px_intg '
            f'WHERE code IN ({placeholders}) '
            f'AND date >= ? AND date <= ?',
            con, params=codes + [period_start, period_end]
        )
    px_tmp = (px_tmp
              .drop_duplicates(['date', 'code'], keep='last')
              .pivot(index='date', columns='code', values='price'))
    px_tmp.index = pd.to_datetime(px_tmp.index)

    # FAR NAV 시계열 (해당 구간)
    far_nav = far_nav_results[str(TARGET_PROFILE)].dropna()
    far_nav.index = pd.to_datetime(far_nav.index)
    far_nav_period = far_nav[
        (far_nav.index >= pd.to_datetime(period_start)) &
        (far_nav.index <= pd.to_datetime(period_end))
    ]

    # 기준가 정규화 (시작일=100)
    def normalize(series):
        s = series.dropna()
        if len(s) == 0 or s.iloc[0] == 0:
            return s
        return (s / s.iloc[0] - 1) * 100  # 수익률 %

    nav_norm = normalize(far_nav_period)

    # 코드 → 펀드명 매핑
    with sqlite3.connect(DB_PATH) as con:
        df_names = pd.read_sql(
            f'SELECT 펀드코드 AS code, "펀드명(Full)" AS fund_name '
            f'FROM cpfund_info WHERE 펀드코드 IN ({placeholders})',
            con, params=codes
        )
    name_map = df_names.set_index('code')['fund_name'].to_dict()

    # ── 시각화 ───────────────────────────────────
    fig, (ax_top, ax_bot) = plt.subplots(
        2, 1, figsize=(18, 12),
        gridspec_kw={'height_ratios': [2.5, 1]}
    )
    fig.suptitle(
        f'ERB {TARGET_PROFILE}  |  #{TARGET_REB_NO}회차  '
        f'({period_start} ~ {period_end})  '
        f'NAV {nav_start:,.0f} → {nav_end:,.0f}  ({nav_ret:+.2f}%)',
        fontsize=13, fontweight='bold'
    )

    # ── 상단: 종목별 수익률 꺾은선 ───────────────
    cmap   = plt.cm.get_cmap('tab20', len(codes))
    colors = [cmap(i) for i in range(len(codes))]

    for i, row in df_slots.iterrows():
        code = row['code']
        wt   = row['weight']
        name = name_map.get(code, code)
        slot = row['slot']
        if code not in px_tmp.columns:
            continue
        ret_series = normalize(px_tmp[code])
        if ret_series.empty:
            continue
        final_ret = ret_series.iloc[-1]
        color = colors[list(codes).index(code) % len(colors)]
        ax_top.plot(ret_series.index, ret_series.values,
                    linewidth=1.6, color=color,
                    label=f'{slot} | {name[:12]} ({wt*100:.1f}%) → {final_ret:+.1f}%')
        # 종료 시점 레이블
        ax_top.annotate(
            f'{final_ret:+.1f}%',
            xy=(ret_series.index[-1], ret_series.iloc[-1]),
            xytext=(5, 0), textcoords='offset points',
            fontsize=7, color=color, va='center', fontweight='bold'
        )

    # FAR NAV 굵게
    ax_top.plot(nav_norm.index, nav_norm.values,
                linewidth=2.8, color='#1D3557', linestyle='--',
                label=f'FAR NAV ({nav_ret:+.2f}%)', zorder=5)
    ax_top.annotate(
        f'NAV {nav_ret:+.2f}%',
        xy=(nav_norm.index[-1], nav_norm.iloc[-1]),
        xytext=(5, 0), textcoords='offset points',
        fontsize=9, color='#1D3557', fontweight='bold', va='center'
    )

    ax_top.axhline(0, color='black', linewidth=0.8, linestyle=':', alpha=0.5)
    ax_top.set_ylabel('수익률 (%)', fontsize=10)
    ax_top.yaxis.set_major_formatter(
        mticker.FuncFormatter(lambda x, _: f'{x:+.1f}%')
    )
    ax_top.legend(loc='upper left', fontsize=7.5,
                  bbox_to_anchor=(1.01, 1), borderaxespad=0)
    ax_top.yaxis.grid(True, linestyle='--', alpha=0.3)
    ax_top.spines[['top', 'right']].set_visible(False)
    ax_top.tick_params(axis='x', rotation=20, labelsize=8)

    # ── 하단: 종목별 기여율 바 차트 ──────────────
    bar_data = df_slots[['slot', 'code', 'fund_name', 'weight',
                          'fund_ret_pct', 'contrib_pct']].copy()
    bar_data['contrib_pct_show']  = bar_data['contrib_pct'] * 100
    bar_data['fund_ret_pct_show'] = bar_data['fund_ret_pct']
    bar_data = bar_data.sort_values('contrib_pct_show', ascending=False).reset_index(drop=True)

    x     = np.arange(len(bar_data))
    width = 0.4

    # colors_bar — code 컬럼 직접 참조
    bar_colors = [
        colors[list(codes).index(c) % len(colors)] if c in codes else 'gray'
        for c in bar_data['code']
    ]

    bars1 = ax_bot.bar(x - width/2, bar_data['fund_ret_pct_show'],
                        width, label='펀드 수익률',
                        color=bar_colors, alpha=0.5, edgecolor='white')
    bars2 = ax_bot.bar(x + width/2, bar_data['contrib_pct_show'],
                        width, label='기여율',
                        color=bar_colors, alpha=0.9, edgecolor='white')
    for _, brow in bar_data.iterrows():
        slot = brow['slot']
        matched = df_slots[df_slots['slot'] == slot]
        if not matched.empty:
            code = matched.iloc[0]['code']
            if code in codes:
                bar_colors.append(colors[list(codes).index(code) % len(colors)])
            else:
                bar_colors.append('gray')
        else:
            bar_colors.append('gray')

    bars1 = ax_bot.bar(x - width/2, bar_data['fund_ret_pct_show'],
                        width, label='펀드 수익률',
                        color=bar_colors, alpha=0.5, edgecolor='white')
    bars2 = ax_bot.bar(x + width/2, bar_data['contrib_pct_show'],
                        width, label='기여율',
                        color=bar_colors, alpha=0.9, edgecolor='white')

    # 바 위 수치
    for bar in bars1:
        h = bar.get_height()
        if abs(h) > 0.05:
            ax_bot.text(bar.get_x() + bar.get_width()/2, h,
                        f'{h:+.1f}%', ha='center',
                        va='bottom' if h >= 0 else 'top',
                        fontsize=6.5)
    for bar in bars2:
        h = bar.get_height()
        if abs(h) > 0.05:
            ax_bot.text(bar.get_x() + bar.get_width()/2, h,
                        f'{h:+.2f}%', ha='center',
                        va='bottom' if h >= 0 else 'top',
                        fontsize=6.5, fontweight='bold')

    ax_bot.set_xticks(x)
    ax_bot.set_xticklabels(
        [f"{r['slot']}\n({r['weight']*100:.0f}%)"
         for _, r in bar_data.iterrows()],
        fontsize=7, rotation=30, ha='right'
    )
    ax_bot.axhline(0, color='black', linewidth=0.8)
    ax_bot.set_ylabel('수익률 / 기여율 (%)', fontsize=9)
    ax_bot.yaxis.set_major_formatter(
        mticker.FuncFormatter(lambda x, _: f'{x:+.1f}%')
    )
    ax_bot.legend(fontsize=8)
    ax_bot.yaxis.grid(True, linestyle='--', alpha=0.3)
    ax_bot.spines[['top', 'right']].set_visible(False)

    plt.tight_layout()
    plt.show()

    # ── 수치 테이블 ──────────────────────────────
    print(f'\n[ ERB {TARGET_PROFILE} #{TARGET_REB_NO}회차 종목별 상세 ]')
    tbl = df_slots[['slot', 'fund_name', 'weight',
                    'fund_ret_pct', 'contrib_pct']].copy()
    tbl['weight']      = (tbl['weight'] * 100).round(2)
    tbl['fund_ret_pct']= tbl['fund_ret_pct'].round(4)
    tbl['contrib_pct'] = (tbl['contrib_pct'] * 100).round(4)
    tbl.columns        = ['슬롯', '펀드명', '비중(%)',
                           '펀드수익률(%)', '기여율(%)']
    tbl = tbl.sort_values('기여율(%)', ascending=False).reset_index(drop=True)

    display(tbl.style
        .format({'비중(%)': '{:.2f}%',
                 '펀드수익률(%)': '{:+.2f}%',
                 '기여율(%)': '{:+.3f}%'})
        .background_gradient(cmap='RdYlGn', subset=['펀드수익률(%)', '기여율(%)'])
        .bar(subset=['비중(%)'], color='#ADE8F4')
    )


# In[ ]:




