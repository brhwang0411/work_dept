'''
ERB(Efficient Risk Budgeting) Asset Allocation Model by Eunseok Yang

- Version  : 2.3.1
- 업데이트 : 2021.07.08
- ERB Weights(Raw) 저장모듈 추가


'''

# 필요 module & data import
import numpy as np
import pandas as pd
import os, sqlite3
import win32com.client as win32
from datetime import datetime
from scipy.optimize import minimize


################
# 파라미터 설정 #
################

# Path & Hyper Parameters
dbpath = 'D:/flaskapps/rasite/db/gqsdb.sqlite3'
dbpath_bm_weight = 'D:/flaskapps/rasite/db/gqsoffdb.sqlite3'
xlpath = 'D:/flaskapps/rasite/models/ERB_Allocation.xlsm'
numasset = 12 # 자산군 갯수(단기자금 제외)
window = 52 # 시계열 윈도우
decayfactor = 0.94 # EWMA 적용 Lambda
assetcap = 0.60
assetfloor = 0.01
pd.set_option('display.float_format', '{:.2f}'.format)

# ERB Profile Parameters
erb_profile = {
'80': [2.0,    0.0,    0.1],
'65': [2.0,    0.0,    0.5],
'50': [2.0,    0.0,    1.0],
'35': [1.0,    0.0,    1.0],
'20': [1.0,    1.0,    1.0],
'10': [0.0,  'MDD',    1.0]
}

erb_profile2 = {
'80': [2.0,   0.0,   0.1,   0.80], # 알파, 베타, 감마, 델타
'65': [2.0,   0.0,   0.2,   0.65],
'50': [2.0,   0.0,   0.3,   0.50],
'35': [1.0,   0.0,   0.3,   0.50],
'20': [1.0,   0.0,   1.0,   0.20],
'10': [0.0, 'MDD',   1.0,   0.35]
}

kr_target_wgt = 0.15

###########
# 함수정의 #
###########

# 기준일에 적용되는 BM Policy Weight
def get_bm_weight(dt_when):
    conn = sqlite3.connect(dbpath_bm_weight)
    dt_list = pd.read_sql_query('SELECT "date" FROM bm_weight', conn, index_col='date').index.tolist()
    dt_list.sort(reverse=True)
    bm_weight = pd.DataFrame
    for i in dt_list:
        if dt_when >= i:
            bm_weight = pd.read_sql_query("""SELECT * FROM bm_weight WHERE date='{}'""".format(i), conn, index_col='date').values.ravel().tolist()
            ms_norm_eq_kr = [kr_target_wgt]
            ms_norm_eq_nonkr = [bm_weight[10]-bm_weight[3]]+bm_weight[5:8]+[bm_weight[8]-bm_weight[-1]]+[bm_weight[9]]+[bm_weight[3]]
            ms_norm_eq_nonkr = [e * (100 - kr_target_wgt) / (100 - bm_weight[-1]) for e in ms_norm_eq_nonkr]
            ms_norm_hybnd = [bm_weight[11]+bm_weight[13]/4]+[bm_weight[12]+bm_weight[13]/4]
            ms_norm = ms_norm_eq_kr + ms_norm_eq_nonkr + ms_norm_hybnd
            ms_norm = ms_norm+[(100-sum(ms_norm[8:10]))/2]*2
            ms_roff = ms_norm[:6]+[0.0]+[ms_norm[7]]+[0.0,0.0]+[50+sum(bm_weight[11:13])]+[50-sum(bm_weight[11:13])]
            return ms_norm, ms_roff
            break
    if bm_weight.empty:
        print('참고할 BM Policy 비중이 존재하지 않습니다.')
    conn.close()

# Read DB
def read_table(table_name):
    with sqlite3.connect(dbpath) as conn:
        df = pd.read_sql_query('SELECT * FROM "{}"'.format(table_name), conn, index_col='date')
    return df

# Get Weight from DB
def get_weight(table_name, date):
    with sqlite3.connect(dbpath) as conn:
        weight = pd.read_sql_query('SELECT * FROM "{}" WHERE date="{}"'.format(table_name, date), conn, index_col='date')
    return weight

# Make Input Table
def make_input(df_ts, dt_when):
    num_when = np.where(df_ts.index==dt_when)[0][0]
    df_input = df_ts[num_when-window:num_when+1]
    return df_input

# Decay Factor array
def make_array_decay(decayfactor):
    array_decay = np.ones(1)*(1-decayfactor)
    while len(array_decay) < window:
        array_decay = np.append(array_decay,array_decay[-1] * decayfactor)
    array_decay = np.sort(array_decay)
    return array_decay

# Portfolio Risk
def portrisk(weight, covmat):
    weight = np.array(weight)
    portfolio_risk = (weight.T @ covmat @ weight) ** 0.5
    return portfolio_risk

# Assets Risk Contribution(%)
def riskcontribution(weight, covmat):
    weight = np.array(weight)
    portfolio_risk = portrisk(weight, covmat)
    marginal_rc = 1 / portfolio_risk * (covmat @ weight)
    rc_abs = weight * marginal_rc
    rc_pct = rc_abs / rc_abs.sum()
    return rc_pct

# ERC Objective Function
def objective_rb(weight, args):
    weight = np.array(weight)
    covmat = args[0]
    rc_target = args[1]
    error = np.sum(np.square(riskcontribution(weight, covmat) - rc_target))
    return error

# get Risk Budgeted Portfolio
def rbweight(covmat, rc_target, initial_weight):
    constraints = ({'type': 'eq', 'fun': lambda x: np.sum(x) - 1.0},
                   {'type': 'ineq', 'fun': lambda x: x})
    options = {'ftol': 1e-12, 'maxiter': 500, 'disp' : False}
    b = (0.0 , 1.0)
    bnds = (b,) * numasset
    opt_result = minimize(fun = objective_rb,
                          args = [covmat, rc_target],
                          x0 = initial_weight,
                          bounds = bnds,
                          method = 'SLSQP',
                          constraints = constraints,
                          options = options)
    return opt_result.x

# ERB Alternative Objective Function
def objective_rb_alt(weight):
    weight = np.array(weight)
    error = np.sum(np.square(weight))
    return error

# get Risk Budgeted Portfolio Alternatively
def rbweight_alt(covmat, rc_target, initial_weight):
    constraints = ({'type': 'eq', 'fun': lambda x: np.sum(x) - 1.0},
                   {'type': 'ineq', 'fun': lambda x: x},
                   {'type': 'eq', 'fun': lambda x: np.sum(np.square(riskcontribution(x, covmat) - rc_target))})
    options = {'ftol': 1e-12, 'maxiter': 500, 'disp' : False}
    b = (0.0 , 1.0)
    bnds = (b,) * numasset
    opt_result = minimize(fun = objective_rb_alt,
                          x0 = initial_weight,
                          bounds = bnds,
                          method = 'SLSQP',
                          constraints = constraints,
                          options = options)
    return opt_result.x


################
# 주요변수 생성 #
################

# 데이터베이스 읽기
df_weekly = read_table('idx_weekly')
print(df_weekly.tail())

# 자산배분 기준일 정의
dt_when = df_weekly.index[-1]
dt_when = '2025-06-15'
print("자산배분 기준일은", dt_when, "입니다.")

# Market Size by Asset Classes
bm_weight = get_bm_weight(dt_when)
ms_norm = pd.Series(
    data=bm_weight[0],
    index=['국내주식', '미국주식', '유럽주식', '일본주식', '중국주식', '신흥국주식', '원자재', '글로벌리츠', '하이일드채권', '신흥국채권', '선진국채권', '국내채권']
    )

ms_roff = pd.Series(
    data=bm_weight[1],
    index=['국내주식', '미국주식', '유럽주식', '일본주식', '중국주식', '신흥국주식', '원자재', '글로벌리츠', '하이일드채권', '신흥국채권', '선진국채권', '국내채권']
    )

# 기준일로부터 최근 52주 input table 생성
df_input = make_input(df_weekly, dt_when).iloc[:,:numasset]

# Weekly Return 생성
df_input_ret = df_input.pct_change(1)[1:]

# Decay Factor 적용 array 생성
array_decay = make_array_decay(decayfactor)

# Risk free rate
rf_rate = df_weekly['기준금리'][dt_when]/100

# Realized Returns
hist_ret = df_input.iloc[-1]/df_input.iloc[0]-1
avg_ret = np.mean(df_input_ret)
mom_ret = (df_input.iloc[-1]/df_input.iloc[-21])**(52/(21))-1

# Realized Volatility
hist_sd = np.sqrt(52)*np.std(df_input_ret)
ewma_sd = np.sqrt(np.sum((df_input_ret**2)*np.tile(array_decay, numasset).reshape(numasset, window).T)*52)
mdd = np.min((df_input/np.maximum.accumulate(df_input))-1)

# Upside Volatility
up_count = np.sum(np.ceil(np.maximum(df_input_ret-avg_ret,0)))
up_avg = np.sum(np.ceil(np.maximum(df_input_ret-avg_ret,0))*df_input_ret)/up_count
up_vol = np.sqrt((np.sum((np.ceil(np.maximum(df_input_ret-avg_ret,0))*df_input_ret-up_avg)**2)-(window-up_count)*(up_avg**2))*window/up_count)

# Downside Volatility
down_count = np.sum(np.ceil(np.maximum(avg_ret-df_input_ret,0)))
down_avg = np.sum(np.ceil(np.maximum(avg_ret-df_input_ret,0))*df_input_ret)/down_count
down_vol = np.sqrt((np.sum((np.ceil(np.maximum(avg_ret-df_input_ret,0))*df_input_ret-down_avg)**2)-(window-down_count)*(down_avg**2))*window/down_count)


# Adjusted Volatility
# sd_adj_delta = ((ewma_sd - hist_sd).where(ewma_sd - hist_sd >= 0)*down_vol/up_vol).replace(np.nan,(ewma_sd-hist_sd).where(ewma_sd-hist_sd<0)*up_vol/down_vol)
sd_adj_delta = ((ewma_sd-hist_sd).where(ewma_sd-hist_sd>=0)*down_vol/up_vol).replace(np.nan, 0) + ((ewma_sd-hist_sd).where(ewma_sd-hist_sd<0)*up_vol/down_vol).replace(np.nan, 0)
sd_adj_min = np.sqrt(52)*np.std(make_input(df_weekly, dt_when).단기자금.pct_change(1)[1:])
adj_sd = np.maximum(hist_sd+sd_adj_delta, sd_adj_min)

# Correlation & Covariance Matrix
hist_cor = df_input_ret.corr()
adj_cov = np.matrix(adj_sd).T*np.matrix(adj_sd)*hist_cor


################
# ERC 비중 산출 #
################

# IVP: Initial Weight
ivp = 1/adj_sd/np.sum(1/adj_sd)

# 직전 ERC Weight 불러오기
erc0 = get_weight('erc_hist', df_weekly.index[np.where(df_weekly.index==dt_when)[0][0]-1]).iloc[:,:-2]

# ERC: Equilibrium Weight
erc_ivp = rbweight(adj_cov, np.repeat(1/numasset,numasset), ivp)
erc_erc0 = rbweight(adj_cov, np.repeat(1/numasset,numasset), erc0.values.ravel())

if objective_rb(erc_ivp, args = (adj_cov, np.repeat(1/numasset,numasset))) < 1e-6:
    erc = erc_ivp.tolist()
elif objective_rb(erc_erc0, args = (adj_cov, np.repeat(1/numasset,numasset))) < 1e-6:
    erc = erc_erc0.tolist()
else:
    print('ERC Portfolio is invalid')
    exit()

# ERC로부터 위험회피계수 Lambda(λ) 추정
erc_risk = portrisk(erc, adj_cov)
erc_sr = (np.sum(hist_ret*erc)-rf_rate)/erc_risk
erc_lambda = (erc_sr/erc_risk) if erc_sr >= 0 else ((np.sum(mom_ret*erc)-rf_rate)/erc_risk**2)
apc_werc = ((hist_cor.sum()-1)/(numasset-1)*erc).sum()

# DB 저장
erc.insert(0, dt_when)
erc = erc+[erc_sr]+[apc_werc]
with sqlite3.connect(dbpath) as conn:
    # cur = conn.cursor()
    # cur.execute("""DELETE FROM erc_hist WHERE date='{}'""".format(dt_when))
    # cur.execute('INSERT INTO "erc_hist" VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)', erc)
    # conn.commit()
    print('ERC Portfolio has been validated and saved')
erc = erc[1:-2]


################
# ERB 비중 산출 #
################

# # ERB by Excel
# xltmp = win32.Dispatch('Excel.Application')
# xltmp.Visible = True
# # wb_tmp = xltmp.Workbooks.Open('D:/flaskapps/rasite/models/temp.xlsx')
# # wb_tmp.Close(False)
#
# excel = win32.Dispatch('Excel.Application')
# excel.Visible = False
# wb = excel.Workbooks.Open(xlpath)
# ws1 = wb.Worksheets('Weekly')
# ws2 = wb.Worksheets('Allocator')
# ws80 = wb.Worksheets('W(80)')
# ws65 = wb.Worksheets('W(65)')
# ws50 = wb.Worksheets('W(50)')
# ws35 = wb.Worksheets('W(35)')
# ws20 = wb.Worksheets('W(20)')
# ws10 = wb.Worksheets('W(10)')
#
# rows_w = np.count_nonzero(np.ravel(ws1.Range('A:A').Value))
# rows_80 = np.count_nonzero(np.ravel(ws80.Range('A:A').Value))
# dt_recent = ws1.Range('A1').Cells(rows_w+1, 1).Value.date()
#
# with sqlite3.connect(dbpath) as conn:
#     df_weekly_add = pd.read_sql_query('SELECT * FROM idx_weekly WHERE "date">"{}"'.format(dt_recent), conn, index_col=None).iloc[:1,:-2]
#
# if len(df_weekly_add.values)==0:
#     erbxl = {
#     '80': np.ravel(ws80.Range(ws80.Range('A1').Cells(rows_80+1,2),ws80.Range('A1').Cells(rows_80+1,13)).Value).tolist(),
#     '65': np.ravel(ws65.Range(ws65.Range('A1').Cells(rows_80+1,2),ws65.Range('A1').Cells(rows_80+1,13)).Value).tolist(),
#     '50': np.ravel(ws50.Range(ws50.Range('A1').Cells(rows_80+1,2),ws50.Range('A1').Cells(rows_80+1,13)).Value).tolist(),
#     '35': np.ravel(ws35.Range(ws35.Range('A1').Cells(rows_80+1,2),ws35.Range('A1').Cells(rows_80+1,13)).Value).tolist(),
#     '20': np.ravel(ws20.Range(ws20.Range('A1').Cells(rows_80+1,2),ws20.Range('A1').Cells(rows_80+1,13)).Value).tolist(),
#     '10': np.ravel(ws10.Range(ws10.Range('A1').Cells(rows_80+1,2),ws10.Range('A1').Cells(rows_80+1,13)).Value).tolist()
#     }
#
# else:
#     with sqlite3.connect(dbpath_bm_weight) as conn:
#         bm_weight = pd.read_sql_query('SELECT * FROM bm_weight WHERE "date">"{}"'.format(dt_recent), conn, index_col=None).iloc[:1,:]
#
#     for i in range(16):
#         ws1.Range('A1').Cells(rows_w+2,i+1).Value = df_weekly_add.iloc[0].values[i]
#     ws2.Range('BH32').Value = bm_weight.미국리츠[0]
#     ws2.Range('BH33').Value = bm_weight.유럽주식[0]
#     ws2.Range('BH34').Value = bm_weight.일본주식[0]
#     ws2.Range('BH35').Value = bm_weight.중국주식[0]
#     ws2.Range('BH36').Value = bm_weight.신흥국주식[0]
#     ws2.Range('BH37').Value = bm_weight.원자재[0]
#     ws2.Range('BH38').Value = bm_weight.금[0]
#     ws2.Range('BH39').Value = bm_weight.하이일드채권[0]
#     ws2.Range('BH40').Value = bm_weight.신흥국채권[0]
#     ws2.Range('BH41').Value = bm_weight.미국회사채[0]
#     ws2.Range('BH42').Value = bm_weight.한국주식[0]
#
#     try:
#         excel.Application.Run('All_In_One')
#     except:
#         pass
#
#     erbxl = {
#     '80': np.ravel(ws80.Range(ws80.Range('A1').Cells(rows_80+2,2),ws80.Range('A1').Cells(rows_80+2,13)).Value).tolist(),
#     '65': np.ravel(ws65.Range(ws65.Range('A1').Cells(rows_80+2,2),ws65.Range('A1').Cells(rows_80+2,13)).Value).tolist(),
#     '50': np.ravel(ws50.Range(ws50.Range('A1').Cells(rows_80+2,2),ws50.Range('A1').Cells(rows_80+2,13)).Value).tolist(),
#     '35': np.ravel(ws35.Range(ws35.Range('A1').Cells(rows_80+2,2),ws35.Range('A1').Cells(rows_80+2,13)).Value).tolist(),
#     '20': np.ravel(ws20.Range(ws20.Range('A1').Cells(rows_80+2,2),ws20.Range('A1').Cells(rows_80+2,13)).Value).tolist(),
#     '10': np.ravel(ws10.Range(ws10.Range('A1').Cells(rows_80+2,2),ws10.Range('A1').Cells(rows_80+2,13)).Value).tolist()
#     }
#
# erbxl = pd.DataFrame(index=ms_norm.index, data=erbxl, dtype='f')
# print('ERB by Excel 도출이 완료되었습니다.')

# Reverse Optimization (find μ)
mu_hat = (erc_lambda*adj_cov)@erc

if dt_when < '2020-06-07': # before

    # adj correlations & covariance matrix
    adj_cor = {
    '80': np.maximum(hist_cor, (hist_cor*erb_profile['80'][2])),
    '65': np.maximum(hist_cor, (hist_cor*erb_profile['65'][2])),
    '50': np.maximum(hist_cor, (hist_cor*erb_profile['50'][2])),
    '35': np.maximum(hist_cor, (hist_cor*erb_profile['35'][2])),
    '20': np.maximum(hist_cor, (hist_cor*erb_profile['20'][2])),
    '10': np.maximum(hist_cor, (hist_cor*erb_profile['10'][2]))
    }

    adj_cov2 = {
    '80': np.matrix(adj_sd).T*np.matrix(adj_sd)*adj_cor['80'],
    '65': np.matrix(adj_sd).T*np.matrix(adj_sd)*adj_cor['65'],
    '50': np.matrix(adj_sd).T*np.matrix(adj_sd)*adj_cor['50'],
    '35': np.matrix(adj_sd).T*np.matrix(adj_sd)*adj_cor['35'],
    '20': np.matrix(adj_sd).T*np.matrix(adj_sd)*adj_cor['20'],
    '10': np.matrix(adj_sd).T*np.matrix(adj_sd)*adj_cor['10']
    }

    # Risk Budgeting Rule
    sumpos = np.sum([0 if r < 0 else r for r in mu_hat])
    array_bgt = pd.Series(np.ones(numasset), index = ivp.index)
    array_bgt['원자재':'글로벌리츠'] = 0.5

    bgt_score = {
    '80': 1/adj_sd*array_bgt if sumpos==0 else mu_hat**erb_profile['80'][0]/adj_sd**erb_profile['80'][1]*array_bgt,
    '65': 1/adj_sd*array_bgt if sumpos==0 else mu_hat**erb_profile['65'][0]/adj_sd**erb_profile['65'][1]*array_bgt,
    '50': 1/adj_sd*array_bgt if sumpos==0 else mu_hat**erb_profile['50'][0]/adj_sd**erb_profile['50'][1]*array_bgt,
    '35': 1/adj_sd*array_bgt if sumpos==0 else mu_hat**erb_profile['35'][0]/adj_sd**erb_profile['35'][1]*array_bgt,
    '20': 1/adj_sd*array_bgt if sumpos==0 else mu_hat**erb_profile['20'][0]/adj_sd**erb_profile['20'][1]*array_bgt,
    '10': mu_hat**erb_profile['10'][0]/-mdd*array_bgt
    }

    rc_target = {
    '80': bgt_score['80']/sum(bgt_score['80']),
    '65': bgt_score['65']/sum(bgt_score['65']),
    '50': bgt_score['50']/sum(bgt_score['50']),
    '35': bgt_score['35']/sum(bgt_score['35']),
    '20': bgt_score['20']/sum(bgt_score['20']),
    '10': bgt_score['10']/sum(bgt_score['10'])
    }

    # calculate erb
    erb = {
    '80': rbweight(adj_cov2['80'], rc_target['80'], erc),
    '65': rbweight(adj_cov2['65'], rc_target['65'], erc),
    '50': rbweight(adj_cov2['50'], rc_target['50'], erc),
    '35': rbweight(adj_cov2['35'], rc_target['35'], erc),
    '20': rbweight(adj_cov2['20'], rc_target['20'], erc),
    '10': rbweight(adj_cov2['10'], rc_target['10'], erc)
    }

else: # after

    # 정책비중 책정
    mksize = ms_norm if erc_lambda > 0 else ms_roff
    weight_policy = {
    '80': pd.concat([mksize[:-4]/sum(mksize[:-4])*erb_profile2['80'][3]*100,mksize[-4:]/sum(mksize[-4:])*(1-erb_profile2['80'][3])*100]),
    '65': pd.concat([mksize[:-4]/sum(mksize[:-4])*erb_profile2['65'][3]*100,mksize[-4:]/sum(mksize[-4:])*(1-erb_profile2['65'][3])*100]),
    '50': pd.concat([mksize[:-4]/sum(mksize[:-4])*erb_profile2['50'][3]*100,mksize[-4:]/sum(mksize[-4:])*(1-erb_profile2['50'][3])*100]),
    '35': pd.concat([mksize[:-4]/sum(mksize[:-4])*erb_profile2['35'][3]*100,mksize[-4:]/sum(mksize[-4:])*(1-erb_profile2['35'][3])*100]),
    '20': pd.concat([mksize[:-4]/sum(mksize[:-4])*erb_profile2['20'][3]*100,mksize[-4:]/sum(mksize[-4:])*(1-erb_profile2['20'][3])*100]),
    '10': pd.concat([mksize[:-4]/sum(mksize[:-4])*erb_profile2['10'][3]*100,mksize[-4:]/sum(mksize[-4:])*(1-erb_profile2['10'][3])*100])
    }

    # adj correlations & covariance matrix
    ## if less than 1: pick positive correl + adjusted smaller-in-abs negative correl
    adj_cor = {
    '80': np.maximum(hist_cor, (hist_cor*erb_profile2['80'][2])),
    '65': np.maximum(hist_cor, (hist_cor*erb_profile2['65'][2])),
    '50': np.maximum(hist_cor, (hist_cor*erb_profile2['50'][2])),
    '35': np.maximum(hist_cor, (hist_cor*erb_profile2['35'][2])),
    '20': np.maximum(hist_cor, (hist_cor*erb_profile2['20'][2])),
    '10': np.maximum(hist_cor, (hist_cor*erb_profile2['10'][2]))
    }

    avg_pair_cor = {
    '80': (adj_cor['80'].sum()-1)/(numasset-1),
    '65': (adj_cor['65'].sum()-1)/(numasset-1),
    '50': (adj_cor['50'].sum()-1)/(numasset-1),
    '35': (adj_cor['35'].sum()-1)/(numasset-1),
    '20': (adj_cor['20'].sum()-1)/(numasset-1),
    '10': (adj_cor['10'].sum()-1)/(numasset-1)
    }

    wtd_sum_apc = {
    '80': (avg_pair_cor['80']*weight_policy['80']).sum()/100,
    '65': (avg_pair_cor['65']*weight_policy['65']).sum()/100,
    '50': (avg_pair_cor['50']*weight_policy['50']).sum()/100,
    '35': (avg_pair_cor['35']*weight_policy['35']).sum()/100,
    '20': (avg_pair_cor['20']*weight_policy['20']).sum()/100,
    '10': (avg_pair_cor['10']*weight_policy['10']).sum()/100
    }

    adj_cov2 = {
    '80': np.matrix(adj_sd).T * np.matrix(adj_sd) * adj_cor['80'],
    '65': np.matrix(adj_sd).T * np.matrix(adj_sd) * adj_cor['65'],
    '50': np.matrix(adj_sd).T * np.matrix(adj_sd) * adj_cor['50'],
    '35': np.matrix(adj_sd).T * np.matrix(adj_sd) * adj_cor['35'],
    '20': np.matrix(adj_sd).T * np.matrix(adj_sd) * adj_cor['20'],
    '10': np.matrix(adj_sd).T * np.matrix(adj_sd) * adj_cor['10']
    }

    # Risk Budgeting Rule
    sumpos = np.sum([0 if r < 0 else r for r in mu_hat])
    array_bgt = pd.Series(np.ones(numasset), index = ivp.index)
    array_bgt['원자재':'글로벌리츠'] = 0.5

    bgt_score = {
    '80': 1/adj_sd*array_bgt if sumpos==0 else mu_hat**erb_profile2['80'][0]/adj_sd**erb_profile2['80'][1]*array_bgt,
    '65': 1/adj_sd*array_bgt if sumpos==0 else mu_hat**erb_profile2['65'][0]/adj_sd**erb_profile2['65'][1]*array_bgt,
    '50': 1/adj_sd*array_bgt if sumpos==0 else mu_hat**erb_profile2['50'][0]/adj_sd**erb_profile2['50'][1]*array_bgt,
    '35': 1/adj_sd*array_bgt if sumpos==0 else mu_hat**erb_profile2['35'][0]/adj_sd**erb_profile2['35'][1]*array_bgt,
    '20': 1/adj_sd*array_bgt if sumpos==0 else mu_hat**erb_profile2['20'][0]/adj_sd**erb_profile2['20'][1]*array_bgt,
    '10': mu_hat**erb_profile2['10'][0]/-mdd*array_bgt
    }

    bgt_score = {
    '80': bgt_score['80']*weight_policy['80']**(0.5+wtd_sum_apc['80']),
    '65': bgt_score['65']*weight_policy['65']**(0.5+wtd_sum_apc['65']),
    '50': bgt_score['50']*weight_policy['50']**(0.5+wtd_sum_apc['50']),
    '35': bgt_score['35']*weight_policy['35']**(0.5+wtd_sum_apc['35']),
    '20': bgt_score['20']*weight_policy['20']**(0.5+wtd_sum_apc['20']),
    '10': bgt_score['10']*weight_policy['10']**(0.5+wtd_sum_apc['10'])
    }

    for t in bgt_score:
        for c in bgt_score[t].index:
            bgt_score[t][c] = 0.0 if bgt_score[t][c]<=0.1 else round(bgt_score[t][c],6)

    rc_target = {
    '80': bgt_score['80']/sum(bgt_score['80']),
    '65': bgt_score['65']/sum(bgt_score['65']),
    '50': bgt_score['50']/sum(bgt_score['50']),
    '35': bgt_score['35']/sum(bgt_score['35']),
    '20': bgt_score['20']/sum(bgt_score['20']),
    '10': bgt_score['10']/sum(bgt_score['10'])
    }

    # calculate erb
    erb = {
    '80': rbweight(adj_cov2['80'], rc_target['80'], erc),
    '65': rbweight(adj_cov2['65'], rc_target['65'], erc),
    '50': rbweight(adj_cov2['50'], rc_target['50'], erc),
    '35': rbweight(adj_cov2['35'], rc_target['35'], erc),
    '20': rbweight(adj_cov2['20'], rc_target['20'], erc),
    '10': rbweight(adj_cov2['10'], rc_target['10'], erc)
    }

# 비중 오류 검사
if objective_rb(erb['80'], args = (adj_cov2['80'], rc_target['80'])) < 1e-11:
    print('Python: ERB_80 Portfolio has been validated')
else:
    erb['80'] = rbweight_alt(adj_cov2['80'], rc_target['80'], erc)
    if objective_rb(erb['80'], args = (adj_cov2['80'], rc_target['80'])) < 1e-11:
        print('Python: ERB_80 Portfolio has been validated')
    else:
        erb['80'] = rbweight(adj_cov2['80'], rc_target['80'], ivp)
        if objective_rb(erb['80'], args = (adj_cov2['80'], rc_target['80'])) < 1e-11:
            print('Python: ERB_80 Portfolio has been validated')
        else:
            erb['80'] = rbweight(adj_cov2['80'], rc_target['80'], get_weight('erb80_hist', df_weekly.index[np.where(df_weekly.index==dt_when)[0][0]-1]).values.ravel()[:-1])
            if objective_rb(erb['80'], args = (adj_cov2['80'], rc_target['80'])) < 1e-11:
                print('Python: ERB_80 Portfolio has been validated')
            else:
                print('Python: ERB_80 Portfolio - Failure')

if objective_rb(erb['65'], args = (adj_cov2['65'], rc_target['65'])) < 1e-11:
    print('Python: ERB_65 Portfolio has been validated')
else:
    erb['65'] = rbweight_alt(adj_cov2['65'], rc_target['65'], erc)
    if objective_rb(erb['65'], args = (adj_cov2['65'], rc_target['65'])) < 1e-11:
        print('Python: ERB_65 Portfolio has been validated')
    else:
        erb['65'] = rbweight(adj_cov2['65'], rc_target['65'], ivp)
        if objective_rb(erb['65'], args = (adj_cov2['65'], rc_target['65'])) < 1e-11:
            print('Python: ERB_65 Portfolio has been validated')
        else:
            erb['65'] = rbweight(adj_cov2['65'], rc_target['65'], get_weight('erb65_hist', df_weekly.index[np.where(df_weekly.index==dt_when)[0][0]-1]).values.ravel()[:-1])
            if objective_rb(erb['65'], args = (adj_cov2['65'], rc_target['65'])) < 1e-11:
                print('Python: ERB_65 Portfolio has been validated')
            else:
                print('Python: ERB_65 Portfolio - Failure')

if objective_rb(erb['50'], args = (adj_cov2['50'], rc_target['50'])) < 1e-11:
    print('Python: ERB_50 Portfolio has been validated')
else:
    erb['50'] = rbweight_alt(adj_cov2['50'], rc_target['50'], erc)
    if objective_rb(erb['50'], args = (adj_cov2['50'], rc_target['50'])) < 1e-11:
        print('Python: ERB_50 Portfolio has been validated')
    else:
        erb['50'] = rbweight(adj_cov2['50'], rc_target['50'], ivp)
        if objective_rb(erb['50'], args = (adj_cov2['50'], rc_target['50'])) < 1e-11:
            print('Python: ERB_50 Portfolio has been validated')
        else:
            erb['50'] = rbweight(adj_cov2['50'], rc_target['50'], get_weight('erb50_hist', df_weekly.index[np.where(df_weekly.index==dt_when)[0][0]-1]).values.ravel()[:-1])
            if objective_rb(erb['50'], args = (adj_cov2['50'], rc_target['50'])) < 1e-11:
                print('Python: ERB_50 Portfolio has been validated')
            else:
                print('Python: ERB_50 Portfolio - Failure')

if objective_rb(erb['35'], args = (adj_cov2['35'], rc_target['35'])) < 1e-11:
    print('Python: ERB_35 Portfolio has been validated')
else:
    erb['35'] = rbweight_alt(adj_cov2['35'], rc_target['35'], erc)
    if objective_rb(erb['35'], args = (adj_cov2['35'], rc_target['35'])) < 1e-11:
        print('Python: ERB_35 Portfolio has been validated')
    else:
        erb['35'] = rbweight(adj_cov2['35'], rc_target['35'], ivp)
        if objective_rb(erb['35'], args = (adj_cov2['35'], rc_target['35'])) < 1e-11:
            print('Python: ERB_35 Portfolio has been validated')
        else:
            erb['35'] = rbweight(adj_cov2['35'], rc_target['35'], get_weight('erb35_hist', df_weekly.index[np.where(df_weekly.index==dt_when)[0][0]-1]).values.ravel()[:-1])
            if objective_rb(erb['35'], args = (adj_cov2['35'], rc_target['35'])) < 1e-11:
                print('Python: ERB_35 Portfolio has been validated')
            else:
                print('Python: ERB_35 Portfolio - Failure')

if objective_rb(erb['20'], args = (adj_cov2['20'], rc_target['20'])) < 1e-11:
    print('Python: ERB_20 Portfolio has been validated')
else:
    erb['20'] = rbweight_alt(adj_cov2['20'], rc_target['20'], erc)
    if objective_rb(erb['20'], args = (adj_cov2['20'], rc_target['20'])) < 1e-11:
        print('Python: ERB_20 Portfolio has been validated')
    else:
        erb['20'] = rbweight(adj_cov2['20'], rc_target['20'], ivp)
        if objective_rb(erb['20'], args = (adj_cov2['20'], rc_target['20'])) < 1e-11:
            print('Python: ERB_20 Portfolio has been validated')
        else:
            erb['20'] = rbweight(adj_cov2['20'], rc_target['20'], get_weight('erb20_hist', df_weekly.index[np.where(df_weekly.index==dt_when)[0][0]-1]).values.ravel()[:-1])
            if objective_rb(erb['20'], args = (adj_cov2['20'], rc_target['20'])) < 1e-11:
                print('Python: ERB_20 Portfolio has been validated')
            else:
                print('Python: ERB_20 Portfolio - Failure')

if objective_rb(erb['10'], args = (adj_cov2['10'], rc_target['10'])) < 1e-11:
    print('Python: ERB_10 Portfolio has been validated')
else:
    erb['10'] = rbweight_alt(adj_cov2['10'], rc_target['10'], erc)
    if objective_rb(erb['10'], args = (adj_cov2['10'], rc_target['10'])) < 1e-11:
        print('Python: ERB_10 Portfolio has been validated')
    else:
        erb['10'] = rbweight(adj_cov2['10'], rc_target['10'], ivp)
        if objective_rb(erb['10'], args = (adj_cov2['10'], rc_target['10'])) < 1e-11:
            print('Python: ERB_10 Portfolio has been validated')
        else:
            erb['10'] = rbweight(adj_cov2['10'], rc_target['10'], get_weight('erb10_hist', df_weekly.index[np.where(df_weekly.index==dt_when)[0][0]-1]).values.ravel()[:-1])
            if objective_rb(erb['10'], args = (adj_cov2['10'], rc_target['10'])) < 1e-11:
                print('Python: ERB_10 Portfolio has been validated')
            else:
                print('Python: ERB_10 Portfolio - Failure')

# # Excel or Python
# ws2.Range('AK46:AK49').Value = ws2.Range('AK52:AK55').Value
# for i in range(12):
#     ws2.Range('AR'+'{}'.format(32+i)).Value = erb['80'][i]
# err80_py = ws2.Range('AR44').Value
# err80_xl = ws80.Range('N'+'{}'.format(rows_80+1)).Value if len(df_weekly_add.values)==0 else ws80.Range('N'+'{}'.format(rows_80+2)).Value
# if err80_py>err80_xl:
#     erb['80'] = erbxl['80'].values
#
# ws2.Range('AK46:AK49').Value = ws2.Range('AK56:AK59').Value
# for i in range(12):
#     ws2.Range('AR'+'{}'.format(32+i)).Value = erb['65'][i]
# err65_py = ws2.Range('AR44').Value
# err65_xl = ws65.Range('N'+'{}'.format(rows_80+1)).Value if len(df_weekly_add.values)==0 else ws65.Range('N'+'{}'.format(rows_80+2)).Value
# if err65_py>err65_xl:
#     erb['65'] = erbxl['65'].values
#
# ws2.Range('AK46:AK49').Value = ws2.Range('AK60:AK63').Value
# for i in range(12):
#     ws2.Range('AR'+'{}'.format(32+i)).Value = erb['50'][i]
# err50_py = ws2.Range('AR44').Value
# err50_xl = ws50.Range('N'+'{}'.format(rows_80+1)).Value if len(df_weekly_add.values)==0 else ws50.Range('N'+'{}'.format(rows_80+2)).Value
# if err50_py>err50_xl:
#     erb['50'] = erbxl['50'].values
#
# ws2.Range('AK46:AK49').Value = ws2.Range('AK64:AK67').Value
# for i in range(12):
#     ws2.Range('AR'+'{}'.format(32+i)).Value = erb['35'][i]
# err35_py = ws2.Range('AR44').Value
# err35_xl = ws35.Range('N'+'{}'.format(rows_80+1)).Value if len(df_weekly_add.values)==0 else ws35.Range('N'+'{}'.format(rows_80+2)).Value
# if err35_py>err35_xl:
#     erb['35'] = erbxl['35'].values
#
# ws2.Range('AK46:AK49').Value = ws2.Range('AK68:AK71').Value
# for i in range(12):
#     ws2.Range('AR'+'{}'.format(32+i)).Value = erb['20'][i]
# err20_py = ws2.Range('AR44').Value
# err20_xl = ws20.Range('N'+'{}'.format(rows_80+1)).Value if len(df_weekly_add.values)==0 else ws20.Range('N'+'{}'.format(rows_80+2)).Value
# if err20_py>err20_xl:
#     erb['20'] = erbxl['20'].values
#
# ws2.Range('AK46:AK49').Value = ws2.Range('AK72:AK75').Value
# for i in range(12):
#     ws2.Range('AR'+'{}'.format(32+i)).Value = erb['10'][i]
# err10_py = ws2.Range('AR44').Value
# err10_xl = ws10.Range('N'+'{}'.format(rows_80+1)).Value if len(df_weekly_add.values)==0 else ws10.Range('N'+'{}'.format(rows_80+2)).Value
# if err10_py>err10_xl:
#     erb['10'] = erbxl['10'].values
#
# wb.Close(False)
# excel.Application.Quit()
# del ws1, ws2, ws80, ws65,ws50, ws35, ws20, ws10, wb, wb_tmp, xltmp, excel
# os.system('taskkill /f /im excel.exe')


################
# ERB 비중 저장 #
################

# erb modification
if dt_when < '2020-06-07':
    # before
    assetcap = np.round(1 / np.sqrt(numasset), 2)
    erb_mod = np.minimum(pd.DataFrame(erb, index = ivp.index), assetcap)
    erb_cash = 1 - np.sum(erb_mod)
    erb_mod = erb_mod.append(erb_cash, ignore_index = True)
    assetnames = ivp.index.values.tolist()
    assetnames.append('단기자금')
    erb_mod.index = assetnames
    print(erb_mod)

else:
    # after
    erb_bef_mod = pd.DataFrame(erb, index = ivp.index)
    erb_aft_mod = pd.DataFrame(data=None, columns=erb, index = ivp.index)

    erb_aft_mod.loc['국내주식'][erb_bef_mod.loc['국내주식']<assetfloor] = 0.0
    erb_aft_mod.loc['국내주식'][erb_bef_mod.loc['국내주식']>=assetfloor] = erb_bef_mod.loc['국내주식']
    erb_aft_mod.loc['국내주식'][erb_aft_mod.loc['국내주식']>=assetcap] = assetcap

    erb_aft_mod.loc['유럽주식'][erb_bef_mod.loc['유럽주식']<assetfloor] = 0.0
    erb_aft_mod.loc['유럽주식'][erb_bef_mod.loc['유럽주식']>=assetfloor] = erb_bef_mod.loc['유럽주식']
    erb_aft_mod.loc['유럽주식'][erb_aft_mod.loc['유럽주식']>=assetcap] = assetcap

    erb_aft_mod.loc['일본주식'][erb_bef_mod.loc['일본주식']<assetfloor] = 0.0
    erb_aft_mod.loc['일본주식'][erb_bef_mod.loc['일본주식']>=assetfloor] = erb_bef_mod.loc['일본주식']
    erb_aft_mod.loc['일본주식'][erb_aft_mod.loc['일본주식']>=assetcap] = assetcap

    erb_aft_mod.loc['중국주식'][erb_bef_mod.loc['중국주식']<assetfloor] = 0.0
    erb_aft_mod.loc['중국주식'][erb_bef_mod.loc['중국주식']>=assetfloor] = erb_bef_mod.loc['중국주식']
    erb_aft_mod.loc['중국주식'][erb_aft_mod.loc['중국주식']>=assetcap] = assetcap

    erb_aft_mod.loc['원자재'][erb_bef_mod.loc['원자재']<assetfloor] = 0.0
    erb_aft_mod.loc['원자재'][erb_bef_mod.loc['원자재']>=assetfloor] = erb_bef_mod.loc['원자재']
    erb_aft_mod.loc['원자재'][erb_aft_mod.loc['원자재']>=assetcap] = assetcap

    erb_aft_mod.loc['글로벌리츠'][erb_bef_mod.loc['글로벌리츠']<assetfloor] = 0.0
    erb_aft_mod.loc['글로벌리츠'][erb_bef_mod.loc['글로벌리츠']>=assetfloor] = erb_bef_mod.loc['글로벌리츠']
    erb_aft_mod.loc['글로벌리츠'][erb_aft_mod.loc['글로벌리츠']>=assetcap] = assetcap

    erb_aft_mod.loc['신흥국주식'][pd.concat([erb_bef_mod.iloc[:1],erb_bef_mod.iloc[4:7]]).sum()-pd.concat([erb_aft_mod.iloc[:1],erb_aft_mod.iloc[4:5],erb_aft_mod.iloc[6:7]]).sum()<assetfloor] = 0.0
    erb_aft_mod.loc['신흥국주식'][pd.concat([erb_bef_mod.iloc[:1],erb_bef_mod.iloc[4:7]]).sum()-pd.concat([erb_aft_mod.iloc[:1],erb_aft_mod.iloc[4:5],erb_aft_mod.iloc[6:7]]).sum()>=assetfloor] = pd.concat([erb_bef_mod.iloc[:1],erb_bef_mod.iloc[4:7]]).sum()-pd.concat([erb_aft_mod.iloc[:1],erb_aft_mod.iloc[4:5],erb_aft_mod.iloc[6:7]]).sum()
    erb_aft_mod.loc['신흥국주식'][erb_aft_mod.loc['신흥국주식']>=assetcap] = assetcap

    erb_aft_mod.loc['미국주식'] = erb_bef_mod.iloc[:8].sum()-pd.concat([erb_aft_mod.iloc[:1],erb_aft_mod.iloc[2:8]]).sum()
    erb_aft_mod.loc['미국주식'][erb_aft_mod.loc['미국주식']>=assetcap] = assetcap

    erb_aft_mod.loc['하이일드채권'][erb_bef_mod.loc['하이일드채권']<assetfloor] = 0.0
    erb_aft_mod.loc['하이일드채권'][erb_bef_mod.loc['하이일드채권']>=assetfloor] = erb_bef_mod.loc['하이일드채권']
    erb_aft_mod.loc['하이일드채권'][erb_aft_mod.loc['하이일드채권']>=assetcap] = assetcap

    erb_aft_mod.loc['신흥국채권'][erb_bef_mod.loc['신흥국채권']<assetfloor] = 0.0
    erb_aft_mod.loc['신흥국채권'][erb_bef_mod.loc['신흥국채권']>=assetfloor] = erb_bef_mod.loc['신흥국채권']
    erb_aft_mod.loc['신흥국채권'][erb_aft_mod.loc['신흥국채권']>=assetcap] = assetcap

    erb_aft_mod.loc['선진국채권'] = erb_bef_mod.iloc[8:11].sum()-erb_aft_mod.iloc[8:10].sum()
    erb_aft_mod.loc['선진국채권'][erb_bef_mod.loc['선진국채권']<assetfloor] = 0.0
    erb_aft_mod.loc['선진국채권'][erb_aft_mod.loc['선진국채권']>=assetcap] = assetcap

    erb_aft_mod.loc['국내채권'][erb_bef_mod.loc['국내채권']<assetfloor] = 0.0
    erb_aft_mod.loc['국내채권'][erb_bef_mod.loc['국내채권']>=assetfloor] = erb_bef_mod.loc['국내채권']
    erb_aft_mod.loc['국내채권'][erb_aft_mod.loc['국내채권']>=assetcap] = assetcap

    erb_aft_cash = 1.0 - erb_aft_mod.sum()
    for t in erb_aft_cash.index:
        erb_aft_cash[t] = erb_aft_cash[t] if erb_aft_cash[t] >=0.0 else 0.0
    erb_aft_mod = erb_aft_mod.append(erb_aft_cash, ignore_index=True)
    assetnames = ivp.index.values.tolist()
    assetnames.append('단기자금')
    erb_aft_mod.index = assetnames
    erb_mod = erb_aft_mod
    print(erb_mod*100)

erb80 = erb_mod['80'].tolist()
erb65 = erb_mod['65'].tolist()
erb50 = erb_mod['50'].tolist()
erb35 = erb_mod['35'].tolist()
erb20 = erb_mod['20'].tolist()
erb10 = erb_mod['10'].tolist()

erb80.insert(0, dt_when)
erb65.insert(0, dt_when)
erb50.insert(0, dt_when)
erb35.insert(0, dt_when)
erb20.insert(0, dt_when)
erb10.insert(0, dt_when)

# # ERB 결과값 저장
# with sqlite3.connect(dbpath) as conn:
#     cur = conn.cursor()
#     cur.execute("""DELETE FROM erb80_hist_raw WHERE date='{}'""".format(dt_when))
#     cur.execute('INSERT INTO "erb80_hist_raw" VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)', [dt_when]+erb['80'].tolist())
#     cur.execute("""DELETE FROM erb65_hist_raw WHERE date='{}'""".format(dt_when))
#     cur.execute('INSERT INTO "erb65_hist_raw" VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)', [dt_when]+erb['65'].tolist())
#     cur.execute("""DELETE FROM erb50_hist_raw WHERE date='{}'""".format(dt_when))
#     cur.execute('INSERT INTO "erb50_hist_raw" VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)', [dt_when]+erb['50'].tolist())
#     cur.execute("""DELETE FROM erb35_hist_raw WHERE date='{}'""".format(dt_when))
#     cur.execute('INSERT INTO "erb35_hist_raw" VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)', [dt_when]+erb['35'].tolist())
#     cur.execute("""DELETE FROM erb20_hist_raw WHERE date='{}'""".format(dt_when))
#     cur.execute('INSERT INTO "erb20_hist_raw" VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)', [dt_when]+erb['20'].tolist())
#     cur.execute("""DELETE FROM erb10_hist_raw WHERE date='{}'""".format(dt_when))
#     cur.execute('INSERT INTO "erb10_hist_raw" VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)', [dt_when]+erb['10'].tolist())
#     cur.execute("""DELETE FROM erb80_hist WHERE date='{}'""".format(dt_when))
#     cur.execute('INSERT INTO "erb80_hist" VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)', erb80)
#     cur.execute("""DELETE FROM erb65_hist WHERE date='{}'""".format(dt_when))
#     cur.execute('INSERT INTO "erb65_hist" VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)', erb65)
#     cur.execute("""DELETE FROM erb50_hist WHERE date='{}'""".format(dt_when))
#     cur.execute('INSERT INTO "erb50_hist" VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)', erb50)
#     cur.execute("""DELETE FROM erb35_hist WHERE date='{}'""".format(dt_when))
#     cur.execute('INSERT INTO "erb35_hist" VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)', erb35)
#     cur.execute("""DELETE FROM erb20_hist WHERE date='{}'""".format(dt_when))
#     cur.execute('INSERT INTO "erb20_hist" VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)', erb20)
#     cur.execute("""DELETE FROM erb10_hist WHERE date='{}'""".format(dt_when))
#     cur.execute('INSERT INTO "erb10_hist" VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)', erb10)
#     conn.commit()
#     print('ERB Portfolios are saved')
