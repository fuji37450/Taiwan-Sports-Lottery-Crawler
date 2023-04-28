import requests
from datetime import datetime
import pytz
import pandas as pd
import time
from copy import deepcopy

utc_plus_8 = pytz.timezone('Asia/Taipei')

def color_different_red(val):
    prev = val.shift(-1)
    is_equal = val == prev
    is_equal[val.last_valid_index()] = True
    return [f'color: black' if v else 'color: red' for v in is_equal]

def get_data(game_time, atn_name, htn_name, all_data, odds_column):
  now_time = datetime.now().astimezone(utc_plus_8).strftime('%H:%M:%S')
  name = f"{game_time.strftime('%m%d,%H%M')}{atn_name}-{htn_name}"

  if name in all_data:
    data = all_data[name]
  else:  
    data = dict()
    data['比賽日期'] = game_time.date()
    data['比賽時間'] = game_time.strftime('%H:%M')
    data['客隊隊伍名稱'] = atn_name
    data['主隊隊伍名稱'] = htn_name
    
    data['當有更新時的更新時間'] = now_time
    data['客隊勝分差當前賠率'] = [0] * len(odds_column['atn'])
    data['主隊勝分差當前賠率'] = [0] * len(odds_column['htn'])
    data['客隊更新時間'] = []
    for col_name in odds_column['atn']:
      data[col_name] = []
    data['主隊更新時間'] = []
    for col_name in odds_column['htn']:
      data[col_name] = []

  data['客隊更新時間'].insert(0, now_time)
  data['主隊更新時間'].insert(0, now_time)

  return data, name

def crawler(filename, odds_column, all_data):
  requestID = '4092' if filename == 'mlb' else '4104'
  res = requests.get(f'https://blob.sportslottery.com.tw/apidata/Pre/ListByLeague/{requestID}.json')
  games = res.json()
  for game in games:
    if game['li'] == 0: # 冠軍賽
      continue
    game_time = datetime.fromtimestamp(game['kdt']/1000.0).astimezone(utc_plus_8)
    odds = next(filter(lambda g: g['id'] == 10, game['ms']), None) # 勝分差玩法
    if odds == None:
      continue
    else:
      odds = odds['cs']
    
    data, name = get_data(game_time, game['atn'][0], game['htn'][0], all_data, odds_column)

    updated = False
    for i, odd in enumerate(odds):
      prefix = odds_column['atn'][i][odds_column['atn'][i].find('贏'):odds_column['atn'][i].find('變')]
      data['客隊勝分差當前賠率'][i] = f"{prefix}: {odd[0]['o']}"
      if len(data[odds_column['atn'][i]]) and data[odds_column['atn'][i]][0] != float(odd[0]['o']):
        updated = True
      data[odds_column['atn'][i]].insert(0, float(odd[0]['o']))
      data['主隊勝分差當前賠率'][i] = f"{prefix}: {odd[1]['o']}"
      if len(data[odds_column['htn'][i]]) and data[odds_column['htn'][i]][0] != float(odd[1]['o']):
        updated = True
      data[odds_column['htn'][i]].insert(0, float(odd[1]['o']))
    
    if updated:
      data['當有更新時的更新時間'] = data['客隊更新時間'][0]

    all_data[name] = deepcopy(data)
  
  now = datetime.now().astimezone(utc_plus_8).strftime('%H:%M:%S')
  if len(all_data.keys()) == 0:
    print(now, 'No data for', filename)
  else:
    writer = pd.ExcelWriter(f'{filename}.xlsx')
    for key, data in all_data.items():
      t = dict()
      for k, v in data.items():
        t[k] = pd.Series(v)
      df = pd.DataFrame.from_dict(t)
      df = df.style.apply(color_different_red, subset=odds_column['atn']+odds_column['htn'])
      df.to_excel(writer, sheet_name=key, index=False)
    writer.close()

if __name__ == '__main__':
  mlb_odds_column = {
      'atn': [f'客隊贏{i}分賠率變化歷史紀錄' for i in range(1, 7)], 
      'htn': [f'主隊贏{i}分賠率變化歷史紀錄' for i in range(1, 7)]
  }
  mlb_odds_column['atn'].append('客隊贏7+分賠率變化歷史紀錄')
  mlb_odds_column['htn'].append('主隊贏7+分賠率變化歷史紀錄')
  nba_odds_column = {
      'atn': [f'客隊贏{i}-{i+4}分賠率變化歷史紀錄' for i in range(1, 26, 5)], 
      'htn': [f'主隊贏{i}-{i+4}分賠率變化歷史紀錄' for i in range(1, 26, 5)]
  }
  nba_odds_column['atn'].append('客隊贏26+分賠率變化歷史紀錄')
  nba_odds_column['htn'].append('主隊贏26+分賠率變化歷史紀錄')

  mlb_data = dict()
  nba_data = dict()

  while True:
    crawler('mlb', mlb_odds_column, mlb_data)
    crawler('nba', nba_odds_column, nba_data)
    time.sleep(58)