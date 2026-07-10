import os
import yfinance as yf
import pandas as pd
import subprocess

ROLLING_WINDOW_MONTHS = 1.5
ROLLING_WINDOW_DAYS = int(21 * ROLLING_WINDOW_MONTHS)
Z_ENTRY = 2.0
MIN_SIGNALS_TO_SHOW = 15
OUTPUT_FILE = 'daily_signal_sheet.xlsx'

backtest_summary = pd.read_csv('backtest_reference.csv')
qualified = backtest_summary[backtest_summary['n_signals'] >= MIN_SIGNALS_TO_SHOW]

all_tickers = pd.unique(qualified[['stock_a', 'stock_b']].values.ravel())
prices = {}
for t in all_tickers:
    data = yf.download(t, period='4mo', auto_adjust=False, progress=False)['Close'].squeeze()
    if len(data) > 0:
        prices[t] = data

signal_rows = []
for _, row in qualified.iterrows():
    a, b = row['stock_a'], row['stock_b']
    if a not in prices or b not in prices:
        continue
    pair_data = pd.DataFrame({a: prices[a], b: prices[b]}).dropna()
    if len(pair_data) < ROLLING_WINDOW_DAYS + 5:
        continue
    ratio = pair_data[a] / pair_data[b]
    roll_mean = ratio.rolling(ROLLING_WINDOW_DAYS).mean()
    roll_std = ratio.rolling(ROLLING_WINDOW_DAYS).std()
    z = (ratio - roll_mean) / roll_std
    latest_z = z.iloc[-1]

    if abs(latest_z) <= Z_ENTRY:
        continue

    signal_rows.append({
        'sector': row['sector'], 'pair': f'{a}/{b}',
        'as_of_close_date': z.index[-1].date(),
        'current_ratio': round(ratio.iloc[-1], 4),
        'current_z': round(latest_z, 2),
        'signal_for_next_open': 'SHORT_A_LONG_B' if latest_z > 0 else 'LONG_A_SHORT_B',
        'backtest_n_signals': row['n_signals'],
        'backtest_win_rate_%': row['win_rate_%'],
        'backtest_avg_pnl_%': row['avg_pnl_%']
    })

daily_signals = pd.DataFrame(signal_rows)
if len(daily_signals):
    daily_signals = daily_signals.sort_values('current_z', key=abs, ascending=False)

with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
    daily_signals.to_excel(writer, sheet_name='SIGNALS_TODAY', index=False)
    backtest_summary.to_excel(writer, sheet_name='BACKTEST_REFERENCE', index=False)

print(f"Active signals today: {len(daily_signals)}")
