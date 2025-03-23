import numpy as np
import pandas as pd

# パラメータ設定
sampling_rate = 1000  # サンプリングレート (Hz)
duration = 1.0  # 波形の長さ (秒)
frequency = 5  # 周波数 (Hz)
amplitude = 1.0  # 振幅

# 時間軸の作成
t = np.linspace(0, duration, int(sampling_rate * duration), endpoint=False)

# 波形データの作成
sin_wave = amplitude * np.sin(2 * np.pi * frequency * t)  # サイン波
square_wave = amplitude * np.sign(np.sin(2 * np.pi * frequency * t))  # 矩形波

# データフレームに変換
df = pd.DataFrame({
    "Time (s)": t,
    "Sine Wave": sin_wave,
    "Square Wave": square_wave
})

# CSVに保存
csv_filename = "waveform_sample.csv"
df.to_csv(csv_filename, index=False, encoding="utf-8")

print(f"CSVファイル '{csv_filename}' を作成しました。")