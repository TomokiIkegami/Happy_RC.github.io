# Pythonのpandasを使って、csvデータからグラフを描画するプログラム

import pandas as pd
import matplotlib.pyplot as plt

# CSVファイルの読み込み
csv_filename = "waveform_sample.csv"
df = pd.read_csv(csv_filename)

# グラフの描画
x = df["Time (s)"]
y = df["Sine Wave"]

plt.plot(x, y, label="Sine Wave")
plt.xlabel("Time (s)")
plt.ylabel("Amplitude")
plt.title("Sine Wave")
plt.legend()
plt.grid()

plt.show()