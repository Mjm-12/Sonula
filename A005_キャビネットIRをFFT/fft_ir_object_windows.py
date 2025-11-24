import numpy as np
import matplotlib.pyplot as plt
import scipy.io.wavfile as wavfile
import os
from matplotlib.ticker import EngFormatter, MultipleLocator


class ImpulseResponsePlotter:
    def __init__(self, original_data, adjusted_data, rates, time_axes, file_names, fft_size=2**18):
        self.original_data = original_data
        self.adjusted_data = adjusted_data
        self.rates = rates
        self.time_axes = time_axes
        self.file_names = file_names
        self.fft_size = fft_size  # FFTサイズ（ゼロパディング後）

    def plot(self, mode='original'):  # modeを追加して波形の種類を選択
        fig, axs = plt.subplots(2, 1, figsize=(15, 8))
        self._plot_waveform(axs[0], mode)
        self._plot_fft(axs[1])
        plt.tight_layout()
        plt.show()

    def _plot_waveform(self, ax, mode):  # 波形プロット（originalまたはadjustedを選択）
        if mode == 'original':
            data_set = self.original_data
            title = 'Impulse Responses (Original)'
        elif mode == 'adjusted':
            data_set = self.adjusted_data
            title = 'Impulse Responses (Adjusted)'
        else:
            raise ValueError("Invalid mode. Choose 'original' or 'adjusted'.")
        
        ax.set_title(title)  # タイトル
        ax.set_xlabel('Time (ms)')  # x軸ラベル
        ax.set_ylabel('Amplitude (Normalized)')  # y軸ラベル
        ax.set_xlim(0, 5)  # x軸の表示範囲の設定
        ax.set_xticks(np.arange(0, 10.1, step=1))  # x軸の目盛り
        ax.set_yticks(np.arange(-1, 1.1, step=0.5))  # y軸の目盛り
        ax.xaxis.set_minor_locator(MultipleLocator(0.1))  # サブグリッド（x軸）
        ax.yaxis.set_minor_locator(MultipleLocator(0.1))  # サブグリッド（y軸）
        ax.grid(True, which='major', linestyle='-')  # メジャーグリッド線
        ax.grid(True, which='minor', linestyle=':')  # マイナーグリッド線
        for i, data in enumerate(data_set):  # 選択されたデータのプロット
            ax.plot(self.time_axes[i], data, label=self.file_names[i])
        ax.legend(bbox_to_anchor=(1.02, 1), loc='upper left', fontsize='small')  # 凡例の設定

    def _plot_fft(self, ax): # インパルス応答のFFTプロット

        # グラフにプロットする周波数範囲
        min_freq = 20  # 最小周波数
        max_freq = 20000  # 最大周波数

        # FFT分解能を計算して表示（最初のファイルのサンプリングレートで計算）
        fft_resolution = self.rates[0] / self.fft_size
        print(f"\n=== FFT解析情報 ===")
        print(f"FFTサイズ: {self.fft_size:,} サンプル")
        print(f"FFT分解能: {fft_resolution:.4f} Hz")
        print(f"=================\n")

        ax.set_title(f'FFT of Impulse Responses (Resolution: {fft_resolution:.4f} Hz)')  # タイトル
        ax.set_xlabel('Frequency (Hz)')  # x軸ラベル
        ax.set_ylabel('Magnitude (dB)')  # y軸ラベル
        ax.set_xscale('log')  # x軸を対数スケールに設定
        ax.set_xticks([min_freq, 50, 100, 500, 1000, 5000, 10000, max_freq])  # 対数スケール用の目盛り
        ax.set_xlim(min_freq, max_freq)  # x軸の表示範囲の設定
        ax.set_ylim(-40, 5)  # y軸の表示範囲の設定
        ax.grid(True, which='both', linestyle='-')  # メジャー＆マイナーグリッド線
        formatter = EngFormatter(unit='', sep='')  # 工学形式のフォーマッタ
        ax.xaxis.set_major_formatter(formatter)  # x軸フォーマットの設定

        # FFTを実行しグラフにプロット
        for i, data in enumerate(self.original_data):
            N_original = len(data)

            # ゼロパディングを適用
            if N_original < self.fft_size:
                # データの末尾にゼロを追加してFFTサイズまで拡張
                padded_data = np.pad(data, (0, self.fft_size - N_original), mode='constant')
            else:
                # データがFFTサイズより大きい場合は切り詰め
                padded_data = data[:self.fft_size]

            # ゼロパディング後のFFT実行
            fft_result = np.fft.fft(padded_data)
            freqs = np.fft.fftfreq(self.fft_size, 1 / self.rates[i])[:self.fft_size // 2]
            fft_magnitude = np.abs(fft_result)[:self.fft_size // 2] # 絶対値を取得

            # 指定周波数範囲内のインデックスを取得
            valid_indices = np.logical_and(freqs >= min_freq, freqs <= max_freq)

            max_in_range = np.max(fft_magnitude[valid_indices])  # 周波数範囲内の最大値で正規化
            fft_magnitude_db = 20 * np.log10(fft_magnitude / max_in_range)  # デシベル変換

            ax.plot(freqs, fft_magnitude_db, label=self.file_names[i])  # データのプロット

            # 各ファイルのFFT情報を出力
            print(f"ファイル: {self.file_names[i]}")
            print(f"  元のサンプル数: {N_original:,}")
            print(f"  サンプリングレート: {self.rates[i]:,} Hz")
            print(f"  FFT分解能: {self.rates[i] / self.fft_size:.4f} Hz")

        ax.legend(bbox_to_anchor=(1.02, 1), loc='upper left', fontsize='small')  # 凡例の設定

# データを保持するリスト
original_data = []
adjusted_data = []
rates = []
time_axes = []
peak_positions = []
file_name_x = []

# スクリプトファイルの場所を基準にした相対パス
script_dir = os.path.dirname(os.path.abspath(__file__))
folder_path = os.path.join(script_dir, "IR")

# フォルダ内のすべてのWAVファイルを処理
for file_name in os.listdir(folder_path):
    if file_name.endswith('.wav'):
        file_path = os.path.join(folder_path, file_name)
        rate, data = wavfile.read(file_path)
        
        if data.ndim > 1:  #ステレオの場合、片方のチャンネルを取得
            data = data[:, 0]
        
        data = data / np.max(np.abs(data))  # データの規格化
        time_axis = np.arange(len(data)) / rate * 1000  #時間軸をmsに
        
        #ピーク位置を合わせるためのデータを取得
        peak_index = np.argmax(data)
        peak_time = time_axis[peak_index]
        
        original_data.append(data)
        rates.append(rate)
        time_axes.append(time_axis)
        peak_positions.append(peak_time)
        file_name_x.append(file_name)

# 基準となる最速のピーク位置を取得
min_peak_time = min(peak_positions)

# 時間軸を調整
for i, data in enumerate(original_data):
    shift_ms = peak_positions[i] - min_peak_time
    shift_samples = int(shift_ms * rates[i] / 1000)
    adjusted = np.roll(data, -shift_samples)
    adjusted_data.append(adjusted)

# 描画
# FFTサイズを変更する場合は fft_size パラメータを指定 (デフォルト: 2^18 = 262144)
# 例: plotter = ImpulseResponsePlotter(original_data, adjusted_data, rates, time_axes, file_name_x, fft_size=2**16)
plotter = ImpulseResponsePlotter(original_data, adjusted_data, rates, time_axes, file_name_x)
plotter.plot(mode='original')  # 'original' または 'adjusted' を指定