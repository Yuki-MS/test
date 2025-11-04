import bioread
import os


os.chdir(r"C:\Users\r-kajihara\Desktop\test")

data = bioread.read_file("001.acq")

channel = data.channels[0]

print(f"チャンネル名: {channel.name}")
print(f"時間データ: {channel.time_index[:5]}")
print(f"サンプリングレート: {channel.samples_per_second} Hz")
print(f"データポイント数: {len(channel.data)}")
print(f"データ: {channel.data[:5]}")  # 最初の10データを表示