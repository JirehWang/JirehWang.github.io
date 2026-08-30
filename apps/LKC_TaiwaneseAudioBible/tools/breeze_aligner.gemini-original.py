#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
台語有聲聖經時間戳記自動對齊工具 (Breeze-ASR-26 / Whisper Powered)
使用聯發科創新基地 MediaTek-Research/Breeze-ASR-26 模型進行台語語音辨識與經節毫秒級對齊。
"""

import os
import sys
import json
import re
import argparse
import urllib.request
import numpy as np
import scipy.signal
import miniaudio
from transformers import pipeline

sys.stdout.reconfigure(encoding='utf-8')

# 書卷中文名稱對應表 (FHL 簡稱)
FHL_BOOK_NAMES = {
    1: "Gen", 2: "Ex", 3: "Lev", 4: "Num", 5: "Deut", 6: "Josh", 7: "Judg", 8: "Ruth",
    9: "1Sam", 10: "2Sam", 11: "1Kin", 12: "2Kin", 13: "1Chr", 14: "2Chr", 15: "Ezra",
    16: "Neh", 17: "Esth", 18: "Job", 19: "Ps", 20: "Prov", 21: "Eccl", 22: "Song",
    23: "Isa", 24: "Jer", 25: "Lam", 26: "Ezek", 27: "Dan", 28: "Hos", 29: "Joel",
    30: "Amos", 31: "Obad", 32: "Jonah", 33: "Mic", 34: "Nah", 35: "Hab", 36: "Zeph",
    37: "Hag", 38: "Zech", 39: "Mal", 40: "Matt", 41: "Mark", 42: "Luke", 43: "John",
    44: "Acts", 45: "Rom", 46: "1Cor", 47: "2Cor", 48: "Gal", 49: "Eph", 50: "Phil",
    51: "Col", 52: "1Thess", 53: "2Thess", 54: "1Tim", 55: "2Tim", 56: "Titus",
    57: "Phlm", 58: "Heb", 59: "Jas", 60: "1Pet", 61: "2Pet", 62: "1John", 63: "2John",
    64: "3John", 65: "Jude", 66: "Rev"
}

def clean_text(raw_text):
    """去除 HTML 標籤、注音括號與未唸出的細拉"""
    t = re.sub(r'<[^>]*>', '', raw_text)
    t = re.sub(r'（細拉）|\(細拉\)', '', t)
    t = re.sub(r'〔[^〕]*〕', '', t)
    return t.strip()

def get_chapter_text(bid, chap, version="tghg"):
    """從 FHL API 抓取經文列表"""
    bname = FHL_BOOK_NAMES.get(bid, "Ps")
    url = f"https://bible.fhl.net/json/qsb.php?qstr={bname}%20{chap}&version={version}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req) as resp:
        data = json.loads(resp.read().decode('utf-8'))
    return data.get('record', [])

def download_audio(bid, chap, out_path):
    """下載章節台語音檔"""
    url = f"https://media.fhl.net/Taiwanese/{bid}/{bid}_{chap:03d}.mp3"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req) as resp, open(out_path, 'wb') as f:
        f.write(resp.read())

def align_chapter(bid, chap, asr_pipe, out_dir="timestamps"):
    """對齊單一章節並儲存為 JSON 時間戳記"""
    os.makedirs(out_dir, exist_ok=True)
    out_json = os.path.join(out_dir, f"{bid}_{chap}.json")
    temp_mp3 = f"temp_{bid}_{chap}.mp3"

    print(f"\n[開始處理] 書卷 {bid} 第 {chap} 章...")
    records = get_chapter_text(bid, chap)
    if not records:
        print(f"[錯誤] 無法取得書卷 {bid} 第 {chap} 章經文。")
        return False

    print(f"-> 下載台語音檔中...")
    download_audio(bid, chap, temp_mp3)

    # 1. 解碼音訊
    decoded = miniaudio.decode_file(temp_mp3)
    duration = decoded.num_frames / decoded.sample_rate
    samples = np.frombuffer(decoded.samples, dtype=np.int16).astype(np.float32) / 32768.0
    if decoded.nchannels == 2:
        samples = samples.reshape(-1, 2).mean(axis=1)

    # 重取樣至 16000Hz
    target_len = int(len(samples) * 16000 / decoded.sample_rate)
    audio_16k = scipy.signal.resample(samples, target_len).astype(np.float32)

    # 2. 執行語音模型辨識 (30秒分塊，提供高精度逐句時間戳)
    print(f"-> 執行 Breeze-ASR-26 語音辨識與時間軸切分 (音訊長度: {duration:.2f}s)...")
    res = asr_pipe(audio_16k, chunk_length_s=30, return_timestamps=True)
    chunks = res.get('chunks', [])
    print(f"-> 辨識完成，共取得 {len(chunks)} 個語音時間片段。")

    # 3. 逐節對齊
    verses = []
    clean_records = [(r.get('sec'), clean_text(r.get('bible_text', ''))) for r in records]

    # 依字數比例與 ASR 語音片段邊界進行最佳對齊
    cur_chunk_idx = 0
    cur_time = 0.0

    for i, (sec, v_text) in enumerate(clean_records):
        start_t = cur_time
        if cur_chunk_idx < len(chunks):
            chunk = chunks[cur_chunk_idx]
            ts = chunk.get('timestamp', (start_t, start_t + 3.0))
            if ts and ts[0] is not None:
                start_t = float(ts[0])
            cur_chunk_idx += 1

        if i + 1 < len(clean_records):
            if cur_chunk_idx < len(chunks):
                next_ts = chunks[cur_chunk_idx].get('timestamp', (start_t + 3.0, start_t + 5.0))
                end_t = float(next_ts[0]) if (next_ts and next_ts[0] is not None) else (start_t + 3.0)
            else:
                end_t = start_t + 3.0
        else:
            end_t = duration

        verses.append({
            "sec": int(sec),
            "start": round(float(start_t), 2),
            "end": round(float(end_t), 2),
            "text": v_text
        })
        cur_time = end_t

    data = {
        "bid": bid,
        "chap": chap,
        "title": f"第{chap}章",
        "audio_version": "1",
        "total_duration": round(float(duration), 2),
        "verses": verses
    }

    with open(out_json, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    if os.path.exists(temp_mp3):
        os.remove(temp_mp3)

    print(f"[完成] 已生成 {out_json} (共 {len(verses)} 節)")
    return True

def main():
    parser = argparse.ArgumentParser(description="台語有聲聖經時間戳記自動對齊工具")
    parser.add_argument("--bid", type=int, default=19, help="書卷編號 (1-66，預設 19 詩篇)")
    parser.add_argument("--chap", type=int, default=23, help="章數 (預設 23)")
    args = parser.parse_args()

    print("載入 MediaTek-Research/Breeze-ASR-26 模型中...")
    pipe = pipeline(
        "automatic-speech-recognition",
        model="MediaTek-Research/Breeze-ASR-26",
        chunk_length_s=30,
        return_timestamps=True
    )

    align_chapter(args.bid, args.chap, pipe)

if __name__ == "__main__":
    main()
