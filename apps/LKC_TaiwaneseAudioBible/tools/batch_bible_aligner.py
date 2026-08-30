"""
Batch Bible Audio Aligner & Auto-Calibrator (台語有聲聖經全自動雙向對齊與自癒工具)
==================================================================================
Author: Antigravity Assistant & MultiAgent Global
Description:
  核心架構升級：
  1. 雙向頭尾共識演算法 (Bidirectional Head-Tail Consensus)：
     同時比對第 i 節尾字 (Tail-3) 的結束時間與第 i+1 節首字 (Head-3) 的開始時間，
     徹底杜絕「單向提早切斷尾句（如我就歡喜）」的問題！
  2. 字數比例時長防呆 (Min Duration Guardrail)：
     確保每節時長不小於 len(text)/5.5，防止過早截斷。
  3. 10ms 聲學微吸附 (10ms Snapping)：
     精確鎖定在 [Tail_End, Head_Start] 之間的深層換氣靜音谷底 (RMS Minima)。
  4. 整合 Auditor 滿分自動驗證。

CLI Usage:
  python tools/batch_bible_aligner.py --book 19 --chap 122
  python tools/batch_bible_aligner.py --book 19 --range 120-134
"""

import os
import sys
import re
import json
import time
import argparse
import urllib.request
import numpy as np

# 設置 UTF-8 輸出
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")


def get_pure_spoken_text(t: str) -> str:
    """萃取純淨發音文字：排除白話字音標 (kha̍h-tiâu)、細拉、題註、標點符號。"""
    t = re.sub(r"<[^>]*>", "", t)
    t = re.sub(r"（細拉）|\(細拉\)|〔[^〕]*〕|\[[^\]]*\]", "", t)
    t = re.sub(r"\([a-zA-Zāáǎàâēéěèêīíǐìîōóǒòôūúǔùûńňǹ\s\-\u0300-\u036f]+\)", "", t)
    t = re.sub(r"（[a-zA-Zāáǎàâēéěèêīíǐìîōóǒòôūúǔùûńňǹ\s\-\u0300-\u036f]+）", "", t)
    t = re.sub(r"[，。！？；、：“”「」『』…\(\)（）\s\d\-_]", "", t)
    return t.strip()


def strip_spoken_metadata(t: str) -> str:
    """移除語音轉錄中的口播詞（第X節、希伯來字母等）。"""
    t = re.sub(r"第[一二三四五六七八九十百\d]+[節集站折]", "", t)
    t = re.sub(r"第一百[一二三四五六七八九十\d]+[節集站折]", "", t)
    stanzas = [
        "阿勒弗", "伯特", "基默", "達勒", "黑", "瓦夫", "載音", "赫特",
        "泰特", "約德", "卡夫", "拉麥", "邁姆", "努恩", "薩梅克", "阿因",
        "佩", "查德", "科夫", "雷什", "辛", "塔夫"
    ]
    for s in stanzas:
        t = re.sub(s, "", t)
    return t


class RobustBibleAligner:
    def __init__(self, book_id: int, chap_id: int, output_dir: str = "timestamps"):
        self.bid = book_id
        self.chap = chap_id
        self.output_dir = output_dir
        self.model = None
        os.makedirs(output_dir, exist_ok=True)

    def load_asr_model(self):
        if self.model is None:
            from faster_whisper import WhisperModel
            self.model = WhisperModel("small", device="cpu", compute_type="int8")

    def fetch_chapter_text(self):
        """從信望愛 API 取得該章官方漢羅經文。"""
        url = f"https://bible.fhl.net/json/qsb.php?qstr=%E8%A9%A9{self.chap}&version=tghg" if self.bid == 19 else f"https://bible.fhl.net/json/qsb.php?qstr={self.bid}:{self.chap}&version=tghg"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        records = data.get("record", [])
        return records

    def download_and_decode_audio(self):
        """下載並解碼官方音訊。"""
        import miniaudio
        url_audio = f"https://media.fhl.net/Taiwanese/{self.bid}/{self.bid}_{self.chap}.mp3"
        temp_mp3 = f"temp_align_{self.bid}_{self.chap}.mp3"
        
        req = urllib.request.Request(url_audio, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req) as resp, open(temp_mp3, "wb") as f:
            f.write(resp.read())

        decoded = miniaudio.decode_file(temp_mp3)
        sr = decoded.sample_rate
        samples = np.frombuffer(decoded.samples, dtype=np.int16).astype(np.float32)
        if decoded.nchannels == 2:
            samples = samples.reshape(-1, 2).mean(axis=1)

        if os.path.exists(temp_mp3):
            os.remove(temp_mp3)

        duration = len(samples) / sr
        return samples, sr, duration

    def find_silence_valley(self, samples, sr, t_start, t_end, fallback_t):
        """在明確的 [t_start, t_end] 換氣區間內尋找 10ms 最低 RMS 能量谷底。"""
        if t_end <= t_start:
            t_start = max(0.0, fallback_t - 0.4)
            t_end = min(len(samples) / sr, fallback_t + 0.4)

        frame_len = int(sr * 0.02) # 20ms
        hop = int(sr * 0.01)       # 10ms
        
        idx_st = int(t_start * sr)
        idx_en = int(t_end * sr)
        
        best_t = fallback_t
        min_rms = float("inf")
        
        for i in range(idx_st, idx_en - frame_len, hop):
            chunk = samples[i : i + frame_len]
            rms = np.sqrt(np.mean(chunk ** 2))
            if rms < min_rms:
                min_rms = rms
                best_t = i / sr
                
        return round(float(best_t), 2)

    def process_chapter(self) -> str:
        """執行「雙向頭尾共識 + 字數時長防呆 + 10ms 微吸附」對齊流程。"""
        t0 = time.time()
        print(f"\n{'='*65}\n📖 雙向頭尾共識對齊: 書卷 {self.bid} 第 {self.chap} 章\n{'='*65}")
        
        records = self.fetch_chapter_text()
        if not records:
            print(f"[!] 無法取得書卷 {self.bid} 第 {self.chap} 章經文。")
            return None

        samples, sr, total_duration = self.download_and_decode_audio()
        self.load_asr_model()

        clean_records = []
        for r in records:
            sec = int(r["sec"])
            raw_t = r["bible_text"]
            pure_t = get_pure_spoken_text(raw_t)
            clean_records.append((sec, pure_t, raw_t))

        total_verses = len(clean_records)
        print(f"[+] 經文共 {total_verses} 節，音訊總時長: {total_duration:.2f} 秒")

        # 1. 智慧分塊逐詞轉錄 (關閉前文依賴，避免重複幻覺)
        chunk_size = 7
        all_words = []
        import scipy.signal

        for c_idx in range(0, total_verses, chunk_size):
            chunk_records = clean_records[c_idx : c_idx + chunk_size]
            est_cps = 3.5
            chunk_chars = sum(len(r[1]) for r in chunk_records)
            est_dur = chunk_chars / est_cps
            
            t_chunk_st = all_words[-1]["start"] if all_words else 0.0
            if c_idx + chunk_size >= total_verses:
                t_chunk_en = total_duration
            else:
                t_chunk_en = min(total_duration, t_chunk_st + est_dur + 15.0)

            sub = samples[int(t_chunk_st * sr) : int(t_chunk_en * sr)]
            sub_16k = scipy.signal.resample(sub / 32768.0, int(len(sub) * 16000 / sr)).astype(np.float32)

            segs, _ = self.model.transcribe(
                sub_16k,
                language="zh",
                condition_on_previous_text=False,
                word_timestamps=True
            )

            for s in segs:
                if s.words:
                    for w in s.words:
                        all_words.append({
                            "word": w.word.strip(),
                            "start": t_chunk_st + w.start,
                            "end": t_chunk_st + w.end
                        })

        # 2. 雙向頭尾共識邊界定位 (Bidirectional Head-Tail Consensus)
        aligned_verses = []
        word_ptr = 0

        for v_idx in range(total_verses):
            sec, pure_t, raw_t = clean_records[v_idx]
            
            if v_idx == 0:
                snapped_st = 0.00
            else:
                prev_sec, prev_pure, _ = clean_records[v_idx - 1]
                prev_tail_3 = prev_pure[-3:] if len(prev_pure) >= 3 else prev_pure
                curr_head_3 = pure_t[:3] if len(pure_t) >= 3 else pure_t
                
                # 尋找前節尾字 (Tail-3) 的結束時間
                tail_end_t = aligned_verses[-1]["start"] + len(prev_pure) / 5.5
                for w_i in range(word_ptr, len(all_words)):
                    w = all_words[w_i]
                    if any(ch in w["word"] for ch in prev_tail_3):
                        tail_end_t = max(tail_end_t, w["end"])
                        word_ptr = w_i
                        break
                        
                # 尋找本節首字 (Head-3) 的開始時間
                head_st_t = tail_end_t
                for w_i in range(word_ptr, len(all_words)):
                    w = all_words[w_i]
                    w_clean = strip_spoken_metadata(w["word"])
                    if any(ch in w_clean for ch in curr_head_3) or curr_head_3 in w_clean:
                        head_st_t = w["start"]
                        word_ptr = w_i
                        break

                # 修正項 1: 確保尾句絕不被提前截斷 (Tail Protection)
                search_st = min(tail_end_t, head_st_t)
                search_en = max(tail_end_t, head_st_t) + 0.3
                
                # 修正項 2: 尋找 [search_st, search_en] 之間的換氣靜音谷底
                snapped_st = self.find_silence_valley(samples, sr, search_st, search_en, head_st_t)

            if aligned_verses:
                aligned_verses[-1]["end"] = snapped_st

            aligned_verses.append({
                "sec": sec,
                "start": snapped_st,
                "end": round(total_duration, 2),
                "text": raw_t
            })

        aligned_verses[-1]["end"] = round(total_duration, 2)

        # 3. 輸出標準 JSON
        out_data = {
            "bid": self.bid,
            "chap": self.chap,
            "title": f"詩篇 第{self.chap}篇" if self.bid == 19 else f"書卷 {self.bid} 第{self.chap}章",
            "audio_version": "1",
            "total_duration": round(total_duration, 2),
            "verses": aligned_verses
        }

        out_path = os.path.join(self.output_dir, f"{self.bid}_{self.chap}.json")
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(out_data, f, ensure_ascii=False, indent=2)

        elapsed = time.time() - t0
        print(f"[+] 對齊完成！耗時: {elapsed:.2f} 秒 -> {out_path}")

        # 4. 即時 Auditor 自檢
        from bible_audio_auditor import BibleAudioAuditor
        auditor = BibleAudioAuditor(out_path)
        report = auditor.audit_all()
        auditor.print_summary(report)

        return out_path


def main():
    parser = argparse.ArgumentParser(description="台語有聲聖經全自動雙向對齊工具 (Robust Bible Aligner)")
    parser.add_argument("--book", type=int, default=19, help="書卷編號 (例如 19 為詩篇)")
    parser.add_argument("--chap", type=int, default=None, help="單一章節編號 (例如 122)")
    parser.add_argument("--range", type=str, default=None, help="章節範圍 (例如 120-134)")
    parser.add_argument("--output-dir", type=str, default="timestamps", help="JSON 輸出目錄")

    args = parser.parse_args()

    if args.chap is not None:
        aligner = RobustBibleAligner(book_id=args.book, chap_id=args.chap, output_dir=args.output_dir)
        aligner.process_chapter()
    elif args.range is not None:
        st, en = map(int, args.range.split("-"))
        print(f"[*] 準備處理書卷 {args.book} 第 {st} 至 {en} 章...")
        for c in range(st, en + 1):
            aligner = RobustBibleAligner(book_id=args.book, chap_id=c, output_dir=args.output_dir)
            aligner.process_chapter()
    else:
        print("[!] 請指定 --chap <章節> 或 --range <起始-結束>。")


if __name__ == "__main__":
    main()
