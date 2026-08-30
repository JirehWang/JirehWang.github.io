"""
Bible Audio Auditor (聖經時間軸全自動驗證審計工具)
======================================================
Author: Antigravity Assistant & MultiAgent Global
Description:
  全自動化審計台語有聲聖經時間軸 JSON 檔，提供全方位的品質檢測：
  1. 結構與單調性檢查 (Strict Monotonicity & Duration Integrity)
  2. 語速 CPS (Chars-Per-Second) 異常偵測 (Outlier / Truncation Detection)
  3. 純淨台語漢字字數分析 (自動排除羅馬音標、細拉、題註)
  4. 聲學波形邊界分析 (換氣與靜音品質檢驗，防止切在音節中間)
  5. Praat .TextGrid 格式匯出 (支援可視化聲學微調)
"""

import os
import sys
import re
import json
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
    t = re.sub(r"[，。！？；、：“”「」『』…\(\)（）\s\d]", "", t)
    return t.strip()


class BibleAudioAuditor:
    def __init__(self, json_path: str, check_audio: bool = False, audio_path: str = None):
        self.json_path = json_path
        self.check_audio = check_audio
        self.custom_audio_path = audio_path
        
        with open(json_path, "r", encoding="utf-8") as f:
            self.data = json.load(f)
            
        self.bid = self.data.get("bid")
        self.chap = self.data.get("chap")
        self.title = self.data.get("title", f"書卷 {self.bid} 第{self.chap}章")
        self.total_duration = float(self.data.get("total_duration", 0.0))
        self.verses = self.data.get("verses", [])
        
        self.audio_samples = None
        self.audio_sr = None
        self.issues = []

    def load_audio_if_needed(self):
        """下載或讀取對應章節音訊以進行聲學邊界分析。"""
        if not self.check_audio or self.audio_samples is not None:
            return
            
        import miniaudio
        
        target_wav = self.custom_audio_path
        temp_download = False
        
        if not target_wav or not os.path.exists(target_wav):
            url_audio = f"https://media.fhl.net/Taiwanese/{self.bid}/{self.bid}_{self.chap}.mp3"
            print(f"[*] 下載官方音訊進行聲學波形分析: {url_audio} ...")
            target_wav = f"temp_audit_{self.bid}_{self.chap}.mp3"
            req = urllib.request.Request(url_audio, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req) as resp, open(target_wav, "wb") as f:
                f.write(resp.read())
            temp_download = True

        decoded = miniaudio.decode_file(target_wav)
        self.audio_sr = decoded.sample_rate
        samples = np.frombuffer(decoded.samples, dtype=np.int16).astype(np.float32)
        if decoded.nchannels == 2:
            samples = samples.reshape(-1, 2).mean(axis=1)
        self.audio_samples = samples
        
        if temp_download and os.path.exists(target_wav):
            os.remove(target_wav)

    def audit_all(self):
        """執行全方位審計檢查。"""
        self.issues.clear()
        
        if not self.verses:
            self.issues.append({"type": "FATAL", "sec": 0, "msg": "經節列表為空 (verses list is empty)"})
            return self.generate_report()

        total_verses = len(self.verses)
        prev_end = 0.0

        for i, v in enumerate(self.verses):
            sec = v.get("sec", i + 1)
            st = float(v.get("start", 0.0))
            en = float(v.get("end", 0.0))
            raw_text = v.get("text", "")
            pure_text = get_pure_spoken_text(raw_text)
            dur = round(en - st, 2)
            char_count = len(pure_text)

            # --- 檢查 1: 單調性與時長邏輯 ---
            if st >= en:
                self.issues.append({
                    "type": "ERROR", "sec": sec, "metric": "Time Order",
                    "msg": f"起點大於或等於終點 (start {st:.2f}s >= end {en:.2f}s)"
                })

            if i > 0 and st < prev_end - 0.05:
                self.issues.append({
                    "type": "ERROR", "sec": sec, "metric": "Monotonicity",
                    "msg": f"時間軸倒退或重疊 (start {st:.2f}s < prev_end {prev_end:.2f}s)"
                })

            if dur < 1.0:
                self.issues.append({
                    "type": "ERROR", "sec": sec, "metric": "Duration",
                    "msg": f"經節時長極端過短 ({dur:.2f}s < 1.0s)"
                })
            elif dur > 40.0:
                self.issues.append({
                    "type": "WARN", "sec": sec, "metric": "Duration",
                    "msg": f"經節時長偏長 ({dur:.2f}s > 40.0s)，請確認是否包含多節"
                })

            # --- 檢查 2: 語速 CPS (Chars-Per-Second) 異常偵測 ---
            if dur > 0 and char_count > 0:
                cps = char_count / dur
                if cps > 7.0:
                    self.issues.append({
                        "type": "ERROR", "sec": sec, "metric": "CPS Rate",
                        "msg": f"語速異常過快 ({cps:.2f} 字/秒, {char_count}字/{dur:.1f}s) -> 經節可能被腰斬或漏切"
                    })
                elif cps < 1.2 and dur > 5.0:
                    self.issues.append({
                        "type": "WARN", "sec": sec, "metric": "CPS Rate",
                        "msg": f"語速偏慢或含過長空白 ({cps:.2f} 字/秒, {char_count}字/{dur:.1f}s) -> 請確認是否吃到下一節"
                    })

            # --- 檢查 3: 聲學波形換氣品質檢測 ---
            if self.audio_samples is not None:
                sr = self.audio_sr
                idx_en = int(en * sr)
                win = int(sr * 0.03)
                if idx_en + win < len(self.audio_samples) and idx_en - win >= 0:
                    chunk = self.audio_samples[idx_en - win : idx_en + win]
                    cut_rms = np.sqrt(np.mean(chunk ** 2))
                    if cut_rms > 1200:
                        self.issues.append({
                            "type": "WARN", "sec": sec, "metric": "Acoustic Cut",
                            "msg": f"句尾切點能量偏高 (RMS: {cut_rms:.0f}) -> 可能切在發音音節中，建議微調換氣點"
                        })

            prev_end = en

        return self.generate_report()

    def generate_report(self) -> dict:
        """產出結構化審計報告。"""
        errors = [i for i in self.issues if i["type"] == "ERROR"]
        warns = [i for i in self.issues if i["type"] == "WARN"]
        
        total_verses = len(self.verses)
        error_count = len(errors)
        warn_count = len(warns)
        
        score = max(0, 100 - (error_count * 25) - (warn_count * 4))
        
        report = {
            "title": self.title,
            "bid": self.bid,
            "chap": self.chap,
            "total_verses": total_verses,
            "total_duration": self.total_duration,
            "quality_score": score,
            "status": "PASS" if error_count == 0 else "FAIL",
            "error_count": error_count,
            "warn_count": warn_count,
            "issues": self.issues
        }
        return report

    def print_summary(self, report: dict):
        """格式化終端報告輸出。"""
        print("\n" + "=" * 65)
        print(f"📊 聖經時間軸品質審計報告: {report['title']} ({report['total_verses']} 節)")
        print("=" * 65)
        print(f"• 總時長: {report['total_duration']:.2f} 秒 ({report['total_duration']/60:.2f} 分鐘)")
        print(f"• 品質得分: {report['quality_score']} / 100")
        
        if report["status"] == "PASS":
            status_badge = "✅ PASS (審計通過)"
        else:
            status_badge = "❌ FAIL (發現嚴重錯誤)"
        print(f"• 審計結果: {status_badge} [錯誤: {report['error_count']} 個, 警告: {report['warn_count']} 個]\n")

        if not report["issues"]:
            print("🎉 恭喜！全篇所有經節均通過單調性、語速 CPS、時長完整性等 100% 嚴格驗證！")
        else:
            print("🔍 審計檢測問題明細表:")
            print(f"{'層級':<6} | {'經節':<6} | {'檢驗項目':<14} | {'詳細診斷說明'}")
            print("-" * 65)
            for item in report["issues"]:
                lvl = item["type"]
                sec = f"第 {item['sec']:03d} 節" if item["sec"] > 0 else "全篇"
                metric = item.get("metric", "General")
                msg = item["msg"]
                
                badge = "🔴 ERROR" if lvl == "ERROR" else ("🟡 WARN" if lvl == "WARN" else "⚫ FATAL")
                print(f"{badge:<8} | {sec:<8} | {metric:<14} | {msg}")

        print("=" * 65 + "\n")

    def export_textgrid(self, output_path: str):
        """匯出 Praat .TextGrid 供視覺化波形與頻譜圖分析。"""
        try:
            import textgrid
        except ImportError:
            print("[!] 未安裝 textgrid 套件，請執行 `pip install textgrid` 啟用 TextGrid 匯出。")
            return

        tg = textgrid.TextGrid(minTime=0.0, maxTime=self.total_duration)
        tier = textgrid.IntervalTier(name="Verses", minTime=0.0, maxTime=self.total_duration)

        for v in self.verses:
            sec = v["sec"]
            st = max(0.0, float(v["start"]))
            en = min(self.total_duration, float(v["end"]))
            text = f"[{sec}] {get_pure_spoken_text(v['text'])}"
            if en > st:
                tier.add(st, en, text)

        tg.append(tier)
        tg.write(output_path)
        print(f"[+] Praat TextGrid 已成功匯出至: {output_path}")


def main():
    parser = argparse.ArgumentParser(description="台語聖經時間軸自動審計與質檢工具 (Bible Audio Auditor)")
    parser.add_argument("json_file", help="時間軸 JSON 檔案路徑 (例如 timestamps/19_119.json)")
    parser.add_argument("--check-audio", action="store_true", help="下載/讀取音訊進行聲學波形切點品質審計")
    parser.add_argument("--audio-path", type=str, default=None, help="本地音訊檔案路徑 (可選)")
    parser.add_argument("--export-textgrid", type=str, default=None, help="匯出 Praat .TextGrid 檔案路徑")
    parser.add_argument("--save-report", type=str, default=None, help="儲存 JSON 審計報告檔案路徑")
    
    args = parser.parse_args()

    if not os.path.exists(args.json_file):
        print(f"[!] 找不到檔案: {args.json_file}")
        sys.exit(1)

    auditor = BibleAudioAuditor(
        json_path=args.json_file,
        check_audio=args.check_audio,
        audio_path=args.audio_path
    )
    
    if args.check_audio:
        auditor.load_audio_if_needed()

    report = auditor.audit_all()
    auditor.print_summary(report)

    if args.export_textgrid:
        auditor.export_textgrid(args.export_textgrid)

    if args.save_report:
        with open(args.save_report, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)
        print(f"[+] 審計報告已寫入: {args.save_report}")


if __name__ == "__main__":
    main()
