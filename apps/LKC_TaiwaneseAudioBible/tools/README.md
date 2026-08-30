# 台語有聲聖經時間軸校正工具

## 安裝

在 Python 3.10+ 環境安裝規格中的依賴：

```bash
pip install faster-whisper miniaudio numpy scipy
```

若要使用原本的 Breeze-ASR-26 後端，另外安裝：

```bash
pip install transformers torch
```

## 執行

預設使用 `faster-whisper` 的 `small` 模型、CPU `int8` 推論，輸出到目前工具旁的 `timestamps/`：

```bash
python breeze_aligner.py --bid 19 --chap 23
```

一次處理多章：

```bash
python breeze_aligner.py --bid 19 --chap 23 24 25
```

使用 Breeze-ASR-26：

```bash
python breeze_aligner.py --backend breeze --bid 19 --chap 23
```

也可以指定本機音檔，避免重複下載：

```bash
python breeze_aligner.py --bid 19 --chap 119 --audio ../ps119.mp3
```

`--strict-tail` 會在任一節找不到尾部三字錨點時停止；預設會保留時間軸並列出警告。所有輸出仍會檢查起點單調遞增、相鄰節點連續、最後一節結束於音檔尾端，以及每節 1.5～35 秒的時長限制。

## 測試

不需要 ASR 模型即可執行純函式測試：

```bash
python -m unittest discover -s . -p 'test_*.py' -v
```
