# By Tool Shift 參數調整說明

## 1. 功能目的

By Tool Shift 用來比較「目前 Weekly 資料」與「該 Tool 的歷史 Baseline」，判斷各 Tool 是否出現明顯偏移。

畫面中的 Upper（P95）、Center（P50）、Lower（P05）是固定的規則 ID，報表仍會顯示 P95／P50／P05。即使把 Weekly 從 P95 改成 P90，規則名稱仍是 Upper（P95），實際計算內容請以最右側的 Formula 為準。

## 2. 啟用方式

1. 開啟「圖表處理設定」。
2. 開啟 `Run Tool P05/P50/P95 Shift`。
3. 設定 Alarm K Threshold 與 Peer Gate Bypass。
4. 在 Nominal、Bigger、Smaller 分頁調整各規則。
5. 按「儲存」。按「取消」不會套用本次修改。

設定會保留在使用者設定檔，下次啟動程式仍會沿用。

## 3. K 門檻

| 參數 | 預設值 | 說明 |
|---|---:|---|
| Alarm K Threshold（K >） | 2.00 | K 必須嚴格大於此值，才可能 Alarm。等於門檻時不 Alarm。 |
| Peer Gate Bypass（K >） | 5.00 | K 嚴格大於此值時，略過 Peer Gate。等於此值時仍需檢查 Peer Gate。 |

完整判斷區間：

```text
K <= Alarm K Threshold
    不 Alarm

Alarm K Threshold < K <= Peer Gate Bypass
    K、偏移方向、Resolution、Own Gate、Peer Gate 全部通過才 Alarm

K > Peer Gate Bypass
    K、偏移方向、Resolution、Own Gate 通過後直接 Alarm
    不再檢查 Peer Gate
```

範例：Alarm K=3、Peer Bypass=6。

| K | 結果 |
|---:|---|
| 2.9 | 不 Alarm |
| 3.0 | 不 Alarm |
| 4.0 | 必須通過 Peer Gate |
| 6.0 | 必須通過 Peer Gate |
| 6.1 | 略過 Peer Gate |

程式會限制 Peer Gate Bypass 不得小於 Alarm K Threshold。若兩者設成相同值，所有大於 Alarm Threshold 的結果都會直接略過 Peer Gate。

## 4. Characteristic 分頁

| 分頁 | 不良偏移方向 | 適用情境 |
|---|---|---|
| Nominal | 過高或過低都可能不良 | 有中心目標值的參數 |
| Bigger | 數值往下偏移才是不良 | Bigger is better |
| Smaller | 數值往上偏移才是不良 | Smaller is better |
| Sigma | 使用 Smaller 的設定 | Sigma 不另外提供分頁 |

每個分頁的設定互相獨立。修改 Nominal 不會影響 Bigger 或 Smaller。

## 5. 表格欄位

| 欄位 | 說明 |
|---|---|
| Rule | 固定規則 ID：Upper（P95）、Center（P50）、Lower（P05）。 |
| Weekly | 從目前 Weekly 資料取得哪一個 percentile。可設定 P0～P100。 |
| K Base | 從該 Tool Baseline 取得哪一個 percentile，並與 Weekly 值相減計算 K。 |
| Own Gate | 是否檢查該 Tool 自己的 Baseline gate。關閉後仍會檢查 K、Resolution 與偏移方向。 |
| Own Low | Own Gate 的低側 Baseline percentile。灰色欄位代表該規則不使用此欄。 |
| Own High | Own Gate 的高側 Baseline percentile。灰色欄位代表該規則不使用此欄。 |
| Peer Low | 其他 Tools 合併 Baseline 的低側 percentile。 |
| Peer High | 其他 Tools 合併 Baseline 的高側 percentile。 |
| Formula | 即時顯示實際公式；確認修改是否符合預期時，以此欄為準。 |

### Weekly 與 K Base

例如：

```text
Weekly=P90、K Base=P80
```

表示以 Weekly P90 減掉該 Tool Baseline P80，再除以原本規則的 K denominator。Upper、Center、Lower 使用的 denominator 公式不會因 percentile 設定而改變。

一般建議先讓 Weekly 與 K Base 使用相同 percentile，例如 P90 對 P90；若刻意使用不同 percentile，Alarm 敏感度可能會有明顯變化。

### Own Gate

Own Gate 是「目前 Weekly」相對於「同一個 Tool 自己的 Baseline」的第二層確認。

Nominal 預設：

- Upper：Weekly P95 必須高於自己的 Baseline P95。
- Center：Weekly P50 必須落在自己的 Baseline P25～P75 之外。
- Lower：Weekly P05 必須低於自己的 Baseline P05。

### Peer Gate

Peer Gate 使用「其他 Tools 的 Baseline」，不包含目前正在判斷的 Tool。

- Upper 使用 Peer High，Weekly 必須高於 Whole Baseline 的指定 percentile。
- Lower 使用 Peer Low，Weekly 必須低於 Whole Baseline 的指定 percentile。
- Center 依偏移方向使用 Peer Low 或 Peer High，必須落在 Peer 區間外。
- Peer Baseline 至少需要 10 個資料點；資料不足且 K 未超過 Bypass 時，不會 Alarm。

## 6. Formula 讀法

範例：

```text
W:P50 vs K-Base:P50; K>2;
Own:outside Base P25-P75;
Peer(K<=5):outside Whole P05-P95
```

意思是：

1. Weekly 使用 P50，Tool Baseline 使用 P50 計算 K。
2. K 必須大於 2。
3. Weekly P50 必須在該 Tool Baseline P25～P75 之外。
4. 當 K 小於或等於 5 時，還必須在其他 Tools 的 Whole Baseline P05～P95 之外。
5. 當 K 大於 5 時，略過第 4 步。

Formula 中的 `P05` 與畫面數字 `P5` 意義相同。

## 7. 常見調整方式

### 希望降低 Alarm 數量

- 提高 Alarm K Threshold，例如 2.00 改成 2.50。
- 提高 Peer Gate Bypass，例如 5.00 改成 6.00，讓更多結果必須通過 Peer Gate。
- Nominal Upper：提高 Own High 或 Peer High。
- Nominal Lower：降低 Own Low 或 Peer Low。
- Nominal Center：擴大 Own Low～Own High 或 Peer Low～Peer High 區間。

### 希望提高敏感度

- 降低 Alarm K Threshold，例如 2.00 改成 1.80。
- 降低 Peer Gate Bypass，讓較多結果直接略過 Peer Gate；這可能明顯增加 Alarm。
- Nominal Upper：降低 Own High 或 Peer High。
- Nominal Lower：提高 Own Low 或 Peer Low。
- Nominal Center：縮小 Own Low～Own High 或 Peer Low～Peer High 區間。

### 只想用 K 與方向判斷

可以關閉 Own Gate；若也希望大部分結果不受 Peer Gate 阻擋，可將 Peer Gate Bypass 設得接近 Alarm K Threshold。這種設定會比較敏感，建議先用歷史資料驗證 False Alarm 數量。

## 8. 建議修改流程

1. 先記錄目前 Alarm 數量與主要觸發規則。
2. 一次只修改一類參數，例如只改 K Threshold 或只改 Weekly percentile。
3. 使用相同歷史期間重新分析，比較修改前後差異。
4. 觀察 Formula，確認實際 percentile 與門檻正確。
5. 確認 False Alarm 與漏報風險後再正式使用。

不建議同時大幅修改 Weekly、K Base、Own Gate、Peer Gate 與 K Threshold，否則很難判斷 Alarm 變化是由哪個參數造成。

## 9. 資料量與固定規則

以下規則不受此畫面參數影響：

- Weekly 至少以 5 點進行計算。
- Weekly 只有 1～4 點時，會從最新 Baseline 借點補到 5 點。
- 借點後 Tool Baseline 必須至少保留 10 點。
- Resolution 條件仍會套用。
- Upper／Center／Lower 的 K denominator 公式維持不變。

## 10. Reset to Current Default

按下 `Reset to Current Default` 會將以下內容全部恢復成目前正式預設：

- Nominal、Bigger、Smaller 的全部 percentile 與 Own Gate 設定。
- Alarm K Threshold：2.00。
- Peer Gate Bypass：5.00。

Reset 後仍需按「儲存」才會正式套用；若按「取消」，原本已儲存的設定不會被更動。
