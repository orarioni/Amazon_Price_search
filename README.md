# Amazon Price Search

Amazon Selling Partner API（SP-API）を使って、`data/input.xlsx` にある JAN から **ASIN** と **新品送料込み最安価格** を取得し、`data/output.xlsx` に出力する Windows 向けツールです。  
日次運用を前提に、認証情報の安全な保存・再試行・キャッシュ・履歴保存を組み込んでいます。

## できること

- Excel（B列のJAN）を読み取り、JANごとに Amazon カタログ検索を実行
- JAN から ASIN を解決（見つからない場合は空欄）
- ASINごとに新品オファーを取得し、**送料込み最安**を算出
- 結果を `output.xlsx` の以下列に出力
  - G列: ASIN
  - H列: 新品送料込み最安
  - I列: 価格取得日時（ISO 8601）
- 同一JANの重複をまとめて処理し、APIコール数を削減
- 永続キャッシュ（`cache/price_cache.json`）で24時間以内の再取得を抑制
- 価格履歴を日次追記（`cache/history/prices_YYYY-MM-DD.jsonl`）
- レート制限・一時障害時のリトライ（Retry-After優先 + 指数バックオフ + ジッター）

---

## 必要な前提

- **OS**: Windows 11（Windows 10 でも概ね動作想定）
- **Excel**: Microsoft Excel（COM 自動化を利用）
- **PowerShell**: PowerShell 7 (pwsh) 推奨、または Windows PowerShell 5.1 以上
  - PS 7+ では REST API 呼び出しが標準化され、ヘッダー抽出が効率的
  - PS 5.1 でも完全互換（Invoke-WebRequest -UseBasicParsing 付与）
- **ネットワーク**: Amazon API へ HTTPS 接続できること
- **認証情報**: Amazon SP-API 用の以下3点
  - `client_id`
  - `client_secret`
  - `refresh_token`

> 注意: 本ツールは PowerShell + Excel COM 前提のため、macOS / Linux ではそのまま動きません。

---

## 初回セットアップ（1回だけ）

1. Excel をすべて閉じる
2. リポジトリ直下で `run_init.bat` を実行（ダブルクリック可）
3. プロンプトに従って以下を入力
   - `client_id`
   - `client_secret`
   - `refresh_token`
4. `secrets/lwa_secrets.xml` 作成メッセージを確認

`run_init.bat` は内部で `scripts/00_init_secrets.ps1` を呼び出し、認証情報を保存します。

---

## 実行手順（毎回）

1. `data/input.xlsx` を更新（JANは B列、1行目は見出し）
2. Excel を閉じる
3. `run_update.bat` を実行
4. `data/output.xlsx` を確認
   - 実行中はコンソールに `進捗: 現在件数 / 入力件数` を表示します（約10行ごと）。

`run_update.bat` は内部で `scripts/10_update_excel.ps1` を実行します。

---

## 入出力仕様（Excel）

### 入力: `data/input.xlsx`

- 1枚目のシートを処理対象
- 1行目はヘッダ
- **B列**に JAN を入力（2行目以降）

### 出力: `data/output.xlsx`

入力を元に保存し、次の列を更新します。

- **G列（7列目）**: ASIN
- **H列（8列目）**: 新品送料込み最安価格
- **I列（9列目）**: 価格取得日時（ISO 8601）

補足:

- JAN が空欄の行は G/H/I を空欄化
- 該当商品なし（NotFound/Validation）は G/H/I を空欄化
- 一時エラー（RateLimit/Server など）は当該行のみ空欄で継続（全体停止しない）


### 「最安」の定義（運用ルール）

- 対象コンディション: **New**（`ItemCondition=New`）
- 価格計算: `LandedPrice` を優先。無い場合は `ListingPrice + Shipping`
- Prime可否・出荷元/販売元・ポイント還元は現状の最安判定には含めない
- 上記ルールを変える場合は、社内運用ルールとして事前に合意してから設定/実装を変更してください

---


## SP-API ヘッダ運用（公式推奨に寄せた方針）

本ツールの SP-API 呼び出しは、次のヘッダを毎回付与します。

- `x-amz-access-token`（LWAアクセストークン）
- `x-amz-date`（UTC時刻、`yyyyMMddTHHmmssZ`）
- `User-Agent`（必須）

補足:

- `Authorization: Bearer ...` は付与していません（必須要件ではないため）。
- 将来、PII を扱う restricted operations に拡張する場合は、LWAアクセストークンではなく **RDT（Restricted Data Token）** の利用が必要です。

---

## キャッシュと履歴

### 永続キャッシュ

- ファイル: `cache/price_cache.json`
- アクセストークンキャッシュ: `cache/access_token.json`（有効期限内は再利用）
- JAN→ASIN キャッシュTTL: 7日（`JanAsinCacheTtlHours`）
- ASIN→Offers キャッシュTTL: 24時間（`OfferCacheTtlHours`）
- negative cache TTL: 12時間（`NegativeCacheTtlHours`）
- SP-API応答デバッグ: `DebugSpApiResponse=$true` で要点ログ（`status / errors / request.uri / payload.ASIN / Offers.Count`）をターミナルとログへ出力
- 応答全文の出力上限: `DebugSpApiResponseMaxChars`（既定 4000 文字、機微情報はマスク）
- `ok` / `not_found` はキャッシュ保持
- `transient_error` は永続化しない（次回再取得）

### 価格履歴

- ファイル: `cache/history/prices_YYYY-MM-DD.jsonl`
- 価格取得成功分のみ追記
- 分析用途で時系列比較に利用可能

---

## ログ

- 実行ログ: `logs/run.log`
- 実行メトリクス(JSONL): `logs/metrics.jsonl`
- API失敗時の分類（NotFound/Validation, RateLimit/Server, Other）や件数統計を出力
- 最終サマリに `一時エラー件数` と `未解決件数` を出力
  - `一時エラー件数`: `RateLimit/Server` と `Other` のうち再試行対象になり得る一時的失敗件数（当該行は空欄で継続）
  - `未解決件数`: `NotFound/Validation + RateLimit/Server + Other` の合計件数（最終的に解決できず空欄出力になった件数）
- レート制限関連では `x-amzn-RateLimit-Limit` / `Retry-After` をログに残し、運用時の上限把握に利用
- 終了時に `input_rows / unique_asin / pricing_calls / pricing_reduction_pct / http429_count / avg_wait_sec` を出力

トラブル時はまず `logs/run.log` を確認してください。

---

## よくあるエラーと対処

### 1) `output.xlsx` を保存できない

**症状**: 保存時に失敗メッセージが出る  
**原因**: Excelで対象ファイルが開いたまま  
**対処**: Excelをすべて閉じて再実行

### 2) `invalid_grant`

**症状**: アクセストークン取得に失敗  
**原因**: `refresh_token` の期限切れ・誤入力  
**対処**: `run_init.bat` を再実行し認証情報を再登録

### 3) ASIN が見つからない（結果が空欄）

**症状**: G/H/I が空欄のまま  
**原因**:
- JANが誤っている
- 対象マーケットプレイスで商品が見つからない
- バリエーション/識別子差異で一致しない

**対処**:
- JANの桁・値を再確認
- 本ツールは Catalog API で **JAN検索→未解決分のみEANフォールバック検索** を実施済み（`identifiersType=JAN` 後に `identifiersType=EAN`）
- 必要に応じて Amazon 側で商品存在を確認
- `logs/run.log` の該当JANの分類ログを確認

### 4) PowerShell 実行ポリシーに関する警告

**症状**: スクリプト実行がブロックされる  
**対処**: 本ツールは `-ExecutionPolicy Bypass` 付きで `.bat` から起動するため、基本は `run_init.bat` / `run_update.bat` 経由で実行する

### 5) 日本語の文字化け

**症状**: ログやREADME表示が崩れる  
**対処**:
- エディタの文字コードを UTF-8 に設定
- PowerShell / ターミナルのフォント・文字コード設定を見直す

---

## セキュリティ

- 認証情報は `secrets/lwa_secrets.xml` に保存
- `client_secret` / `refresh_token` は **SecureString + DPAPI** で暗号化
- 復号できるのは原則として **同じWindowsユーザー** のみ
- `secrets/` 配下のファイルは社内共有・メール添付しない
- 退職/端末移行時はトークン失効と再発行を推奨

---

## ディレクトリ構成（主要）

- `run_init.bat` : 初期認証情報登録の起動
- `run_update.bat` : 更新処理の起動
- `config.psd1` : 環境設定（マーケットプレイス、TTL、各パス等）
- `scripts/00_init_secrets.ps1` : 薄い実行スクリプト（入力→保存）
- `scripts/10_update_excel.ps1` : 薄い実行スクリプト（設定読込→実行）
- `scripts/lib/AmazonPriceLib.psm1` : 共通ライブラリ（認証、SP-API、リトライ、キャッシュ、Excel出力）
- `data/` : 入出力Excel配置場所（`input.xlsx` / `output.xlsx`）
- `cache/` : 永続キャッシュ・履歴・アクセストークンキャッシュ
- `logs/` : 実行ログ
- `secrets/` : 認証情報（実行時生成）

---

## 運用メモ

- API呼び出し回数を抑えるため、同一JANは自動で集約処理されます。
- Pricing呼び出しは基本直列で、最小間隔を保ちながら動的スロットリングします。
- 429/503 が増える場合は、バッチサイズを自動で 20→10→5 と段階縮小して成功率を優先します。
- 件数が多い場合は実行時間が伸びるため、更新バッチを分けると切り分けしやすくなります。
- 定期運用前に、少件数データで `output.xlsx` と `logs/run.log` の内容を一度確認することを推奨します。
- 実行後は `logs/run.log` の終了メトリクスで `unique_asin` と `pricing_calls` がどちらも 0 以外であることを確認してください。
