【Amazon Price Search ツール（上司向け）】

■ 初回だけ実施
1) Excelを閉じる
2) run_init.bat をダブルクリック
3) client_id / client_secret / refresh_token を入力して保存完了メッセージを確認

■ 毎回の更新手順
1) data\input.xlsx を更新（JANコードは B列、1行目は見出し）
2) Excelを閉じる
3) run_update.bat を実行
4) data\output.xlsx に結果が出力される
   - G列: ASIN
   - H列: 新品 送料込み最安
   - I列: 価格取得日時（ISO 8601、価格が取得できた行のみ）

■ 注意
- 同じJANが複数行ある場合、API呼び出しは1回だけ行います。
- 価格取得はASINごとの単発APIで順次実行するため、安定性を優先した運用です。
- cache/price_cache.json に24時間の永続キャッシュを保持し、再実行時のAPI呼び出しを削減します。
- cache/history/prices_YYYY-MM-DD.jsonl に価格取得成功分の履歴を日次追記します（統計分析向け）。
- 該当なし（NotFound/Validation）は negative cache として保持し、TTL内は再呼び出ししません。
- 一時エラー（RateLimit/Server, Other）は transient_error として扱い、キャッシュに永続化しません（次回再実行で再取得）。
- エラーがあるJANは G/H を空欄のまま継続します（処理全体は止まりません）。
- 詳細ログは logs\run.log を確認してください。
- 認証情報は secrets\lwa_secrets.xml に暗号化保存されます（同じWindowsユーザーのみ利用可）。

■ よくあるエラー
- output.xlsx を保存できない:
  Excelが開いたままの可能性があります。Excelをすべて閉じて再実行してください。
- invalid_grant:
  refresh_token の期限切れ/誤入力の可能性があります。run_init.bat で認証情報を再登録してください。
