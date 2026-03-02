【Amazon Price Search ツール（上司向け）】

■ 初回だけ実施
1) Excelを閉じる
2) run_init.bat をダブルクリック
3) client_id / client_secret / refresh_token / aws_access_key_id / aws_secret_access_key（必要なら aws_session_token）を入力して保存完了メッセージを確認

■ 毎回の更新手順
1) data\input.xlsx を更新（JANコードは B列、1行目は見出し）
2) Excelを閉じる
3) run_update.bat を実行
4) data\output.xlsx に結果が出力される
   - C列: ASIN
   - D列: 新品 送料込み最安
   - E列: 取得日時（ISO 8601）

■ 注意
- 同じJANが複数行ある場合、API呼び出しは1回だけ行います。
- エラーがあるJANは C/D を空欄のまま継続します（処理全体は止まりません）。
- 詳細ログは logs\run.log を確認してください。
- 認証情報は secrets\lwa_secrets.xml に暗号化保存されます（同じWindowsユーザーのみ利用可）。

■ よくあるエラー
- output.xlsx を保存できない:
  Excelが開いたままの可能性があります。Excelをすべて閉じて再実行してください。
- invalid_grant:
  refresh_token の期限切れ/誤入力の可能性があります。run_init.bat で認証情報を再登録してください。
- 401/403:
  SP-API は LWA トークン単体では認証できません。x-amz-access-token と AWS SigV4 署名（aws_access_key_id / aws_secret_access_key / aws_session_token）を併用できているか、SP-API権限と合わせて確認してください。


■ 実機テスト手順
- 詳細はルートの「TEST_実機手順.txt」を参照してください。
- 初回セットアップ確認 / 最小データ更新確認 / エラー継続確認の順でテストできます。
