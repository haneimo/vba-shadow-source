vba-shadow-source
=================

.xlsmファイルをExcelで作成する。
セキュリティセンターを開き、「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する(V)」をチェックする。
Alt+F11でVBEを開く。
ロジェクトにVBASHADOWSOURCE.basをインポートする。
VBASHADOWSOURCE モジュールをVBEで開く。
initVBASHADOWSOURCE サブルーチンを選択してF5キーを押下して実行する。


xlsmを開いたときに、"ファイル名_拡張子_src"フォルダからソースコードをインポートする。
xlsmを保存したときに、"ファイル名_拡張子_src"フォルダへソースコードをエクスポートする。


subersionやgitでコンフリクトを解決した後にxlsmを開くと、
解決済みコードでコーディングできる。
xlsmを開いている時は、VBEのコードが正となる。
xlsmを閉じている時は"ファイル名_拡張子_src"配下のコードが正となる。