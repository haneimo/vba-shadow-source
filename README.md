vba-shadow-source
=================

日本語ですみません。あとで書き直します;;


####導入に際して
1. .xlsmファイルをExcelで作成する。
2. セキュリティセンターを開き、「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する(V)」をチェックする。
3. Alt+F11でVBEを開く。
4. プロジェクトにVBASHADOWSOURCE.basをインポートする。
5. VBASHADOWSOURCE モジュールをVBEで開く。
6. initVBASHADOWSOURCE サブルーチンを選択してF5キーを押下して実行する。

####導入すると
* xlsmを開いたときに、"ファイル名_拡張子_src"フォルダからソースコードをインポートする。
* xlsmを保存したときに、"ファイル名_拡張子_src"フォルダへソースコードをエクスポートする。

####なんで作った？
subersionやgitでVBAを管理したかったので作った。
コンフリクトを解決した後にxlsmを開くと、解決済みコードでコーディングできる。
xlsmを開いている時は、VBEのコードが正となる。
xlsmを閉じている時は"ファイル名_拡張子_src"配下のコードが正となる。
