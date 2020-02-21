■概要
　対象フォルダからレビュー管理票を集めて分析結果を出力する

□前準備
　config/cfg_write2existRMS.txtを設定する(詳細はコメント参照
　※ドキュメント指摘の場合、TARGET_FOLDERのみ変更すればよい

□実行方法
　バッチファイルを実行

■注意
・以下は消さないでください
　- config
　- template
  - AllReviewAnalysis_Rst.xlsx
・exe, AllReviewAnalysis.pyとAllReviewAnalysis_Rst.xlsxのファイルはカレントの状態で実行してください
・エラーを確認したい場合、エクスプローラ上でShift+右クリックから
  Powershell(コマンドプロンプトでもOK)を開き、コマンドライン上から実行してください。
・①実行時、エクセルオープン・クローズが連続します。
　仮想デスクトップを別の番号に移動すれば、デュアルディスプレイなら片方のディスプレイで作業できるかと思います。

■実行結果
exe実行後、
失敗：　すぐウィンドウが閉じる
成功：　Finish!とメッセージが出る　ただし30秒後には勝手に閉じる
　　　　成功したときだけ、AllReviewAnalysis_Rst.xlsxが更新される

■詳細
　①write2existRMS.exe
　　対象フォルダからレビュー管理票を検索して、以下を実施したファイルを作成
　　・各記録シートの(4,22)セルに数式を埋め込む
　　・レビュー管理シートに上記を参照するセルを追加する
　　・ドキュメント指摘シート追加
　　□コンフィグファイル
　　　cfg_write2existRMS.txt
　②AllReviewAnalysis.exe
　　コンフィグで指定したレビュー管理票のドキュメント指摘シートを合算
　　□コンフィグファイル
　　　cfg_AllReviewAnalysis.txt
