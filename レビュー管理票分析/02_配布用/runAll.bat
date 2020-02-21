@echo off

set DIRNAME=%~dp0
pushd %DIRNAME%

echo "■AllReviewAnalysis_Rst.xlsxのクリア"
call "exefiles/clearRst.exe"

echo "■レビューファイル収集"
call "exefiles/write2existRMS.exe"

echo "■結果移動"
copy "log_write2existRMS.txt" "./config/cfg_AllReviewAnalysis.txt"

echo "■分析実行"
call "exefiles/AllReviewAnalysis.exe"

popd
