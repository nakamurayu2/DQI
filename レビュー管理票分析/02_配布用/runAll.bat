@echo off

set DIRNAME=%~dp0
pushd %DIRNAME%

echo "��AllReviewAnalysis_Rst.xlsx�̃N���A"
call "exefiles/clearRst.exe"

echo "�����r���[�t�@�C�����W"
call "exefiles/write2existRMS.exe"

echo "�����ʈړ�"
copy "log_write2existRMS.txt" "./config/cfg_AllReviewAnalysis.txt"

echo "�����͎��s"
call "exefiles/AllReviewAnalysis.exe"

popd
