@echo off
rem 指定したフォルダ配下のvbac.wsfを実行しマクロのVBAをエクスポートする.

rem vbac.wsfの存在するフォルダの指定
set inputPath=ExcelVBA

:main
echo Excelからマクロのエクスポート処理実施しますか？
echo * ******************************
echo * [y] 続行する場合
echo * [n] 終了する場合
echo * ******************************
set /p input="入力："
if defined input set input=%input:"=%
if /i "%input%" == "y" (goto start)
if /i "%input%" == "n" (goto quit)
rem それ以外ならコンソールをクリアし再度質問
cls
goto main

rem 入力が"y"の場合
:start
cd %inputPath%
rem ExcelよりVBAをエクスポート
cscript vbac.wsf decombine

rem 結果表示 %~dp0はカレントパスを表示する
echo エクスポート処理が完了しました.
echo;
tree %~dp0%inputPath% /f
pause
goto quit

rem 入力が"n"の場合
:quit
exit

