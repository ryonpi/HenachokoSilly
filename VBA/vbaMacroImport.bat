@echo off
rem �w�肵���t�H���_�z����vbac.wsf�����s���}�N����VBA���C���|�[�g����.

rem vbac.wsf�̑��݂���t�H���_�̎w��
set inputPath=ExcelVBA

:main
echo Excel����}�N���̃C���|�[�g�������{���܂����H
echo * ******************************
echo * [y] ���s����ꍇ
echo * [n] �I������ꍇ
echo * ******************************
set /p input="���́F"
if defined input set input=%input:"=%
if /i "%input%" == "y" (goto start)
if /i "%input%" == "n" (goto quit)
rem ����ȊO�Ȃ�R���\�[�����N���A���ēx����
cls
goto main

rem ���͂�"y"�̏ꍇ
:start
cd %inputPath%
rem Excel���VBA���C���|�[�g
cscript vbac.wsf combine

rem ���ʕ\�� %~dp0�̓J�����g�p�X��\������
echo �C���|�[�g�������������܂���.
echo;
tree %~dp0%inputPath% /f
pause
goto quit

rem ���͂�"n"�̏ꍇ
:quit
exit