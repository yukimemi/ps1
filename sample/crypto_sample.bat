@echo off

rem �Í���
call crypto.cmd -Enc -Plain "TestPassword1234" -Path ".\secret_sample.enc"

rem ����
for /f "delims=" %%a in ('crypto.cmd -Path .\secret_sample.enc') do set plain=%%a

echo %plain%

pause

