If /I "%Processor_Architecture%" NEQ "x86" (
%SystemRoot%\SysWoW64\WindowsPowerShell\v1.0\powershell.exe /C %0
goto :eof
)
pushd %~dp0
cscript "C:\Program Files (x86)\Inflectra\Rapise\Engine\SeSExecutor.js" "C:\Users\sysadmin\Desktop\Praveen\Account Creation\Account Creation.sstest" "-eval:g_testSetParams={userName:'', password:''};"
popd
