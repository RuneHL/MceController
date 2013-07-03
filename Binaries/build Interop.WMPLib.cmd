
REM requires .net 3.5 SDK

set tlbimp="C:\Program Files\Microsoft SDKs\Windows\v7.0\Bin\x64\tlbimp"
set out=.\Interop.WMPLib.dll
set projdir=..\
set fw20=C:\Windows\Microsoft.NET\Framework\v2.0.50727
set fw35=C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\v3.5

set ref=/reference:"%projdir%\binaries\Microsoft.MediaCenter.dll" /reference:"%projdir%\binaries\Microsoft.MediaCenter.UI.dll" /reference:%fw20%\mscorlib.dll /reference:"%projdir%\packages\Newtonsoft.Json.5.0.6\lib\net35\Newtonsoft.Json.dll" /reference:%fw20%\System.Configuration.Install.dll /reference:"C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\v3.5\System.Core.dll" /reference:"C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\v3.5\System.Data.DataSetExtensions.dll" /reference:%fw20%\System.Data.dll /reference:%fw20%\System.dll /reference:%fw20%\System.Drawing.dll /reference:%fw20%\System.EnterpriseServices.dll /reference:"C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\v3.0\System.Speech.dll" /reference:%fw20%\System.Web.dll /reference:%fw20%\System.Windows.Forms.dll /reference:%fw20%\System.Xml.dll /reference:"%fw35%\System.Xml.Linq.dll" /reference:"C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\v3.0\WindowsBase.dll" /reference:C:\Windows\assembly\GAC\stdole\7.0.3300.0__b03f5f7f11d50a3a\stdole.dll

set keyfile=/keyfile:"%projdir%\Add-In\keyfile.snk"

%tlbimp% C:\Windows\system32\wmp.dll /namespace:WMPLib /machine:Agnostic /out:%out% /sysarray /transform:DispRet %ref% %keyfile%
pause