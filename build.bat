
@echo off
set CSC="C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"

echo Compiling InvoiceInspector.exe...
%CSC% /target:winexe /out:InvoiceInspector.exe /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Web.Extensions.dll /r:Microsoft.CSharp.dll ui\framework\Program.cs ui\framework\MainForm.cs ui\framework\PythonBridge.cs

if %ERRORLEVEL% EQU 0 (
    echo Build Success! Running...
    start InvoiceInspector.exe
) else (
    echo Build Failed!
    pause
)
