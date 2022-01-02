# **LxaNce👸🤴**
# Office365
# 1.Manual method
Time needed: 1 minute.

Manually activate your Office 365 using KMS client key.

Open command prompt as admin.
First, you need to open command prompt with admin rights, then follow the instruction below step by step. Just copy/paste the commands and do not forget to hit Enter in order to execute them.
open cmd with admin rights - Legal way to use Office 365 totally FREE, without paying a dime

Navigate to your Office folder.
If you install your Office in the ProgramFiles folder, the path will be “%ProgramFiles%\Microsoft Office\Office16” or “%ProgramFiles(x86)%\Microsoft Office\Office16”. It depends on the architecture of the Windows OS you are using. If you are not sure of this issue, don’t worry, just run both of the commands above. One of them will be not executed and an error message will be printed on the screen.
```cd /d %ProgramFiles%\Microsoft Office\Office16```
```cd /d %ProgramFiles(x86)%\Microsoft Office\Office16```

activate office 2019 using manual method kms 1 - Legal way to use Office 365 totally FREE, without paying a dime

Convert your Office license to volume one if possible.
If your Office is got from Microsoft, this step is required. On the contrary, if you install Office from a Volume ISO file, this is optional so just skip it if you want.
```for /f %x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"```

activate office 2016 manually method 2 - Legal way to use Office 365 totally FREE, without paying a dime

Use KMS client key to activate your Office
Make sure your PC is connected to the internet, then run the following command.
```cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99```
```cscript ospp.vbs /unpkey:BTDRB >nul```
```cscript ospp.vbs /unpkey:KHGM9 >nul```
```cscript ospp.vbs /unpkey:CPQVG >nul```
```cscript ospp.vbs /sethst:s8.uk.to```
```cscript ospp.vbs /setprt:1688```
```cscript ospp.vbs /act```


If you see the error 0xC004F074, it means that your internet connection is unstable or the server is busy. Please make sure your device is online and try the command “act” again until you succeed.

Here is all the text you will get in the command prompt window.
>C:\Windows\system32>cd /d %ProgramFiles%\Microsoft Office\Office16
>C:\Program Files\Microsoft Office\Office16>cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
The system cannot find the path specified.
>C:\Program Files\Microsoft Office\Office16>for /f %x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"
>C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ppd.xrm-ms"
Microsoft (R) Windows Script Host Version 5.812
Copyright (C) Microsoft Corporation. All rights reserved.
>---Processing--------------------------
>Installing Office license: ..\root\licenses16\proplusvl_kms_client-ppd.xrm-ms
>Office license installed successfully.
>---Exiting-----------------------------
>C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ul-oob.xrm-ms"
Microsoft (R) Windows Script Host Version 5.812
Copyright (C) Microsoft Corporation. All rights reserved.
>---Processing--------------------------
>Installing Office license: ..\root\licenses16\proplusvl_kms_client-ul-oob.xrm-ms
>Office license installed successfully.
>---Exiting-----------------------------
>C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ul.xrm-ms"
>Microsoft (R) Windows Script Host Version 5.812
>Copyright (C) Microsoft Corporation. All rights reserved.
>---Processing--------------------------
>Installing Office license: ..\root\licenses16\proplusvl_kms_client-ul.xrm-ms
>Office license installed successfully.
>---Exiting-----------------------------
>C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
>Microsoft (R) Windows Script Host Version 5.812
>Copyright (C) Microsoft Corporation. All rights reserved.
>---Processing--------------------------

>---Exiting-----------------------------
>C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /unpkey:BTDRB >nul
>C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /unpkey:KHGM9 >nul
>C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /unpkey:CPQVG >nul
>C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /sethst:s8.uk.to
>Microsoft (R) Windows Script Host Version 5.812
>Copyright (C) Microsoft Corporation. All rights reserved.
>---Processing--------------------------
>Successfully applied setting.
>---Exiting-----------------------------
>C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /setprt:1688
>Microsoft (R) Windows Script Host Version 5.812
>Copyright (C) Microsoft Corporation. All rights reserved.
>---Processing--------------------------
>Successfully applied setting.
>---Exiting-----------------------------
>C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /act
>Microsoft (R) Windows Script Host Version 5.812
>Copyright (C) Microsoft Corporation. All rights reserved.
>---Processing--------------------------
>Installed product key detected - attempting to activate the following product:
>SKU ID: d450596f-894d-49e0-966a-fd39ed4c4c64
>LICENSE NAME: Office 16, Office16ProPlusVL_KMS_Client edition
>LICENSE DESCRIPTION: Office 16, VOLUME_KMSCLIENT channel
>Last 5 characters of installed product key: WFG99


>---Exiting-----------------------------
>C:\Program Files\Microsoft Office\Office16>


# Step 1: 
>Copy the code below into a new text document.

[@echo off
title Activate Office 365 ProPlus for FREE - MSGuides.com&cls&echo =====================================================================================&echo #Project: Activating Microsoft software products for FREE without additional software&echo =====================================================================================&echo.&echo #Supported products: Office 365 ProPlus (x86-x64)&echo.&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&(for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&echo.&echo ============================================================================&echo Activating your Office...&cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:WFG99 >nul&cscript //nologo ospp.vbs /unpkey:DRTFM >nul&cscript //nologo ospp.vbs /unpkey:BTDRB >nul&set i=1&cscript //nologo ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 >nul||cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul||goto notsupported
:skms
if %i% GTR 10 goto busy
if %i% EQU 1 set KMS=kms7.MSGuides.com
if %i% EQU 2 set KMS=s8.uk.to
if %i% EQU 3 set KMS=s9.us.to
if %i% GTR 3 goto ato
cscript //nologo ospp.vbs /sethst:%KMS% >nul
:ato
echo ============================================================================&echo.&echo.&cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo ============================================================================&echo.&echo #My official blog: MSGuides.com&echo.&echo #How it works: bit.ly/kms-server&echo.&echo #Please feel free to contact me at msguides.com@gmail.com if you have any questions or concerns.&echo.&echo #Please consider supporting this project: donate.msguides.com&echo #Your support is helping me keep my servers running 24/7!&echo.&echo ============================================================================&choice /n /c YN /m "Would you like to visit my blog [Y,N]?" & if errorlevel 2 exit) || (echo The connection to my KMS server failed! Trying to connect to another one... & echo Please wait... & echo. & echo. & set /a i+=1 & goto skms)
explorer "http://MSGuides.com"&goto halt
:notsupported
echo ============================================================================&echo.&echo Sorry, your version is not supported.&echo.&goto halt
:busy
echo ============================================================================&echo.&echo Sorry, the server is busy and can't respond to your request. Please try again.&echo.
:halt
pause >nul]


# Step 2: 
>Save it as a batch file. (eg. office365.cmd).

Save the text file
Set name of the batch file
# Step 3: 
>Run the batch file with admin rights. (important!).

Run the batch file as admin
Done! Your Office is activated successfully.

Successfully activate Office 365
# Note:

I only test this method with Office 365 ProPlus version. I am not sure it will work with the others.
# Step 3 
>is flagged “important” because the UAC system will stop this process if you don’t do it.


# Renew Microsoft Office license
**Step 1:** Open command prompt as administrator.

**Step 2:** Copy and run the command below. Note: “Office16” is code name of Office 2016. If you are using Office 2013/2010, just replace it with “Office15” and “Office14”.


 
```cscript "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" /act```
If you see an error, try this command.

```cscript "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" /act```
Renew Office trial license
Done!

# Renew Microsoft Windows license
**Step 1:** Open command prompt as administrator.

**Step 2:** Execute this command.


 
```cscript slmgr.vbs /ato```
Your license is renewed successfully!
