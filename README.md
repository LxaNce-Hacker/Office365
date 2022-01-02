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
