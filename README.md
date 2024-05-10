SITES:

Office Customization Tool
https://config.office.com/deploymentsettings

Office Deployment Tool
https://www.microsoft.com/en-us/download/details.aspx?id=49117


INSTALAÇÃO
Comandos - CMD

1° CD\ (ENTER)
2° CD C:\MS Office Setup (ENTER)
3° Setup.exe /configure Configuração.xml (ENTER)



ATIVAÇÃO
Comandos CMD

cd /d %ProgramFiles%\Microsoft Office\Office16 (ENTER)

cd /d %ProgramFiles(x86)%\Microsoft Office\Office16 (ENTER)


for /f %x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x" (ENTER)



cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
cscript ospp.vbs /unpkey:BTDRB >nul
cscript ospp.vbs /unpkey:KHGM9 >nul
cscript ospp.vbs /unpkey:CPQVG >nul
cscript ospp.vbs /sethst:kms8.msguides.com
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /act 

após processar (ENTER)
<Product activation successful>
