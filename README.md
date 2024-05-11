
# Instalação do Microsoft Office 365 - Licença Official

Passos para instalação do Office 365 de forma official com licença.


## Criar Pastas

Acesse o diretorio "C:" e crie uma pasta com nome "MS Office Setup", ou utilizar o comando abaixo para criar no CMD:

```bash
mkdir C:\MS Office Setup
```
## Configurações

1. Acesse o site da [Ferramenta de Personalização do Office](https://config.office.com/deploymentsettings).
2. Escolha a Arquitetura do seu computador:
3. Selecione o Produto.
4. Em idioma escolha a opção "Corresponder com o Sistema Operacional".
5. Após configurações, Clique em Exportar.
6. Em exportar escolha o "Formatos Office Open XML".
7. Aceite os termos e clique em Exportar.
8. Salve o arquivo na pasta "MS Office Setup", criada Anteriormente.
9. Faça o Download da ferramenta "[Office Deployment Tool](https://www.microsoft.com/en-us/download/details.aspx?id=49117).
10. Salve a ferramenta "Office Deployment Tool" na pasta "MS Office".
11. Execute a ferramenta, onde será gerado 5 novos arquivos. 
## Instalação

```bash
  CD\
  CD C:\MS Office Setup
  Setup.exe /configure Configuração.xml
```
## Ativação 

Execute os comandos no CMD:

```bash
  cd /d %ProgramFiles%\Microsoft Office\Office16

  cd /d %ProgramFiles(x86)%\Microsoft Office\Office16

  for /f %x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"

  cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
  cscript ospp.vbs /unpkey:BTDRB >nul
  cscript ospp.vbs /unpkey:KHGM9 >nul
  cscript ospp.vbs /unpkey:CPQVG >nul
  cscript ospp.vbs /sethst:kms8.msguides.com
  cscript ospp.vbs /setprt:1688
  cscript ospp.vbs /act
```
    
