需要确保计算机已安装以下powershell模块

AzureAD模块 ：
Install-Module -Name AzureAD 
Import-module azuread

ExchangeOnline模块：
Install-Module -Name ExchangeOnlineManagement -RequiredVersion 3.0.0

查询已安装模块：
Get-InstalledModule

首次安装需要设定运用远程安装：
Set-ExecutionPolicy RemoteSigned