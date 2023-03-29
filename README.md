# run_powershell_o365
将公司常用的几个o365的exchange操作，打包成一个python项目

### 环境搭建

需要确保计算机已安装以下powershell模块

AzureAD模块 ：
```PowerShell
Install-Module -Name AzureAD 
Import-module azuread
```

ExchangeOnline模块：

```PowerShell
Install-Module -Name ExchangeOnlineManagement -RequiredVersion 3.0.0
```

查询已安装模块：
```PowerShell
Get-InstalledModule
```

首次安装需要设定允许远程安装：

```PowerShell
Set-ExecutionPolicy RemoteSigned
```

### How to use

1.请修改代码中的管理员账号和密码为实际账号密码

2.用到以下包，如果没有请pip install一下
```Python
import os
from glob import glob
import subprocess as sp
import random
```

3.直接运行Powershell_all_in_one.py

4.有以下功能，请按要求输入即可
```Python
    请选择您要进行哪种操作：\n
    1：重置密码\n
    2：邮件组操作\n
    3：修改部门（批量）\n
    4：修改职位（批量）\n
    5：修改直属领导（批量）\n
```
5.批量修改功能，需要提前编辑好csv文件，文件格式请参考input文件夹中的各文件
