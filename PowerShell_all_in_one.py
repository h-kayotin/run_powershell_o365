"""
run_powershell - 借助subprocess来运行powershell

Author: ahjiang
Date 2023/3/24
"""
import os
from glob import glob
import subprocess as sp
import random


class PowerShell:
    # from scapy
    def __init__(self, coding, ):
        cmd = [self._where('PowerShell.exe'),
               "-NoLogo", "-NonInteractive",  # Do not print headers
               "-Command", "-"]  # Listen commands from stdin
        startupinfo = sp.STARTUPINFO()
        startupinfo.dwFlags |= sp.STARTF_USESHOWWINDOW
        self.popen = sp.Popen(cmd, stdout=sp.PIPE, stdin=sp.PIPE, stderr=sp.STDOUT, startupinfo=startupinfo)
        self.coding = coding

    def __enter__(self):
        return self

    def __exit__(self, a, b, c):
        self.popen.kill()

    def run(self, cmd, timeout=15):
        b_cmd = cmd.encode(encoding=self.coding)
        try:
            b_outs, errs = self.popen.communicate(b_cmd, timeout=timeout)
        except sp.TimeoutExpired:
            self.popen.kill()
            b_outs, errs = self.popen.communicate()
        outs = b_outs.decode(encoding=self.coding)
        return outs, errs

    @staticmethod
    def _where(filename, dirs=None, env="PATH"):
        """Find file in current dir, in deep_lookup cache or in system path"""
        if dirs is None:
            dirs = []
        if not isinstance(dirs, list):
            dirs = [dirs]
        if glob(filename):
            return filename
        paths = [os.curdir] + os.environ[env].split(os.path.pathsep) + dirs
        try:
            return next(os.path.normpath(match)
                        for path in paths
                        for match in glob(os.path.join(path, filename))
                        if match)
        except (StopIteration, RuntimeError):
            raise IOError("File not found: %s" % filename)


def change_password():
    new_password = random_password()
    identy = input("请输入要修改的邮箱完整地址：")
    ps_password = f"""
    #连接AzureAD
    $adpasswd = ConvertTo-SecureString -String '管理员密码' -AsPlainText -Force 
    $AzurePwd = New-Object System.Management.Automation.PSCredential ("管理员账号", $adpasswd)
    Connect-AzureAD -AzureEnvironmentName AzureChinaCloud -Credential $AzurePwd

    #设置密码
    $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
    $PasswordProfile.Password = "{new_password}"
    Set-AzureADUser -ObjectId "{identy}" -PasswordProfile $PasswordProfile
    Write-Host ("密码已重置为：{new_password}" )

    #断开连接AzureAD
    Disconnect-AzureAD -Confirm:$false
        """
    run_powershell(ps_password)


def change_group():
    identy = input("请输入要修改的邮箱完整地址：")
    group_name = input("请输入完整邮件组地址：")
    op_type_num = int(input(f"""
        请选择要进行哪种操作：\n
        1：添加到邮件组
        0：从邮件组移除
    """))
    if op_type_num:
        op_type = "Add"
    else:
        op_type = "Remove"
    ps_group = f"""
    $secpasswd = ConvertTo-SecureString -String '管理员密码' -AsPlainText -Force 
    $o365cred = New-Object System.Management.Automation.PSCredential ("管理员账号", $secpasswd) 
    Connect-ExchangeOnline -Credential $o365cred -ExchangeEnvironmentName O365China
    {op_type}-DistributionGroupMember -Identity "{group_name}" -Member "{identy}" -Confirm:$false
    Write-Host ("成功从邮件组：{group_name}，进行了{op_type}操作，用户：{identy}")
    Disconnect-ExchangeOnline -Confirm:$false
    """
    run_powershell(ps_group)


def change_dep():
    print("请确保input/change_dep.csv文件存在，如有中文，请确保编码是utf-8")
    current_path = os.getcwd().replace('\\', '/')
    ps_dep = f"""
    $adpasswd = ConvertTo-SecureString -String '管理员密码' -AsPlainText -Force 
    $AzurePwd = New-Object System.Management.Automation.PSCredential ("管理员账号", $adpasswd)
    Connect-AzureAD -AzureEnvironmentName AzureChinaCloud -Credential $AzurePwd
    $data = Import-Csv {current_path}/input/change_dep.csv
    foreach ($row in $data)
    {{
    Write-Host "Updating the user :"  $row.'User Username'    " Department "  $row.'Department'  -ForegroundColor Yellow  
    Set-AzureADUser -ObjectId (Get-AzureADUser -ObjectId $row.'User Username').ObjectId -Department $row.'Department'
    Write-Host "Updated." -ForegroundColor Green

    }}
    """
    run_powershell(ps_dep)


def change_title():
    print("请确保input/change_title.csv文件存在，如有中文，请确保编码是utf-8")
    current_path = os.getcwd().replace('\\', '/')
    ps_title = f"""
    $adpasswd = ConvertTo-SecureString -String '管理员密码' -AsPlainText -Force 
    $AzurePwd = New-Object System.Management.Automation.PSCredential ("管理员账号", $adpasswd)
    Connect-AzureAD -AzureEnvironmentName AzureChinaCloud -Credential $AzurePwd
    $data = Import-Csv {current_path}/input/change_title.csv
    foreach ($row in $data)
    {{
    Write-Host "Updating the user :"  $row.'User Username'    " Title "  $row.'Title'  -ForegroundColor Yellow  
    Set-AzureADUser -ObjectId (Get-AzureADUser -ObjectId $row.'User Username').ObjectId -JobTitle $row.'Title'
    Write-Host "Updated." -ForegroundColor Green

    }}
    """
    run_powershell(ps_title)


def change_manager():
    print("请确保input/change_manager.csv文件存在")
    current_path = os.getcwd().replace('\\', '/')
    ps_manager = f"""
    $adpasswd = ConvertTo-SecureString -String '管理员密码' -AsPlainText -Force 
    $AzurePwd = New-Object System.Management.Automation.PSCredential ("管理员账号", $adpasswd)
    Connect-AzureAD -AzureEnvironmentName AzureChinaCloud -Credential $AzurePwd
    $data = Import-Csv {current_path}/input/change_manager.csv
    foreach ($row in $data)
    {{
    Write-Host "Updating the user :"  $row.'User Username'    " Manager "  $row.'Manager'  -ForegroundColor Yellow  
    Set-AzureADUserManager -ObjectId (Get-AzureADUser -ObjectId $row.'User Username').Objectid -RefObjectId (Get-AzureADUser -ObjectId $row.'Manager').ObjectId
    Write-Host "Updated." -ForegroundColor Green

    }}
    """
    run_powershell(ps_manager)


def run_powershell(ps_text):
    with PowerShell('GBK') as ps:
        outs, errs = ps.run(ps_text)
    print('error:', os.linesep, errs)
    print('output:', os.linesep, outs)
    print("请按回车键退出--->")
    sp.getoutput("pause")


def random_password(digits=8):
    """
    生成随机密码
    :param digits:密码位数，默认为8
    :return: 密码字符串
    """
    upper = ["A", "B", "C", "D", "E", "F", "G", "H", "J", "K",
             "M", "N", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]
    lower = []
    for i in range(len(upper)):
        lower.append(upper[i].lower())
    nums = [num for num in range(2, 10)]
    password = ""
    for dig in range(digits):
        if dig == 0:
            password += upper[random.randrange(0, len(upper))]
        elif dig < 4:
            password += lower[random.randrange(0, len(lower))]
        else:
            password += str(nums[random.randrange(0, len(nums))])
    return password


PS_dict = {
    "1": change_password,
    "2": change_group,
    "3": change_dep,
    "4": change_title,
    "5": change_manager
}

if __name__ == '__main__':
    print(f"""
    请选择您要进行哪种操作：\n
    1：重置密码\n
    2：邮件组操作\n
    3：修改部门（批量）\n
    4：修改职位（批量）\n
    5：修改直属领导（批量）\n
    """)
    op_nums = input("请输入：")

    PS_dict[str(op_nums)]()






