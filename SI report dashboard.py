import pandas as pd
import os
import io
from openpyxl import load_workbook
from datetime import datetime
import requests
from azure.identity  import ClientSecretCredential

import re
from pathlib import Path
from typing import Optional

def newest_matching_excel(
        folder: str | Path,
        pattern: str = r".*\.xlsx?$",          # 默认匹配所有 Excel 文件
        use_mtime: bool = True                # True 用修改时间，False 用创建时间
) -> Optional[Path]:
    """
    返回 folder 中文件名匹配 pattern 且日期最新的 Excel 文件路径。
    若未找到匹配文件，返回 None。
    """
    folder = Path(folder).expanduser().resolve()
    if not folder.is_dir():
        raise NotADirectoryError(f"{folder} 不是有效目录")

    # 1. 列出所有文件并筛选
    candidates = [
        f for f in folder.iterdir()
        if f.is_file() and re.fullmatch(pattern, f.name, flags=re.I)
    ]
    if not candidates:
        return None

    # 2. 按时间排序并取最新
    key = lambda p: p.stat().st_mtime if use_mtime else p.stat().st_ctime
    newest = max(candidates, key=key)
    return newest
folder_path="C:/Users/liudon3x/OneDrive - Intel Corporation/AT SI - SI Dashboard"
latest_file = newest_matching_excel(folder_path, pattern=r"^SI report for the coming dashboard*\.xlsx$")
if latest_file:
        print("最新文件:", latest_file)

tenant_id = "46c98d88-e344-4ed4-8496-4ed7712e255d"
client_id = "ec9b8d51-97ef-47fc-9331-6ddb1ac4b3a8"
client_secret = "Wc-8Q~-N6iM-fumMuvzdJbeEGAxkpUduEXSx8cBj"
def get_graph_access_token(tenant_id, client_id, client_secret):
    """获取 Microsoft Graph API 的访问令牌"""
    credential = ClientSecretCredential(tenant_id, client_id, client_secret)
    token = credential.get_token("https://graph.microsoft.com/.default")
    return token.token

def read_sharepoint_excel_graph_api(
    sharepoint_site_id: str,
    drive_id: str,
    file_path: str,
    sheet_name: str
) -> pd.DataFrame:
    """
    使用 Microsoft Graph API 读取 SharePoint Online 中的 Excel 文件。

    参数
    ----
    tenant_id       : Azure AD 租户 ID
    client_id       : 应用注册的 Application (client) ID
    client_secret   : 应用注册的 Client Secret
    sharepoint_site_id : SharePoint 站点 ID (可通过 Graph API 获取)
    drive_id        : SharePoint 文档库 Drive ID (可通过 Graph API 获取)
    file_path       : 文件在 SharePoint 中的相对路径，例如 'Shared Documents/data.xlsx'
    sheet_name      : 要读取的工作表名称，默认为 'Sheet1'

    返回
    ----
    pd.DataFrame
    """
    token = get_graph_access_token(tenant_id, client_id, client_secret)
    headers = {"Authorization": f"Bearer {token}"}

    # 构造 Graph API 请求 URL
    url = (
        f"https://graph.microsoft.com/v1.0/sites/{sharepoint_site_id}"
        f"/drives/{drive_id}/root:/{file_path}:/content"
    )

    # 下载文件内容
    response = requests.get(url, headers=headers)
    response.raise_for_status()

    # 使用 pandas 解析 Excel 文件
    excel_bytes = io.BytesIO(response.content)
    df = pd.read_excel(excel_bytes, sheet_name=sheet_name)

    return df

sheetNames=['CD1','CD6','SS1','PG8','PG7','KM1','KM2','KM5','KM8','CR1','CR3']
#sheetNames=['CD1']
# 获取当前登录的用户名
username = os.getlogin()

from collections import namedtuple

country=namedtuple('country',['ID','Name'])
ACCP_Country=[country(ID=1,Name='China')]
ACCP_Country.append(country(ID=2,Name='Costa Rica'))
ACCP_Country.append(country(ID=3,Name='Malaysia'))
ACCP_Country.append(country(ID=4,Name='Vietnam'))
df_Country=pd.DataFrame(ACCP_Country)

city=namedtuple('city',['ID','Name','P_id'])
ACCP_City=[city(ID=1,Name='Chengdu',P_id=1)]
ACCP_City.append(city(ID=2,Name='Costa Rica',P_id=2))
ACCP_City.append(city(ID=3,Name='Kulim',P_id=3))
ACCP_City.append(city(ID=4,Name='Penang',P_id=3))
ACCP_City.append(city(ID=5,Name='Ho Chi Minh',P_id=4))
df_City=pd.DataFrame(ACCP_City)

site=namedtuple('site',['ID','Domain','P_id'])
ACCP_Domain=[site(ID=1,Domain='CD1',P_id=1)]
ACCP_Domain.append(site(ID=2,Domain='CD6',P_id=1))
ACCP_Domain.append(site(ID=3,Domain='SS1',P_id=5))
ACCP_Domain.append(site(ID=4,Domain='PG8',P_id=4))
ACCP_Domain.append(site(ID=5,Domain='PG7',P_id=4))
ACCP_Domain.append(site(ID=6,Domain='KM1',P_id=3))
ACCP_Domain.append(site(ID=7,Domain='KM2',P_id=3))
ACCP_Domain.append(site(ID=8,Domain='KM5',P_id=3))
ACCP_Domain.append(site(ID=9,Domain='KM8',P_id=3))
ACCP_Domain.append(site(ID=10,Domain='CR1',P_id=2))
ACCP_Domain.append(site(ID=11,Domain='CR3',P_id=2))
df_Site=pd.DataFrame(ACCP_Domain)
file_url = "https://intel.sharepoint.com/sites/atsi/Shared%20Documents/AT%20SI%20Automation/Power%20BI%20for%20SI%20Report/SI%20report%20for%20the%20coming%20dashboard.xlsx"
file_path = 'C:/Users/'+username+'/OneDrive - Intel Corporation/AT SI Automation/Power BI for SI Report/SI report for the coming dashboard.xlsx'
system=namedtuple('system',['ID','Name','Facility'])
ACCP_System=[system(ID=0,Name='',Facility='0')]
ACCP_System.remove(system(ID=0,Name='',Facility='0'))
install_DF=namedtuple('install_DF',['P_id','QuarterName','Install_Capacity','DF','Domain','ID'])
ACCP_install_DF=[install_DF(P_id=0,QuarterName='',Install_Capacity=0,DF=0,Domain='CD1',ID=0)]
ACCP_install_DF.remove(install_DF(P_id=0,QuarterName='',Install_Capacity=0,DF=0,Domain='CD1',ID=0))
utility=namedtuple('utility',['P_id','Utility','DataTime','Name','Action_Plan','Domain','Install_ID'])
ACCP_utility=[utility(P_id=0,Utility=0,DataTime=1991/7/20,Name='',Action_Plan='',Domain='CD1',Install_ID=0)]
ACCP_utility.remove(utility(P_id=0,Utility=0,DataTime=1991/7/20,Name='',Action_Plan='',Domain='CD1',Install_ID=0))
autoincrement_ID=1
Installid=1
my_dict = {'Q1':'1-1','Q2':'4-1','Q3':'7-1','Q4':'10-1'}
str_QName=''
for item in sheetNames:
    df=read_sharepoint_excel_graph_api("https://intel.sharepoint.com/","atsi","Shared%20Documents/AT SI Automation/SI report for the coming dashboard.xlsx",item)
    #df = pd.read_excel(file_url, sheet_name=item)
    colCount=df.columns.size
    rowCount=len(df)
    for row in range(1,rowCount,1):
        if df.values[row][0]==0:
            break
        ACCP_System.append(system(ID=autoincrement_ID,Name=df.values[row][0],Facility=item))
        for col in range(1,colCount,1):
            current_data=2024/4/1
            str_Name=''
            if df.columns[col]=='Action Plan':
                break
            if 'Q' in df.columns[col]:
                str_Name=df.columns[col]
                #str_mouth=str_Name[:2]
                #str_year = str_Name[3:5]
                #str_data=my_dict[str_mouth]
                #current_data='20'+str_year+'-'+str_data+' 0:0:0'
                #parsed_time = datetime.strptime(current_data, "%Y-%m-%d %H:%M:%S")
            if  'Capacity' in df.values[0][col] or 'capacity' in df.values[0][col]:
                if str_Name=='':
                    str_Name=df.columns[col-1]
                if '.' in str_Name:
                    str_Name=str_Name[:len(str_Name)-2]
                ACCP_install_DF.append(
                    install_DF(P_id=autoincrement_ID, QuarterName=str_Name, Install_Capacity=df.values[row][col], DF=df.values[row][col+1],Domain=item,ID=Installid))
                Installid=Installid+1
            elif df.values[0][col]=='DF' or df.values[0][col]=='System':
                continue
            else:
                if 'Q' in df.columns[col]:
                    str_QName=df.columns[col]
                if '.' in str_QName:
                    str_QName=str_QName[:len(str_QName)-2]
                str_Name = df.values[0][col]
                str_mouth = str_Name[:2]
                str_year = str_Name[3:5]
                if str_Name[:1]!='Q':
                    str_mouth = str_Name[5:]
                    str_year = str_Name[2:4]
                str_data = my_dict[str_mouth]
                current_data='20'+str_year+'-'+str_data+' 0:0:0'
                parsed_time = datetime.strptime(current_data, "%Y-%m-%d %H:%M:%S")
                #filtered_df=ACCP_install_DF[(ACCP_install_DF['Domain']==item)&(ACCP_install_DF['QuarterName'] ==str_QName)]
                filtered_df = [aa for aa in ACCP_install_DF if aa.Domain == item and aa.QuarterName==str_QName]
                pidforinstall=0
                if len(filtered_df)>0:
                    pidforinstall=filtered_df[0].ID
                ACCP_utility.append(utility(P_id=autoincrement_ID, Utility=df.values[row][col], DataTime=parsed_time,Name=str_QName,Action_Plan=df.values[1][colCount-1],Domain=item,Install_ID=pidforinstall))
        autoincrement_ID=autoincrement_ID+1
df_system=pd.DataFrame(ACCP_System)
df_install_DF=pd.DataFrame(ACCP_install_DF)
df_utility=pd.DataFrame(ACCP_utility)

