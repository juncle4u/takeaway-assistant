# cython: language_level=3
import os
import datetime
import base64
import zlib
import uuid
import time
import tempfile
from Crypto.Cipher import AES
import json
import sqlite3
import ctypes
import ctypes.wintypes
import winreg
import os

def getToken():

    reg = winreg.OpenKey(winreg.HKEY_CURRENT_USER,r"SOFTWARE\Google\Chrome\BLBeacon")
    version = int( winreg.QueryValueEx(reg,"version")[0].split(".")[0] )
    if version < 96:
        input("当前的谷歌Chrome浏览器版本过低，请升级至96版以上, 然后重试")
        reg.Close()
        
    with open("门店.txt", "rb") as f:
        shop_dic = eval( f.read().decode() )
        mt_shop_dic = shop_dic["美团"]
        #ele_shop_dic = shop_dic["饿了么"]
    path = "cookie"    
    if os.path.exists(path):
        with open(path,"rb") as f:            
            token_dic = eval( handleText( f.read().decode(), False ) )
            if isValid( token_dic["acctId"], len(mt_shop_dic) ) == False:
                return False  # 授权提示     
            setting_date = datetime.date.fromisoformat(token_dic["token_date"]) 
            delta = datetime.date.today() - setting_date            
            if delta.days == 0:  # 如果当天抓过则直接返回，不必重复抓     
                return token_dic                

    cmd = os.popen("chrome e.waimai.meituan.com") # 每天更新1次cookie
    if cmd.read().find("不是内部或外部命令") != -1:
        reg = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe")
        path = winreg.QueryValueEx(reg,"path")[0]
        reg.Close()
        os.popen( 'setx path "%path%;{}"'.format(path) )
        
    token, bsid, acctId, region_id,  region_version = "","","","",""         
    item_dict = chrome("e.waimai.meituan.com")
    token = item_dict["token"]                  
    bsid = item_dict["bsid"]                     
    acctId = item_dict["acctId"]
    region_id = item_dict["region_id"]
    region_version = item_dict["region_version"]
       
    # 饿了么
    chainId, ksid = "", ""
    visitId = str(uuid.uuid4()).upper().replace("-", "") + "|" + str( round( time.time() * 1000 ) )
    item_dic = chrome(".ele.me")
    ksid = item_dic["ksid"]
    
    token_dic = {
        "token_date":datetime.date.today().isoformat(),
        "mt_shop_dic":mt_shop_dic,   "acctId":acctId,    "token":token,    "bsid":bsid,
        "region_version":region_version,   "region_id": region_id,
        "cookie":"acctId="+ acctId + ";bsid=" + bsid + ";token=" + token + ";wmPoiId=",         
        "ele_shop_dic":ele_shop_dic,  "ksid":ksid,  "visitId":visitId
        }
    data = handleText(repr(token_dic), True)  
    with open(path,"wb+") as f:        
        f.write(data.encode())  # token保存到文件
    return token_dic

def isValid(acctId,shop_len):

    if shop_len == 1: # 单店永久免费
        return True
    setting_path = "setting.ini"    
    if os.path.exists(setting_path):        
        with open(setting_path, "rb") as f:        
            data = f.read().decode()
            key_list = data.split("|")
            setting_acctId = handleText( key_list[0], False )
            timestamp = int( handleText( key_list[1], False ) )             
            exp = datetime.date.fromtimestamp(timestamp)            
            delta = exp - datetime.date.today()            
            if delta.days < 0 or acctId != setting_acctId : # 两个acctId不一致 或  超过授权日期                
                return False            
            else:                
                return True 
    else: # 无setting,则生成setting,然后判断是否授权        
        with open(setting_path, "wb+") as f: 
            default_exp = 1656518400
            key = handleText(acctId, True) + "|" + handleText(str(default_exp), True)
            f.write(key.encode())
            delta = datetime.date.fromtimestamp( default_exp ) - datetime.date.today() 
            if delta.days < 0 :                
                return False           
            else:                
                return True


def handleText(text, flag):
	
    in_tab = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz123456789"
    out_tab = in_tab[::-1]
    if flag == True:
        trans_tab = str.maketrans(in_tab,out_tab)
        base_zlib_text = base64.b64encode(zlib.compress(text.encode())).decode()
        trans_text = base_zlib_text.translate(trans_tab)
        out_text = "eJ" + trans_text[2:]
        return out_text                
    if flag == False:
        trans_tab = str.maketrans(out_tab,in_tab)
        trans_text = text.translate(trans_tab)
        out_text = "eJ" + trans_text[2:]
        in_text = zlib.decompress(base64.b64decode(out_text.encode())).decode()
        return in_text

def crypt(cipher_text=b'', is_key=False):

    class DataBlob(ctypes.Structure):
        _fields_ = [
            ('cbData', ctypes.wintypes.DWORD),
            ('pbData', ctypes.POINTER(ctypes.c_char))
        ]

    blob_in, blob_entropy, blob_out = map(
        lambda x: DataBlob(len(x), ctypes.create_string_buffer(x)),
        [cipher_text, b'', b'']
    )
    desc = ctypes.c_wchar_p()
    if not ctypes.windll.crypt32.CryptUnprotectData( ctypes.byref(blob_in), ctypes.byref(desc), ctypes.byref(blob_entropy),None, None, 0x01, ctypes.byref( blob_out) ):
         raise RuntimeError('Failed to decrypt the cipher text with DPAPI')
    description = desc.value
    buffer_out = ctypes.create_string_buffer(int(blob_out.cbData))
    ctypes.memmove(buffer_out, blob_out.pbData, blob_out.cbData)
    map(ctypes.windll.kernel32.LocalFree, [desc, blob_out.pbData])
    if is_key:
            return description, buffer_out.raw
    else:
            return description, buffer_out.value   


def chrome(domain_name=""):
            
        for pre_path in  [ 'LOCALAPPDATA', 'APPDATA']:    
                path =   os.getenv(pre_path) + '\\Google\\Chrome\\User Data\\Local State'
                if os.path.exists(path):
                        with open(path,'rb') as f:                    
                            key_file_json = json.load(f)                   
                            key_encry = base64.b64decode(key_file_json['os_crypt']['encrypted_key'])[5:]
                            _, key = crypt(key_encry, is_key=True)
                        break
        cookie_file = os.getenv("localappdata") + r"\Google\Chrome\User Data\Default\Network\Cookies"
        if os.path.exists(cookie_file):
                tmp_cookie_file = tempfile.NamedTemporaryFile(suffix='.sqlite').name
                open(tmp_cookie_file, 'wb').write(open(cookie_file, 'rb').read())
        else:
                print("找不到以下文件:",  cookie_file)

        con = sqlite3.connect(tmp_cookie_file)
        cur = con.cursor()
        cur.execute("SELECT host_key, name,  encrypted_value FROM cookies WHERE host_key like  '%{}%'".format(domain_name))
        item_dict = {}
        for item in cur.fetchall():
            data = item[2]
            cipher = AES.new(key,AES.MODE_GCM, data[3:15])
            item_dict[ item[1] ] = cipher.decrypt(data[15:])[:-16].decode()
        con.close()
        return item_dict
