using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;


namespace DCOMSecurity
{

    unsafe static class regUnit
    {
        #region
        const int OWNER_SECURITY_INFORMATION = 0x00000001;
        const int GROUP_SECURITY_INFORMATION = 0x00000002;
        const int DACL_SECURITY_INFORMATION = 0x00000004;
        const int SACL_SECURITY_INFORMATION = 0x00000008;
        const int LABEL_SECURITY_INFORMATION = 0x00000010;
        const int ATTRIBUTE_SECURITY_INFORMATION = 0x00000020;
        const int SCOPE_SECURITY_INFORMATION = 0x00000040;
        const int PROCESS_TRUST_LABEL_SECURITY_INFORMATION = 0x00000080;
        const int BACKUP_SECURITY_INFORMATION = 0x00010000;
        const int SDDL_REVISION_1 = 1;

        [DllImport("advapi32.dll", EntryPoint = "ConvertSecurityDescriptorToStringSecurityDescriptorW", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern int ConvertSdToStringSd(byte[] sd, uint revision, uint information,
            out IntPtr resultString, ref uint resultStringLength);

        [DllImport("advapi32.dll", EntryPoint = "ConvertStringSecurityDescriptorToSecurityDescriptorW", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern int ConvertStringSdToSd(string stringSd, uint revision, out Byte* resultSd, ref uint resultSdLength);

        #endregion

        /// <summary>
        /// 0：读取正确，     -1：keyName 不存在，     -2：path不存在
        /// </summary>
        /// <param name="hkey"></param>
        /// <param name="path"></param>
        /// <param name="keyName"></param>
        /// <param name="regData"></param>
        /// <returns></returns>
        static public int getRegData(RegistryKey rootReg, string path, string keyName, out object regData)
        {

            try
            {
                RegistryKey RegKey = rootReg.OpenSubKey(path, false);
                regData = RegKey.GetValue(keyName);
                RegKey.Close();
                if (regData != null) {
                    return 0;
                }
                return -1;
            }
            catch (Exception e)
            {
                regData = null;
                return -2;
            }


        }

        static public bool deleteAppIDRegKey(string uuid, string keyName) {
            try
            {
                object regData;
                int flag = getRegData(Registry.ClassesRoot, "AppID\\{" + uuid + "}", keyName, out regData);
                if (flag == 0)
                {
                    RegistryKey delKey = Registry.ClassesRoot.OpenSubKey("AppID\\{" + uuid + "}", true);
                    delKey.DeleteValue(keyName,false);
                    delKey.Close();
                }

                //还有一处注册表位置，不过系统好像会自动关联
                //flag = getRegData(Registry.LocalMachine, @"SOFTWARE\Classes\AppID\{" + uuid + "}", keyName, out regData);
                //if (flag == 0)
                //{
                //    RegistryKey delKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Classes\AppID\{" + uuid + "}", true);
                //    delKey.DeleteValue(keyName);
                //    delKey.Close();
                //}
            }
            catch  {
                return false;
            }
            

            return true;
        }

        static public string getCommSDDL(String keyName)
        {
            object regData;
            String path = "";
            String resultSDDL = "";
            if (keyName.Equals("MachineLaunchRestriction"))
            {
                path = @"SOFTWARE\Policies\Microsoft\Windows NT\DCOM";
                if (getRegData(Registry.LocalMachine, path, keyName, out regData)==0)
                {
                    resultSDDL = (string)regData;
                    if (resultSDDL.Length > 0)
                    {
                        return resultSDDL;
                    }
                }
                
            }

            path = @"SOFTWARE\Microsoft\Ole";

            if (getRegData(Registry.LocalMachine, path, keyName, out regData)==0)
            {
                byte[] regByteData = (byte[])regData;
                IntPtr ByteArray;
                uint l = 0;
                ConvertSdToStringSd(regByteData, 1, DACL_SECURITY_INFORMATION | GROUP_SECURITY_INFORMATION | OWNER_SECURITY_INFORMATION | SACL_SECURITY_INFORMATION,
                    out ByteArray,
                    ref l
                );
                resultSDDL = Marshal.PtrToStringUni(ByteArray);

            }
            return resultSDDL;
        }

        static public string getOpcSDDL(String uuid, String keyName,out int result) {
            object regData;
            String path = @"AppID\"+ "{" + uuid + "}";
            String resultSDDL = "";
            result = getRegData(Registry.ClassesRoot, path, keyName, out regData);
            if (result == 0)
            {
                byte[] regByteData = (byte[])regData;
                IntPtr ByteArray;

                uint l = 0;
                ConvertSdToStringSd(regByteData, 1, DACL_SECURITY_INFORMATION | GROUP_SECURITY_INFORMATION | OWNER_SECURITY_INFORMATION | SACL_SECURITY_INFORMATION,
                    out ByteArray,
                    ref l
                );
                resultSDDL = Marshal.PtrToStringUni(ByteArray);
            }
            //else {
            //    path = @"SOFTWARE\Classes\AppID\{" + uuid + "}";
            //    result = getRegData(Registry.LocalMachine, path, keyName, out regData);
            //    if (result == 0) {
            //        byte[] regByteData = (byte[])regData;
            //        IntPtr ByteArray;
            //        uint l = 0;
            //        ConvertSdToStringSd(regByteData, 1, DACL_SECURITY_INFORMATION | GROUP_SECURITY_INFORMATION | OWNER_SECURITY_INFORMATION | SACL_SECURITY_INFORMATION,
            //            out ByteArray,
            //            ref l
            //        );
            //        resultSDDL = Marshal.PtrToStringUni(ByteArray);
            //    }
            //}
            return resultSDDL==null?"": resultSDDL;
        }

        static public bool setCommSDDL(String stringSd, String keyName) {
            try
            {
                RegistryKey RegKey;
                if (keyName.Equals("MachineLaunchRestriction")) {
                    object regData;
                    if (getRegData(Registry.LocalMachine, @"SOFTWARE\Policies\Microsoft\Windows NT\DCOM", keyName, out regData) == 0) {//使用了组策略进行设置启动激活权限
                        RegKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Policies\Microsoft\Windows NT\DCOM", true);
                        RegKey.SetValue(keyName, stringSd, RegistryValueKind.String);
                        RegKey.Close();
                        return true;
                    }
                    
                }
                uint l = 0;
                Byte* sddlbytes;
                ConvertStringSdToSd(stringSd, 1, out sddlbytes, ref l);
                Byte[] Sddl = new Byte[l];
                for (int i = 0; i < l; i++)
                {
                    Sddl[i] = *(sddlbytes + i);
                }

                RegKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Ole", true);
                RegKey.SetValue(keyName, Sddl, RegistryValueKind.Binary);
                RegKey.Close();
                return true;


            }
            catch (Exception e) {
                return false;
            }
            
        }

        static public bool setCommSetting() {
            try
            {
                object regData;
                int flag = getRegData(Registry.LocalMachine, @"SOFTWARE\Microsoft\Ole", "EnableDCOM", out regData);

                if (flag != -2)
                {
                    RegistryKey RegKey;
                    RegKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Ole", true);
                    RegKey.SetValue("EnableDCOM", "Y", RegistryValueKind.String); 
                    RegKey.SetValue("EnableDCOMHTTP", "N", RegistryValueKind.String);

                    RegKey.SetValue("LegacyAuthenticationLevel", 2, RegistryValueKind.DWord);   //连接
                    RegKey.SetValue("LegacyImpersonationLevel", 2, RegistryValueKind.DWord);    //标识

                    RegKey.Close();

                    RegKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Rpc", true);
                    string[] connect = { "ncacn_ip_tcp" };
                    RegKey.SetValue("DCOM Protocols", connect, RegistryValueKind.MultiString);  //tcp/ip协议
                    RegKey.Close();

                }

                return true;

            }
            catch
            {
                return false;
            }
        }

        static public bool setOpcSDDL(String uuid, String stringSd, String keyName) {
            try
            {
                uint l = 0;
                Byte* sddlbytes;
                ConvertStringSdToSd(stringSd, 1, out sddlbytes, ref l);
                Byte[] Sddl = new Byte[l];
                for (int i = 0; i < l; i++)
                {
                    Sddl[i] = *(sddlbytes + i);
                }


                object regData;
                int flag = getRegData(Registry.ClassesRoot, @"AppID" + @"\{" + uuid + "}", keyName, out regData);

                if (flag != -2)
                {
                    RegistryKey RegKey;
                    RegKey = Registry.ClassesRoot.OpenSubKey(@"AppID" + @"\{" + uuid + "}", true);
                    RegKey.SetValue(keyName, Sddl, RegistryValueKind.Binary);
                    RegKey.Close();
                }

                //flag = getRegData(Registry.LocalMachine, @"SOFTWARE\Classes\AppID\{" + uuid + "}", keyName, out regData);
                //if (flag != -2)
                //{
                //    RegistryKey RegKey;
                //    RegKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Classes\AppID\{" + uuid + "}", true);
                //    RegKey.SetValue(keyName, Sddl, RegistryValueKind.Binary);
                //    RegKey.Close();
                //}
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        static public bool setOpcSetting(String uuid)
        {
            try
            {
                object regData;
                int flag = getRegData(Registry.ClassesRoot, @"AppID" + @"\{" + uuid + "}", "AuthenticationLevel", out regData);

                if (flag != -2)
                {
                    RegistryKey RegKey;
                    RegKey = Registry.ClassesRoot.OpenSubKey(@"AppID" + @"\{" + uuid + "}", true);
                    RegKey.SetValue("ActivateAtStorage", "N", RegistryValueKind.String); //在此计算机上运行应用程序
                    RegKey.SetValue("RunAs", "Interactive User");                
                    RegKey.DeleteValue("RemoteServerName", false);     //用于指定数据存储服务器
                    RegKey.DeleteValue("AuthenticationLevel", false);//身份验证 默认
                    RegKey.DeleteValue("EndPoints", false);
                    //string[] connect = { "ncacn_ip_tcp,0,"};
                    //RegKey.SetValue("EndPoints", connect, RegistryValueKind.MultiString);

                    RegKey.Close();
                }

                return true;

            }
            catch
            {
                return false;
            }
        }


        //static public void test(string path, string key)
        //{
        //    //uint l = 0;
        //    //Byte* sddl;
        //    //string stringSd = "O:BAG:BAD:(A;;CCDCLCSWRP;;;BA)";
        //    //ConvertStringSdToSd(stringSd, 1, out sddl, ref l);


        //    //byte[] regData = (byte[])getRegData(@"SOFTWARE\Microsoft\Ole", "DefaultAccessPermission");

        //    //IntPtr ByteArray;
        //    //String resultSddl;
        //    ////string sdStr;
        //    //ConvertSdToStringSd(regData, 1, DACL_SECURITY_INFORMATION | GROUP_SECURITY_INFORMATION | OWNER_SECURITY_INFORMATION | SACL_SECURITY_INFORMATION, out ByteArray,
        //    //    ref l
        //    //);
        //    //resultSddl = Marshal.PtrToStringUni(ByteArray);

        //}


    }
}
