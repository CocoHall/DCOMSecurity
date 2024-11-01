using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;


namespace DCOMSecurity
{
    class mUser {
        public string name;
        public string sid;
        public int remoteAccess = 0;
        public int remoteLaunch = 0;
        public int remoteActive = 0;
        public int safe = 0;

    }

    class mSID{
        public string name;
        public string sid;
    }
    class commACL
    {
        public string sidOrAclName ;//例如 LA s-1-1-500
        public string displayName ;//例如 Administrator
        public int allowDefaultLocalAccess , denyDefaultLocalAccess ;
        public int allowDefaultRemoteAccess , denyDefaultRemoteAccess ;

        public int allowDefaultLocalLaunch , denyDefaultLocalLaunch ;
        public int allowDefaultRemoteLaunch , denyDefaultRemoteLaunch ;
        public int allowDefaultLocalActive , denyDefaultLocalActive ;
        public int allowDefaultRemoteActive , denyDefaultRemoteActive ;

        public int allowRestrictionLocalAccess , denyRestrictionLocalAccess ;
        public int allowRestrictionRemoteAccess , denyRestrictionRemoteAccess ;

        public int allowRestrictionLocalLaunch , denyRestrictionLocalLaunch ;
        public int allowRestrictionRemoteLaunch , denyRestrictionRemoteLaunch ;
        public int allowRestrictionLocalActive , denyRestrictionLocalActive ;
        public int allowRestrictionRemoteActive , denyRestrictionRemoteActive ;
    }

    class opcACL
    {
        public string uuid="";
        public string opcName="";
        public string ACLName="";
        public string accountName="";

        public int allowLocalAccess, denyLocalAccess;
        public int allowRemoteAccess, denyRemoteAccess;

        public int allowLocalLaunch, denyLocalLaunch;
        public int allowRemoteLaunch, denyRemoteLaunch;
        public int allowLocalActive, denyLocalActive;
        public int allowRemoteActive, denyRemoteActive;
    }

    class opcServer {
        public string uuid="";
        public string name="";
        public bool isExist;
        public int AccessFlag;//0 默认 1 非默认
        public int LaunchFlag;
    }

}
