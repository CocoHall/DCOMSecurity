using System;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.DirectoryServices.AccountManagement;


namespace DCOMSecurity
{
    static unsafe class userAndGroup
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct USER_INFO_0
        {
            public string name;
        }

        public struct USER_INFO_23
        {
            public string name;
            public string fullname;
            public string comment;
            public int flags;
            public IntPtr psid;
        }

        [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        internal struct LOCALGROUP_INFO_1
        {
            public IntPtr lpszGroupName;
            public IntPtr lpszComment;
        }

        [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        internal struct LOCALGROUP_MEMBERS_INFO_1
        {
            public IntPtr lgrmi1_sid;
            public IntPtr lgrmi1_sidusage;
            public IntPtr lgrmi1_name;

        }


        [DllImport("Netapi32.dll")]
        internal static extern int NetUserEnum( 
            [MarshalAs(UnmanagedType.LPWStr)]
            string servername,
            int level,
            int filter,
            out IntPtr bufptr,
            int prefmaxlen,
            out int entriesread,
            out int totalentries,
            out int resume_handle);

        [DllImport("Netapi32.dll")]
        internal static extern int NetApiBufferFree(IntPtr Buffer);

        [DllImport("Netapi32.dll")]
        internal extern static  int NetUserGetInfo([MarshalAs(UnmanagedType.LPWStr)] string servername, [MarshalAs(UnmanagedType.LPWStr)] string username, int level, out IntPtr bufptr);

        [DllImport("ADVAPI32.dll")]
        internal static extern bool ConvertSidToStringSid(
            IntPtr Sid,
            out string stringSid
            );

        public static string[] getAllUser() {
            int EntriesRead;
            int TotalEntries;
            int Resume;
            IntPtr bufPtr;

            NetUserEnum(null, 0, 2, out bufPtr, -1, out EntriesRead,
                out TotalEntries, out Resume);

            string[] names = new string[EntriesRead];


            IntPtr iter = bufPtr;
            for (int i = 0; i < EntriesRead; i++)
            {
                USER_INFO_0 User0 = (USER_INFO_0)Marshal.PtrToStructure(iter,
                    typeof(USER_INFO_0));

                iter = (IntPtr)(iter.ToInt64() + Marshal.SizeOf(typeof(USER_INFO_0)));

                names[i] = User0.name;
                
            }
            NetApiBufferFree(bufPtr);
            return names;

        }

        public static string getSidByName(string name) {
            IntPtr puser23;
            string stringSid="";
            userAndGroup.NetUserGetInfo(null, name, 23, out puser23);
            if (puser23 != IntPtr.Zero)
            {
                USER_INFO_23 user23 = (USER_INFO_23)Marshal.PtrToStructure(puser23,
                    typeof(USER_INFO_23));

                ConvertSidToStringSid(user23.psid, out stringSid);
            }

            return stringSid;
        }

        public static int CreateLocalWindowsAccount(string username, string password, string displayName, string description, bool canChangePwd, bool pwdExpires)
        {
            try
            {
                PrincipalContext context = new PrincipalContext(ContextType.Machine);
                UserPrincipal user = new UserPrincipal(context);
                user.SetPassword(password);
                user.DisplayName = displayName;
                user.Name = username;
                user.Description = description;
                user.UserCannotChangePassword = canChangePwd;
                user.PasswordNeverExpires = pwdExpires;
                user.Save();
                //now add user to "Users" group so it displays in Control Panel
                GroupPrincipal group = GroupPrincipal.FindByIdentity(context, "Users");
                group.Members.Add(user);
                group.Save();

                GroupPrincipal group2 = GroupPrincipal.FindByIdentity(context, "Administrators");
                group2.Members.Add(user);
                group2.Save();
                return 1;
            }
            catch (Exception ex)
            {
                return -1;
            }
        }

        [DllImport("netapi32.dll")]
        internal static extern uint NetLocalGroupGetMembers(
        IntPtr ServerName,
        IntPtr GrouprName,
        uint level,
        ref IntPtr siPtr,
        uint prefmaxlen,
        ref uint entriesread,
        ref uint totalentries,
        IntPtr resumeHandle);

        [DllImport("netapi32.dll")]
        internal static extern uint NetLocalGroupEnum(
        IntPtr ServerName,
        uint level,
        ref IntPtr siPtr,
        uint prefmaxlen,
        ref uint entriesread,
        ref uint totalentries,
        IntPtr resumeHandle);
    
    

    public static string[] getGroupsByUsername(string username) {
            uint level = 1, prefmaxlen = 0xFFFFFFFF, entriesread = 0, totalentries = 0;
            string groups = "";
            IntPtr GroupInfoPtr=IntPtr.Zero, UserInfoPtr= IntPtr.Zero;

            NetLocalGroupEnum(
                IntPtr.Zero, 
                level,
                ref GroupInfoPtr,
                prefmaxlen,
                ref entriesread,
                ref totalentries,
                IntPtr.Zero);


            for (int i = 0; i < totalentries; i++)
            {

                long newOffset = GroupInfoPtr.ToInt64() + sizeof(LOCALGROUP_INFO_1) * i;
                LOCALGROUP_INFO_1 groupInfo = (LOCALGROUP_INFO_1)Marshal.PtrToStructure(new IntPtr(newOffset), typeof(LOCALGROUP_INFO_1));
                string currentGroupName = Marshal.PtrToStringAuto(groupInfo.lpszGroupName);

                uint prefmaxlen1 = 0xFFFFFFFF, entriesread1 = 0, totalentries1 = 0;

                NetLocalGroupGetMembers(IntPtr.Zero, groupInfo.lpszGroupName, 1, ref UserInfoPtr, prefmaxlen1, ref entriesread1, ref totalentries1, IntPtr.Zero);

                //getting members name
                for (int j = 0; j < totalentries1; j++)
                {
                    long newOffset1 = UserInfoPtr.ToInt64() + sizeof(LOCALGROUP_MEMBERS_INFO_1) * j;
                    LOCALGROUP_MEMBERS_INFO_1 memberInfo = (LOCALGROUP_MEMBERS_INFO_1)Marshal.PtrToStructure(new IntPtr(newOffset1), typeof(LOCALGROUP_MEMBERS_INFO_1));
                    string currentUserName = Marshal.PtrToStringAuto(memberInfo.lgrmi1_name);
                    if (currentUserName.ToUpper().Equals(username.ToUpper())) {
                        if (groups.Length == 0)
                            groups += currentGroupName ;
                        else
                            groups += ","+currentGroupName;
                        break;
                    }

                }
                NetApiBufferFree(UserInfoPtr);
            }
            NetApiBufferFree(GroupInfoPtr);

            return groups.Split(',');


        }

    }
}
