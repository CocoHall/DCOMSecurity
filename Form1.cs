using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace DCOMSecurity
{
    public partial class Form1 : Form
    {

        string currentName = "";
        string currentSID = "";
        List<mSID> insideSidList = new List<mSID>();//内置的一些账号名/组 和sid/aclname

        List<mUser> localSidList = new List<mUser>();//存储本地的账号名和sid
        List<commACL> commAclTable = new List<commACL>();

        List<opcACL> opcAclTable = new List<opcACL>();
        List<opcServer> opclist = new List<opcServer>();

        DataGridView dvComm, dvOpc, dvUser;

        const int unsafeColor = 0xBF8F00;

        //设置表格列名等
        private void initDataGridView() {

            //账号情况下的datagridview
            dvUser = new DataGridView();
            DataGridViewTextBoxColumn[] colsUser = new DataGridViewTextBoxColumn[5];
            for (int i = 0; i < 5; i++)
            {
                colsUser[i] = new System.Windows.Forms.DataGridViewTextBoxColumn();
            }
            dvUser.RowHeadersVisible = false;
            colsUser[0].HeaderText = "本地账号";
            colsUser[1].HeaderText = "远程访问";
            colsUser[2].HeaderText = "远程启动";
            colsUser[3].HeaderText = "远程激活";
            colsUser[4].HeaderText = "安全性";
            dvUser.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabPage3.Controls.Add(dvUser);

            PropertyInfo piUser = dvUser.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);

            piUser.SetValue(dvUser, true, null);

            dvUser.AllowUserToAddRows = false;
            dvUser.AllowUserToDeleteRows = false;
            dvUser.ReadOnly = true;// 不能增加删除，只读
            dvUser.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;

            dvUser.EnableHeadersVisualStyles = false;
            dvUser.ColumnHeadersDefaultCellStyle.Font = new Font("方正", 8, FontStyle.Bold);
            //dv.ColumnHeadersDefaultCellStyle.s = Color.Red;

            dvUser.AllowUserToResizeColumns = false;//禁止修改宽度
            dvUser.AllowUserToResizeRows = false;//禁止修改高度
            dvUser.BackgroundColor = Color.White;//默认没数据部分太丑了
            dvUser.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            colsUser[0],colsUser[1],colsUser[2],colsUser[3],colsUser[4] });
            for (int i = 0; i < dvUser.Columns.Count; i++)
            {
                dvUser.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            dvUser.ColumnHeadersHeight = 30;
            dvUser.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            

            //通用设置里的datagridview

            dvComm = new DataGridView();

            DataGridViewTextBoxColumn colsName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            DataGridViewComboBoxColumn[] colsComm = new DataGridViewComboBoxColumn[12];
            DataGridViewButtonColumn colsBtn = new DataGridViewButtonColumn();

            colsBtn.UseColumnTextForButtonValue = true;
            colsBtn.Text = "删除";

            for (int i = 0; i < 12; i++)
            {
                colsComm[i] = new System.Windows.Forms.DataGridViewComboBoxColumn();
                colsComm[i].DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            }
            dvComm.RowHeadersVisible = false;

            colsName.HeaderText = "账号/组";

            colsComm[0].HeaderText = "默认本地访问";
            colsComm[1].HeaderText = "默认远程访问";
            colsComm[2].HeaderText = "限制本地访问";
            colsComm[3].HeaderText = "限制远程访问";
            colsComm[4].HeaderText = "默认本地启动";
            colsComm[5].HeaderText = "默认远程启动";
            colsComm[6].HeaderText = "默认本地激活";
            colsComm[7].HeaderText = "默认远程激活";
            colsComm[8].HeaderText = "限制本地启动";
            colsComm[9].HeaderText = "限制远程启动";
            colsComm[10].HeaderText = "限制本地激活";
            colsComm[11].HeaderText = "限制远程激活";
            colsBtn.HeaderText = "删除";


            DataTable source = new DataTable();
            source.Columns.Add("permission");
            source.Rows.Add();
            source.Rows[0][0] = "允许";
            source.Rows.Add();
            source.Rows[1][0] = "允许、拒绝";
            source.Rows.Add();
            source.Rows[2][0] = "拒绝";
            source.Rows.Add();
            source.Rows[3][0] = "无";

            for (int i = 0; i < 12; i++)
            {
                colsComm[i].DataSource = source;
                colsComm[i].DisplayMember = "permission";
                colsComm[i].ValueMember = "permission";
 
            }

            dvComm.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabPage1.Controls.Add(dvComm);

            PropertyInfo piComm = dvComm.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);

            piComm.SetValue(dvComm, true, null);
            dvComm.AllowUserToAddRows = false;
            dvComm.AllowUserToDeleteRows = false;

            dvComm.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;

            dvComm.EnableHeadersVisualStyles = false;
            dvComm.ColumnHeadersDefaultCellStyle.Font = new Font("方正", 8, FontStyle.Bold);

            dvComm.AllowUserToResizeColumns = false;//禁止修改宽度
            dvComm.AllowUserToResizeRows = false;//禁止修改高度
            dvComm.BackgroundColor = Color.White;//默认没数据部分太丑了
            dvComm.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            colsName,colsComm[0],colsComm[1],colsComm[2],colsComm[3],colsComm[4],colsComm[5],colsComm[6],colsComm[7],colsComm[8],colsComm[9],colsComm[10],colsComm[11],colsBtn});
            dvComm.Columns[0].ReadOnly = true;//不能修改用户名/组列
            for (int i = 0; i < dvComm.Columns.Count; i++)
            {
                dvComm.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            dvComm.ColumnHeadersHeight = 30;
            dvComm.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            dvComm.CellContentClick += dvComm_CellContentClick;

            //SUPCON OPC -------------------------------------------------------------------------------------
            dvOpc = new DataGridView();
            DataGridViewTextBoxColumn[] colsOpc = new DataGridViewTextBoxColumn[8];
            for (int i = 0; i < 8; i++)
            {
                colsOpc[i] = new System.Windows.Forms.DataGridViewTextBoxColumn();
            }
            dvOpc.RowHeadersVisible = false;
            colsOpc[0].HeaderText = "OPC Server";
            colsOpc[1].HeaderText = "账号/组";
            colsOpc[2].HeaderText = "本地访问";
            colsOpc[3].HeaderText = "远程访问";
            colsOpc[4].HeaderText = "本地启动";
            colsOpc[5].HeaderText = "远程启动";
            colsOpc[6].HeaderText = "本地激活";
            colsOpc[7].HeaderText = "远程激活";


            dvOpc.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabPage2.Controls.Add(dvOpc);

            PropertyInfo piOpc = dvComm.GetType().GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);

            piOpc.SetValue(dvOpc, true, null);
            dvOpc.AllowUserToAddRows = false;
            dvOpc.AllowUserToDeleteRows = false;
            dvOpc.ReadOnly = true;// 不能增加删除，只读
            dvOpc.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;

            dvOpc.EnableHeadersVisualStyles = false;
            dvOpc.ColumnHeadersDefaultCellStyle.Font = new Font("方正", 8, FontStyle.Bold);
            //dv.ColumnHeadersDefaultCellStyle.s = Color.Red;

            dvOpc.AllowUserToResizeColumns = false;//禁止修改宽度
            dvOpc.AllowUserToResizeRows = false;//禁止修改高度
            dvOpc.BackgroundColor = Color.White;//默认没数据部分太丑了
            dvOpc.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            colsOpc[0],colsOpc[1],colsOpc[2],colsOpc[3],colsOpc[4],colsOpc[5],colsOpc[6],colsOpc[7] });

            for (int i = 0; i < dvOpc.Columns.Count; i++)
            {
                dvOpc.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            dvOpc.ColumnHeadersHeight = 30;
            dvOpc.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            dvOpc.CellPainting += dvOpc_CellPainting;

        }

        //合并单元格
        private void dvOpc_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {

            // 对第1列相同单元格进行合并     
            if (e.ColumnIndex == 0 && e.RowIndex != -1)
            {
                Brush datagridBrush = new SolidBrush(dvOpc.GridColor);
                SolidBrush groupLineBrush = new SolidBrush(e.CellStyle.BackColor);
                using (Pen datagridLinePen = new Pen(datagridBrush))
                {
                    // 清除单元格
                    e.Graphics.FillRectangle(groupLineBrush, e.CellBounds);
                    if (e.RowIndex < dvOpc.Rows.Count - 1 && dvOpc.Rows[e.RowIndex + 1].Cells[e.ColumnIndex].Value != null && dvOpc.Rows[e.RowIndex + 1].Cells[e.ColumnIndex].Value.ToString() != e.Value.ToString())
                    {
                        //绘制底边线
                        e.Graphics.DrawLine(datagridLinePen, e.CellBounds.Left, e.CellBounds.Bottom - 1, e.CellBounds.Right, e.CellBounds.Bottom - 1);
                        // 画右边线
                        e.Graphics.DrawLine(datagridLinePen, e.CellBounds.Right - 1, e.CellBounds.Top, e.CellBounds.Right - 1, e.CellBounds.Bottom);
                    }
                    else
                    {
                        // 画右边线
                        e.Graphics.DrawLine(datagridLinePen, e.CellBounds.Right - 1, e.CellBounds.Top, e.CellBounds.Right - 1, e.CellBounds.Bottom);
                    }
                    //对最后一条记录只画底边线
                    if (e.RowIndex == dvOpc.Rows.Count - 1)
                    {
                        //绘制底边线
                        e.Graphics.DrawLine(datagridLinePen, e.CellBounds.Left, e.CellBounds.Bottom - 1, e.CellBounds.Right, e.CellBounds.Bottom - 1);
                    }
                    //填写单元格内容，相同的内容的单元格只填写第一个
                    if (e.Value != null)
                    {
                        if (!(e.RowIndex > 0 && dvOpc.Rows[e.RowIndex - 1].Cells[e.ColumnIndex].Value.ToString() == e.Value.ToString()))
                        {
                            //绘制单元格内容
                            e.Graphics.DrawString(e.Value.ToString(), e.CellStyle.Font, Brushes.Black, e.CellBounds.X + 2, e.CellBounds.Y + 5, StringFormat.GenericDefault);
                        }
                    }
                    e.Handled = true;
                }
            }

        }

        //删除一行
        private void dvComm_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 13)
            {
                dvComm.Rows.RemoveAt(e.RowIndex);
            }
        }

        //加载内置SDDL和UUID
        private void initList() {

            string[][] insideUser = new string[][]{
                new string[] { "PS","self" },
                new string[] { "BA","Administrators" },
                new string[] { "SY","SYSTEM" },
                new string[] { "LA","Administrator" },
                new string[] { "NU","NETWORK" },
                new string[] { "NS","NETWORD SERVICE" },
                new string[] { "WD","Everyone" },
                new string[] { "AN","ANONYMOUS LOGON" },
                new string[] { "IU","INTERACTIVE" },
                new string[] { "BU","Users" },
                new string[] { "BG","Guests" },
                new string[] { "LG","Guest" },
                new string[] { "LS","LOCAL SERVICE" },
                new string[] { "AU","Authenticated Users" },
                new string[] { "AC","All APPLICATION PACKAGES" },
                new string[] { "LU","Performance Log Users" },
                new string[] { "S-1-5-32-562","Distributed COM Users" }
            };

            foreach (var tmp in insideUser)
            {
                mSID tmpSID = new mSID();
                tmpSID.sid = tmp[0];
                tmpSID.name = tmp[1];
                
                insideSidList.Add(tmpSID);
            }


            string[][] insideOpcServer = new string[][]{
                new string[] { "41EBD53D-36C4-4027-B2B4-09A6E4A362DD","SUPCON.SCRTCore" },
                new string[] { "A8E1EC00-1F75-11d4-9775-0000E8A370F0","SUPCON.JXServer" },
                new string[] { "13486D44-4821-11D2-A494-3CB306C10000","OpcEnum" },
                new string[] { "CC7D4C09-B26B-45B3-B482-2DECA09B832A", "OPC Enum x64 CategoryManager" },
                new string[] { "F8582CF3-88FB-11D0-B850-00C0F0104305", "MatrikonOPC Server for Simulation and Testing" },

            };

            foreach (var tmp in insideOpcServer)
            {
                opcServer tmpOpcServer = new opcServer();
                tmpOpcServer.uuid = tmp[0];
                tmpOpcServer.name = tmp[1];
                opclist.Add(tmpOpcServer);
            }
        }

        private void update() {
            addAllSid();
            updateCommAclTableFromReg();
            updateOpcAclTableFromReg();
            changeName();
            judgeSecurity();
            updateDataGridView();
        }

        public Form1()
        {
            InitializeComponent();
            initDataGridView();
            initList();
            getWhoami();
            updateCurrentName(currentName);
            update();
        }

        //通过whoami /user获得账号名和sid，可以获得域账号
        public void getWhoami()
        {
            string output = "";     //输出字符串
            string dosCommand = "whoami /user";
            if (dosCommand != null && dosCommand != "")
            {
                Process process = new Process();     //创建进程对象
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = "cmd.exe";      //设定需要执行的命令
                startInfo.Arguments = "/C " + dosCommand;   //设定参数，其中的“/C”表示执行完命令后马上退出
                startInfo.UseShellExecute = false;     //不使用系统外壳程序启动
                startInfo.RedirectStandardInput = false;   //不重定向输入
                startInfo.RedirectStandardOutput = true;   //重定向输出
                startInfo.CreateNoWindow = true;     //不创建窗口
                process.StartInfo = startInfo;
                try
                {
                    if (process.Start())       //开始进程
                    {
                        process.WaitForExit();  //当超过这个毫秒后，不管执行结果如何，都不再执行这个DOS命令
                        output = process.StandardOutput.ReadToEnd();//读取进程的输出
                    }
                }
                catch
                {
                }
                finally
                {
                    if (process != null)
                        process.Close();
                }
            }

            Match match = Regex.Match(output, @"\s(.*?) (S-.*?)\s");

            if (match.Groups.Count == 3) {
                currentName = match.Groups[1].ToString();
                currentSID = match.Groups[2].ToString();
                //addSid(currentName, currentSID);
            }

            return;

        }

        //获得本地的用户名和sid到内存localSidList中
        public void addAllSid() {
            int index = comboBox1.SelectedIndex;

            if (comboBox1 != null && comboBox1.Items.Count > 0)
                comboBox1.Items.Clear();

            localSidList.Clear();
            string[] names = userAndGroup.getAllUser();
            foreach (string name in names) {
                string sid = getSidByName(name);
                comboBox1.Items.Add(name);
                addSid(name, sid);
            }

            if (index < comboBox1.Items.Count) {
                comboBox1.SelectedIndex = index;
            }

        }

        private void addSid(string name, string sid) {
            for (int i = 0; i < localSidList.Count; i++) {
                if (localSidList[i].sid.ToUpper().Equals(sid.ToUpper())) {
                    return;
                }
            }
            mUser t = new mUser();
            t.name = name;
            t.sid = sid;
            localSidList.Add(t);
        }

        //从注册表中查询几个OPC的自定义权限，存入内存opcAclTable中
        private void updateOpcAclTableFromReg() {

            opcAclTable.Clear();
            foreach (opcServer tmp in opclist) {
                int result1, result2;
                string AccessPermission = regUnit.getOpcSDDL(tmp.uuid, "AccessPermission", out result1);
                string LaunchPermission = regUnit.getOpcSDDL(tmp.uuid, "LaunchPermission", out result2);
                switch (result1) {
                    case 0:
                        tmp.AccessFlag = 1;
                        tmp.isExist = true;
                        updateOpcAclTableAccess(tmp.uuid, tmp.name, AccessPermission); break;
                    case -1: tmp.isExist = true; tmp.AccessFlag = 0; break;
                    case -2: tmp.isExist = false; break;
                }
                switch (result2)
                {
                    case 0:
                        tmp.LaunchFlag = 1;
                        tmp.isExist = true;
                        updateOpcAclTableLaunch(tmp.uuid, tmp.name, LaunchPermission); break;
                    case -1: tmp.isExist = true; tmp.LaunchFlag = 0; break;
                    case -2: tmp.isExist = false; break;
                }
            }

        }

        //从注册表中获得通用设置权限，放入commAclTable
        private void updateCommAclTableFromReg() {

            commAclTable.Clear();
            string DefaultAccessPermission = regUnit.getCommSDDL("DefaultAccessPermission");
            string DefaultLaunchPermission = regUnit.getCommSDDL("DefaultLaunchPermission");
            string MachineAccessRestriction = regUnit.getCommSDDL("MachineAccessRestriction");
            string MachineLaunchRestriction = regUnit.getCommSDDL("MachineLaunchRestriction");

            updateCommAclTableDefaultAccessFromStr(DefaultAccessPermission);
            updateCommAclTableDefaultLaunchFromStr(DefaultLaunchPermission);
            updateCommAclTableRestrictionAccessFromStr(MachineAccessRestriction);
            updateCommAclTableRestrictionLaunchFromStr(MachineLaunchRestriction);

        }

        #region 将SDDL字符串写入commAclTable

        //传入SDDL string ，设置内存中的权限表
        private void updateCommAclTableDefaultAccessFromStr(string DefaultAccessPermission) {
            string[] permissions = DefaultAccessPermission.Split('(');
            foreach (string permission in permissions) {
                string[] tmps = permission.Split(';');
                if (tmps.Length == 6) {
                    bool allowOrDeny;
                    if (tmps[0].Equals("A")) allowOrDeny = true;
                    else if (tmps[0].Equals("D")) allowOrDeny = false;
                    else return;
                    string acl = tmps[2];
                    string user = tmps[5].Trim(')');

                    #region 解析acl
                    int allowDefaultLocalAccess = 0, denyDefaultLocalAccess = 0;
                    int allowDefaultRemoteAccess = 0, denyDefaultRemoteAccess = 0;
                    if (acl.Contains("DC"))
                    {
                        if (allowOrDeny)
                        {
                            allowDefaultLocalAccess = 1;
                        }
                        else
                        {
                            denyDefaultLocalAccess = 1;
                        }
                    }

                    if (acl.Contains("LC"))
                    {
                        if (allowOrDeny)
                        {
                            allowDefaultRemoteAccess = 1;
                        }
                        else
                        {
                            denyDefaultRemoteAccess = 1;
                        }
                    }
                    #endregion

                    int flag = 0;



                    for (int i = 0; i < commAclTable.Count; i++) {
                        commACL t = commAclTable[i];
                        if (t.sidOrAclName.ToUpper().Equals(user.ToUpper())) {
                            flag = 1;
                            if (allowOrDeny)
                            {
                                t.allowDefaultLocalAccess = allowDefaultLocalAccess;
                                t.allowDefaultRemoteAccess = allowDefaultRemoteAccess;
                            }
                            else {
                                t.denyDefaultLocalAccess = denyDefaultLocalAccess;
                                t.denyDefaultRemoteAccess = denyDefaultRemoteAccess;
                            }
                        }
                    }


                    if (flag == 0) {
                        commACL t = new commACL();
                        t.displayName = t.sidOrAclName = user;
                        if (allowOrDeny)
                        {
                            t.allowDefaultLocalAccess = allowDefaultLocalAccess;
                            t.allowDefaultRemoteAccess = allowDefaultRemoteAccess;
                        }
                        else
                        {
                            t.denyDefaultLocalAccess = denyDefaultLocalAccess;
                            t.denyDefaultRemoteAccess = denyDefaultRemoteAccess;
                        }
                        commAclTable.Add(t);

                    }

                }
            }
        }

        private void updateCommAclTableDefaultLaunchFromStr(string DefaultLaunchPermission)
        {
            string[] permissions = DefaultLaunchPermission.Split('(');
            foreach (string permission in permissions)
            {
                string[] tmps = permission.Split(';');
                if (tmps.Length == 6)
                {
                    bool allowOrDeny;
                    if (tmps[0].Equals("A")) allowOrDeny = true;
                    else if (tmps[0].Equals("D")) allowOrDeny = false;
                    else return;
                    string acl = tmps[2];
                    string user = tmps[5].Trim(')');

                    #region 解析acl
                    int allowDefaultLocalLaunch = 0, denyDefaultLocalLaunch = 0;
                    int allowDefaultRemoteLaunch = 0, denyDefaultRemoteLaunch = 0;
                    int allowDefaultLocalActive = 0, denyDefaultLocalActive = 0;
                    int allowDefaultRemoteActive = 0, denyDefaultRemoteActive = 0;
                    if (acl.Contains("DC"))
                    {
                        if (allowOrDeny)
                        {
                            allowDefaultLocalLaunch = 1;
                        }
                        else
                        {
                            denyDefaultLocalLaunch = 1;
                        }
                    }

                    if (acl.Contains("LC"))
                    {
                        if (allowOrDeny)
                        {
                            allowDefaultRemoteLaunch = 1;
                        }
                        else
                        {
                            denyDefaultRemoteLaunch = 1;
                        }
                    }

                    if (acl.Contains("SW"))
                    {
                        if (allowOrDeny)
                        {
                            allowDefaultLocalActive = 1;
                        }
                        else
                        {
                            denyDefaultLocalActive = 1;
                        }
                    }

                    if (acl.Contains("RP"))
                    {
                        if (allowOrDeny)
                        {
                            allowDefaultRemoteActive = 1;
                        }
                        else
                        {
                            denyDefaultRemoteActive = 1;
                        }
                    }
                    #endregion

                    int flag = 0;



                    for (int i = 0; i < commAclTable.Count; i++)
                    {
                        commACL t = commAclTable[i];
                        if (t.sidOrAclName.ToUpper().Equals(user.ToUpper()))
                        {
                            flag = 1;
                            if (allowOrDeny)
                            {
                                t.allowDefaultLocalLaunch = allowDefaultLocalLaunch;
                                t.allowDefaultRemoteLaunch = allowDefaultRemoteLaunch;
                                t.allowDefaultLocalActive = allowDefaultLocalActive;
                                t.allowDefaultRemoteActive = allowDefaultRemoteActive;
                            }
                            else
                            {
                                t.denyDefaultLocalLaunch = denyDefaultLocalLaunch;
                                t.denyDefaultRemoteLaunch = denyDefaultRemoteLaunch;
                                t.denyDefaultLocalActive = denyDefaultLocalActive;
                                t.denyDefaultRemoteActive = denyDefaultRemoteActive;
                            }
                        }
                    }


                    if (flag == 0)
                    {
                        commACL t = new commACL();
                        t.displayName = t.sidOrAclName = user;
                        if (allowOrDeny)
                        {
                            t.allowDefaultLocalLaunch = allowDefaultLocalLaunch;
                            t.allowDefaultRemoteLaunch = allowDefaultRemoteLaunch;
                            t.allowDefaultLocalActive = allowDefaultLocalActive;
                            t.allowDefaultRemoteActive = allowDefaultRemoteActive;
                        }
                        else
                        {
                            t.denyDefaultLocalLaunch = denyDefaultLocalLaunch;
                            t.denyDefaultRemoteLaunch = denyDefaultRemoteLaunch;
                            t.denyDefaultLocalActive = denyDefaultLocalActive;
                            t.denyDefaultRemoteActive = denyDefaultRemoteActive;
                        }
                        commAclTable.Add(t);

                    }

                }
            }
        }

        private void updateCommAclTableRestrictionAccessFromStr(string MachineAccessRestriction)
        {
            string[] permissions = MachineAccessRestriction.Split('(');
            foreach (string permission in permissions)
            {
                string[] tmps = permission.Split(';');
                if (tmps.Length == 6)
                {
                    bool allowOrDeny;
                    if (tmps[0].Equals("A")) allowOrDeny = true;
                    else if (tmps[0].Equals("D")) allowOrDeny = false;
                    else return;
                    string acl = tmps[2];
                    string user = tmps[5].Trim(')');

                    #region 解析acl
                    int allowRestrictionLocalAccess = 0, denyRestrictionLocalAccess = 0;
                    int allowRestrictionRemoteAccess = 0, denyRestrictionRemoteAccess = 0;
                    if (acl.Contains("DC"))
                    {
                        if (allowOrDeny)
                        {
                            allowRestrictionLocalAccess = 1;
                        }
                        else
                        {
                            denyRestrictionLocalAccess = 1;
                        }
                    }

                    if (acl.Contains("LC"))
                    {
                        if (allowOrDeny)
                        {
                            allowRestrictionRemoteAccess = 1;
                        }
                        else
                        {
                            denyRestrictionRemoteAccess = 1;
                        }
                    }
                    #endregion

                    int flag = 0;



                    for (int i = 0; i < commAclTable.Count; i++)
                    {
                        commACL t = commAclTable[i];
                        if (t.sidOrAclName.ToUpper().Equals(user.ToUpper()))
                        {
                            flag = 1;
                            if (allowOrDeny)
                            {
                                t.allowRestrictionLocalAccess = allowRestrictionLocalAccess;
                                t.allowRestrictionRemoteAccess = allowRestrictionRemoteAccess;
                            }
                            else
                            {
                                t.denyRestrictionLocalAccess = denyRestrictionLocalAccess;
                                t.denyRestrictionRemoteAccess = denyRestrictionRemoteAccess;
                            }
                        }
                    }


                    if (flag == 0)
                    {
                        commACL t = new commACL();
                        t.displayName = t.sidOrAclName = user;
                        if (allowOrDeny)
                        {
                            t.allowRestrictionLocalAccess = allowRestrictionLocalAccess;
                            t.allowRestrictionRemoteAccess = allowRestrictionRemoteAccess;
                        }
                        else
                        {
                            t.denyRestrictionLocalAccess = denyRestrictionLocalAccess;
                            t.denyRestrictionRemoteAccess = denyRestrictionRemoteAccess;
                        }
                        commAclTable.Add(t);

                    }

                }
            }
        }

        private void updateCommAclTableRestrictionLaunchFromStr(string MachineLaunchRestriction)
        {
            string[] permissions = MachineLaunchRestriction.Split('(');
            foreach (string permission in permissions)
            {
                string[] tmps = permission.Split(';');
                if (tmps.Length == 6)
                {
                    bool allowOrDeny;
                    if (tmps[0].Equals("A")) allowOrDeny = true;
                    else if (tmps[0].Equals("D")) allowOrDeny = false;
                    else return;
                    string acl = tmps[2];
                    string user = tmps[5].Trim(')');

                    #region 解析acl
                    int allowRestrictionLocalLaunch = 0, denyRestrictionLocalLaunch = 0;
                    int allowRestrictionRemoteLaunch = 0, denyRestrictionRemoteLaunch = 0;
                    int allowRestrictionLocalActive = 0, denyRestrictionLocalActive = 0;
                    int allowRestrictionRemoteActive = 0, denyRestrictionRemoteActive = 0;
                    if (acl.Contains("DC"))
                    {
                        if (allowOrDeny)
                        {
                            allowRestrictionLocalLaunch = 1;
                        }
                        else
                        {
                            denyRestrictionLocalLaunch = 1;
                        }
                    }

                    if (acl.Contains("LC"))
                    {
                        if (allowOrDeny)
                        {
                            allowRestrictionRemoteLaunch = 1;
                        }
                        else
                        {
                            denyRestrictionRemoteLaunch = 1;
                        }
                    }

                    if (acl.Contains("SW"))
                    {
                        if (allowOrDeny)
                        {
                            allowRestrictionLocalActive = 1;
                        }
                        else
                        {
                            denyRestrictionLocalActive = 1;
                        }
                    }

                    if (acl.Contains("RP"))
                    {
                        if (allowOrDeny)
                        {
                            allowRestrictionRemoteActive = 1;
                        }
                        else
                        {
                            denyRestrictionRemoteActive = 1;
                        }
                    }
                    #endregion

                    int flag = 0;



                    for (int i = 0; i < commAclTable.Count; i++)
                    {
                        commACL t = commAclTable[i];
                        if (t.sidOrAclName.ToUpper().Equals(user.ToUpper()))
                        {
                            flag = 1;
                            if (allowOrDeny)
                            {
                                t.allowRestrictionLocalLaunch = allowRestrictionLocalLaunch;
                                t.allowRestrictionRemoteLaunch = allowRestrictionRemoteLaunch;
                                t.allowRestrictionLocalActive = allowRestrictionLocalActive;
                                t.allowRestrictionRemoteActive = allowRestrictionRemoteActive;
                            }
                            else
                            {
                                t.denyRestrictionLocalLaunch = denyRestrictionLocalLaunch;
                                t.denyRestrictionRemoteLaunch = denyRestrictionRemoteLaunch;
                                t.denyRestrictionLocalActive = denyRestrictionLocalActive;
                                t.denyRestrictionRemoteActive = denyRestrictionRemoteActive;
                            }
                        }
                    }


                    if (flag == 0)
                    {
                        commACL t = new commACL();
                        t.displayName = t.sidOrAclName = user;
                        if (allowOrDeny)
                        {
                            t.allowRestrictionLocalLaunch = allowRestrictionLocalLaunch;
                            t.allowRestrictionRemoteLaunch = allowRestrictionRemoteLaunch;
                            t.allowRestrictionLocalActive = allowRestrictionLocalActive;
                            t.allowRestrictionRemoteActive = allowRestrictionRemoteActive;
                        }
                        else
                        {
                            t.denyRestrictionLocalLaunch = denyRestrictionLocalLaunch;
                            t.denyRestrictionRemoteLaunch = denyRestrictionRemoteLaunch;
                            t.denyRestrictionLocalActive = denyRestrictionLocalActive;
                            t.denyRestrictionRemoteActive = denyRestrictionRemoteActive;
                        }
                        commAclTable.Add(t);

                    }

                }
            }
        }

        private void updateOpcAclTableAccess(string uuid, string opcName, string AccessPermission)
        {
            if (AccessPermission.Length == 0) {
                return;

            }
            string[] permissions = AccessPermission.Split('(');
            foreach (string permission in permissions)
            {
                string[] tmps = permission.Split(';');
                if (tmps.Length == 6)
                {
                    bool allowOrDeny;
                    if (tmps[0].Equals("A")) allowOrDeny = true;
                    else if (tmps[0].Equals("D")) allowOrDeny = false;
                    else return;
                    string acl = tmps[2];
                    string user = tmps[5].Trim(')');

                    #region 解析acl
                    int allowLocalAccess = 0, denyLocalAccess = 0;
                    int allowRemoteAccess = 0, denyRemoteAccess = 0;
                    if (acl.Contains("DC"))
                    {
                        if (allowOrDeny)
                        {
                            allowLocalAccess = 1;
                        }
                        else
                        {
                            denyLocalAccess = 1;
                        }
                    }

                    if (acl.Contains("LC"))
                    {
                        if (allowOrDeny)
                        {
                            allowRemoteAccess = 1;
                        }
                        else
                        {
                            denyRemoteAccess = 1;
                        }
                    }
                    #endregion

                    int flag = 0;



                    for (int i = 0; i < opcAclTable.Count; i++)
                    {
                        opcACL t = opcAclTable[i];
                        if (t.uuid.Equals(uuid))
                        {
                            if (t.ACLName.ToUpper().Equals(user.ToUpper()))
                            {
                                flag = 1;
                                if (allowOrDeny)
                                {
                                    t.allowLocalAccess = allowLocalAccess;
                                    t.allowRemoteAccess = allowRemoteAccess;
                                }
                                else
                                {
                                    t.denyLocalAccess = denyLocalAccess;
                                    t.denyRemoteAccess = denyRemoteAccess;
                                }
                            }
                        }
                    }


                    if (flag == 0)
                    {
                        opcACL t = new opcACL();
                        t.accountName = t.ACLName = user;
                        t.uuid = uuid;
                        t.opcName = opcName;

                        if (allowOrDeny)
                        {
                            t.allowLocalAccess = allowLocalAccess;
                            t.allowRemoteAccess = allowRemoteAccess;
                        }
                        else
                        {
                            t.denyLocalAccess = denyLocalAccess;
                            t.denyRemoteAccess = denyRemoteAccess;
                        }
                        opcAclTable.Add(t);

                    }

                }
            }
        }

        private void updateOpcAclTableLaunch(string uuid, string name, string LaunchPermission)
        {
            if (LaunchPermission.Length == 0)
            {
                return;
            }
            string[] permissions = LaunchPermission.Split('(');
            foreach (string permission in permissions)
            {
                string[] tmps = permission.Split(';');
                if (tmps.Length == 6)
                {
                    bool allowOrDeny;
                    if (tmps[0].Equals("A")) allowOrDeny = true;
                    else if (tmps[0].Equals("D")) allowOrDeny = false;
                    else return;
                    string acl = tmps[2];
                    string user = tmps[5].Trim(')');

                    #region 解析acl
                    int allowLocalLaunch = 0, denyLocalLaunch = 0;
                    int allowRemoteLaunch = 0, denyRemoteLaunch = 0;

                    int allowLocalActive = 0, denyLocalActive = 0;
                    int allowRemoteActive = 0, denyRemoteActive = 0;
                    if (acl.Contains("DC"))
                    {
                        if (allowOrDeny)
                        {
                            allowLocalLaunch = 1;
                        }
                        else
                        {
                            denyLocalLaunch = 1;
                        }
                    }

                    if (acl.Contains("LC"))
                    {
                        if (allowOrDeny)
                        {
                            allowRemoteLaunch = 1;
                        }
                        else
                        {
                            denyRemoteLaunch = 1;
                        }
                    }

                    if (acl.Contains("SW"))
                    {
                        if (allowOrDeny)
                        {
                            allowLocalActive = 1;
                        }
                        else
                        {
                            denyLocalActive = 1;
                        }
                    }

                    if (acl.Contains("RP"))
                    {
                        if (allowOrDeny)
                        {
                            allowRemoteActive = 1;
                        }
                        else
                        {
                            denyRemoteActive = 1;
                        }
                    }
                    #endregion

                    int flag = 0;



                    for (int i = 0; i < opcAclTable.Count; i++)
                    {
                        opcACL t = opcAclTable[i];
                        if (t.uuid.Equals(uuid))
                        {
                            if (t.ACLName.ToUpper().Equals(user.ToUpper()))
                            {
                                flag = 1;
                                if (allowOrDeny)
                                {
                                    t.allowLocalLaunch = allowLocalLaunch;
                                    t.allowRemoteLaunch = allowRemoteLaunch;
                                    t.allowLocalActive = allowLocalActive;
                                    t.allowRemoteActive = allowRemoteActive;
                                }
                                else
                                {
                                    t.denyLocalLaunch = denyLocalLaunch;
                                    t.denyRemoteLaunch = denyRemoteLaunch;
                                    t.denyLocalActive = denyLocalActive;
                                    t.denyRemoteActive = denyRemoteActive;
                                }
                            }
                        }
                    }


                    if (flag == 0)
                    {
                        opcACL t = new opcACL();
                        t.accountName = t.ACLName = user;
                        t.opcName = name;
                        t.uuid = uuid;
                        if (allowOrDeny)
                        {
                            t.allowLocalLaunch = allowLocalLaunch;
                            t.allowRemoteLaunch = allowRemoteLaunch;
                            t.allowLocalActive = allowLocalActive;
                            t.allowRemoteActive = allowRemoteActive;
                        }
                        else
                        {
                            t.denyLocalLaunch = denyLocalLaunch;
                            t.denyRemoteLaunch = denyRemoteLaunch;
                            t.denyLocalActive = denyLocalActive;
                            t.denyRemoteActive = denyRemoteActive;
                        }
                        opcAclTable.Add(t);

                    }

                }
            }
        }
        #endregion

        private void changeName()
        {

            for (int i = 0; i < commAclTable.Count; i++)
            {
                commACL tacl = commAclTable[i];
                int flag = 0;
                for (int j = 0; j < insideSidList.Count; j++)
                {
                    mSID tsid = insideSidList[j];
                    if (tacl.sidOrAclName.Equals(tsid.sid))
                    {
                        tacl.displayName = tsid.name;
                        flag = 1;
                        break;
                    }
                }
                if (flag == 0) {
                    for (int j = 0; j < localSidList.Count; j++)
                    {
                        mUser tsid = localSidList[j];
                        if (tacl.sidOrAclName.Equals(tsid.sid))
                        {
                            tacl.displayName = tsid.name;
                            flag = 1;
                            break;
                        }
                    }
                }
            }

            for (int i = 0; i < opcAclTable.Count; i++)
            {
                for (int j = 0; j < insideSidList.Count; j++)
                {
                    opcACL tacl = opcAclTable[i];
                    mSID tsid = insideSidList[j];
                    if (tacl.ACLName.Equals(tsid.sid))
                    {
                        tacl.accountName = tsid.name;
                        break;
                    }
                }
            }

        }

        private void judgeSecurity()
        {
            for (int i = 0; i < localSidList.Count; i++)
            {
                mUser tUser = localSidList[i];
                int allowDefaultRemoteAccess = 0, denyDefaultRemoteAccess = 0;
                int allowDefaultRemoteLaunch = 0, denyDefaultRemoteLaunch = 0;
                int allowDefaultRemoteActive = 0, denyDefaultRemoteActive = 0;
                int allowRestrictionRemoteAccess = 0, denyRestrictionRemoteAccess = 0;
                int allowRestrictionRemoteLaunch = 0, denyRestrictionRemoteLaunch = 0;
                int allowRestrictionRemoteActive = 0, denyRestrictionRemoteActive = 0;

                string[] groups = userAndGroup.getGroupsByUsername(tUser.name);
                for (int j = 0; j < groups.Length; j++)
                {
                    for (int k = 0; k < commAclTable.Count; k++)
                    {
                        commACL tcom = commAclTable[k];
                        if (tcom.displayName.ToUpper().Equals(groups[j].ToUpper())
                            || (tcom.displayName.ToUpper().Equals("EVERYONE")  && !tcom.displayName.ToUpper().Equals("GUESTS") && !tcom.displayName.ToUpper().Equals("GUEST"))
                            || tcom.displayName.ToUpper().Equals("SELF")
                            || tcom.displayName.ToUpper().Equals("AUTHENTICATED USERS")
                            || tcom.displayName.ToUpper().Equals("NETWORK")

                            )
                        {//访问控制表中的组名和账号所属的某个组相同
                            if (tcom.allowDefaultRemoteAccess == 1) allowDefaultRemoteAccess = 1;
                            if (tcom.allowDefaultRemoteLaunch == 1) allowDefaultRemoteLaunch = 1;
                            if (tcom.allowDefaultRemoteActive == 1) allowDefaultRemoteActive = 1;
                            if (tcom.allowRestrictionRemoteAccess == 1) allowRestrictionRemoteAccess = 1;
                            if (tcom.allowRestrictionRemoteLaunch == 1) allowRestrictionRemoteLaunch = 1;
                            if (tcom.allowRestrictionRemoteActive == 1) allowRestrictionRemoteActive = 1;
                             
                            if (tcom.denyDefaultRemoteAccess == 1) denyDefaultRemoteAccess = 1;
                            if (tcom.denyDefaultRemoteLaunch == 1) denyDefaultRemoteLaunch = 1;
                            if (tcom.denyDefaultRemoteActive == 1) denyDefaultRemoteActive = 1;
                            if (tcom.denyRestrictionRemoteAccess == 1) denyRestrictionRemoteAccess = 1;
                            if (tcom.denyRestrictionRemoteLaunch == 1) denyRestrictionRemoteLaunch = 1;
                            if (tcom.denyRestrictionRemoteActive == 1) denyRestrictionRemoteActive = 1;
                        }
                    }
                }
                if (denyDefaultRemoteAccess == 0 && denyRestrictionRemoteAccess == 0 && allowDefaultRemoteAccess == 1 && allowRestrictionRemoteAccess == 1)
                {
                    tUser.remoteAccess = 1;
                }

                if (denyDefaultRemoteLaunch == 0 && denyRestrictionRemoteLaunch == 0 && allowDefaultRemoteLaunch == 1 && allowRestrictionRemoteLaunch == 1)
                {
                    tUser.remoteLaunch = 1;
                }
                if (denyDefaultRemoteActive == 0 && denyRestrictionRemoteActive == 0 && allowDefaultRemoteActive == 1 && allowRestrictionRemoteActive == 1)
                {
                    tUser.remoteActive = 1;
                }

                if (tUser.remoteAccess == 1 && tUser.remoteLaunch == 1 && tUser.remoteActive == 1)
                {
                    tUser.safe = 0;
                }
                else if (tUser.remoteAccess == 1 && tUser.remoteActive == 1)
                {
                    tUser.safe = 2;
                }
                else {
                    tUser.safe = 1;
                }

                
            }

        
        }

        private void updateDataGridView()
        {


            dvUser.RowCount = 0;
            foreach (mUser t in localSidList) {
                string[] content = new string[5];
                content[0] = t.name;
                content[1] = t.remoteAccess==1?"允许":"无";
                content[2] = t.remoteLaunch == 1 ? "允许" : "无"; ;
                content[3] = t.remoteActive == 1 ? "允许" : "无";
                switch (t.safe) {
                    case 0: content[4] = "存在隐患"; break;
                    case 1: content[4] = "安全"; break;
                    case 2: content[4] = "账号活跃时存在隐患"; break;
                    default:content[4] = "未知";break;
                }

                dvUser.Rows.Add(content);
            }
            for (int i = 0; i < dvUser.RowCount; i++)
            {
                if (!dvUser.Rows[i].Cells[4].Value.ToString().Equals("安全"))
                {
                    dvUser.Rows[i].DefaultCellStyle.ForeColor = Color.FromArgb(unsafeColor);
                }
            }




            //comm设置的显示界面

            dvComm.RowCount = 0;
            foreach (commACL t in commAclTable)
            {
                string[] content = new string[14];
                content[0] = t.displayName;
                #region 显示允许还是拒绝
                if (t.allowDefaultLocalAccess == 1 && t.denyDefaultLocalAccess == 1) { content[1] = "允许、拒绝"; }
                else if (t.allowDefaultLocalAccess == 1 && t.denyDefaultLocalAccess == 0) { content[1] = "允许"; }
                else if (t.allowDefaultLocalAccess == 0 && t.denyDefaultLocalAccess == 1) { content[1] = "拒绝"; }
                else if (t.allowDefaultLocalAccess == 0 && t.denyDefaultLocalAccess == 0) { content[1] = "无"; }

                if (t.allowDefaultRemoteAccess == 1 && t.denyDefaultRemoteAccess == 1) { content[2] = "允许、拒绝"; }
                else if (t.allowDefaultRemoteAccess == 1 && t.denyDefaultRemoteAccess == 0) { content[2] = "允许"; }
                else if (t.allowDefaultRemoteAccess == 0 && t.denyDefaultRemoteAccess == 1) { content[2] = "拒绝"; }
                else if (t.allowDefaultRemoteAccess == 0 && t.denyDefaultRemoteAccess == 0) { content[2] = "无"; }

                if (t.allowRestrictionLocalAccess == 1 && t.denyRestrictionLocalAccess == 1) { content[3] = "允许、拒绝"; }
                else if (t.allowRestrictionLocalAccess == 1 && t.denyRestrictionLocalAccess == 0) { content[3] = "允许"; }
                else if (t.allowRestrictionLocalAccess == 0 && t.denyRestrictionLocalAccess == 1) { content[3] = "拒绝"; }
                else if (t.allowRestrictionLocalAccess == 0 && t.denyRestrictionLocalAccess == 0) { content[3] = "无"; }

                if (t.allowRestrictionRemoteAccess == 1 && t.denyRestrictionRemoteAccess == 1) { content[4] = "允许、拒绝"; }
                else if (t.allowRestrictionRemoteAccess == 1 && t.denyRestrictionRemoteAccess == 0) { content[4] = "允许"; }
                else if (t.allowRestrictionRemoteAccess == 0 && t.denyRestrictionRemoteAccess == 1) { content[4] = "拒绝"; }
                else if (t.allowRestrictionRemoteAccess == 0 && t.denyRestrictionRemoteAccess == 0) { content[4] = "无"; }

                if (t.allowDefaultLocalLaunch == 1 && t.denyDefaultLocalLaunch == 1) { content[5] = "允许、拒绝"; }
                else if (t.allowDefaultLocalLaunch == 1 && t.denyDefaultLocalLaunch == 0) { content[5] = "允许"; }
                else if (t.allowDefaultLocalLaunch == 0 && t.denyDefaultLocalLaunch == 1) { content[5] = "拒绝"; }
                else if (t.allowDefaultLocalLaunch == 0 && t.denyDefaultLocalLaunch == 0) { content[5] = "无"; }

                if (t.allowDefaultRemoteLaunch == 1 && t.denyDefaultRemoteLaunch == 1) { content[6] = "允许、拒绝"; }
                else if (t.allowDefaultRemoteLaunch == 1 && t.denyDefaultRemoteLaunch == 0) { content[6] = "允许"; }
                else if (t.allowDefaultRemoteLaunch == 0 && t.denyDefaultRemoteLaunch == 1) { content[6] = "拒绝"; }
                else if (t.allowDefaultRemoteLaunch == 0 && t.denyDefaultRemoteLaunch == 0) { content[6] = "无"; }

                if (t.allowDefaultLocalActive == 1 && t.denyDefaultLocalActive == 1) { content[7] = "允许、拒绝"; }
                else if (t.allowDefaultLocalActive == 1 && t.denyDefaultLocalActive == 0) { content[7] = "允许"; }
                else if (t.allowDefaultLocalActive == 0 && t.denyDefaultLocalActive == 1) { content[7] = "拒绝"; }
                else if (t.allowDefaultLocalActive == 0 && t.denyDefaultLocalActive == 0) { content[7] = "无"; }

                if (t.allowDefaultRemoteActive == 1 && t.denyDefaultRemoteActive == 1) { content[8] = "允许、拒绝"; }
                else if (t.allowDefaultRemoteActive == 1 && t.denyDefaultRemoteActive == 0) { content[8] = "允许"; }
                else if (t.allowDefaultRemoteActive == 0 && t.denyDefaultRemoteActive == 1) { content[8] = "拒绝"; }
                else if (t.allowDefaultRemoteActive == 0 && t.denyDefaultRemoteActive == 0) { content[8] = "无"; }

                if (t.allowRestrictionLocalLaunch == 1 && t.denyRestrictionLocalLaunch == 1) { content[9] = "允许、拒绝"; }
                else if (t.allowRestrictionLocalLaunch == 1 && t.denyRestrictionLocalLaunch == 0) { content[9] = "允许"; }
                else if (t.allowRestrictionLocalLaunch == 0 && t.denyRestrictionLocalLaunch == 1) { content[9] = "拒绝"; }
                else if (t.allowRestrictionLocalLaunch == 0 && t.denyRestrictionLocalLaunch == 0) { content[9] = "无"; }

                if (t.allowRestrictionRemoteLaunch == 1 && t.denyRestrictionRemoteLaunch == 1) { content[10] = "允许、拒绝"; }
                else if (t.allowRestrictionRemoteLaunch == 1 && t.denyRestrictionRemoteLaunch == 0) { content[10] = "允许"; }
                else if (t.allowRestrictionRemoteLaunch == 0 && t.denyRestrictionRemoteLaunch == 1) { content[10] = "拒绝"; }
                else if (t.allowRestrictionRemoteLaunch == 0 && t.denyRestrictionRemoteLaunch == 0) { content[10] = "无"; }

                if (t.allowRestrictionLocalActive == 1 && t.denyRestrictionLocalActive == 1) { content[11] = "允许、拒绝"; }
                else if (t.allowRestrictionLocalActive == 1 && t.denyRestrictionLocalActive == 0) { content[11] = "允许"; }
                else if (t.allowRestrictionLocalActive == 0 && t.denyRestrictionLocalActive == 1) { content[11] = "拒绝"; }
                else if (t.allowRestrictionLocalActive == 0 && t.denyRestrictionLocalActive == 0) { content[11] = "无"; }

                if (t.allowRestrictionRemoteActive == 1 && t.denyRestrictionRemoteActive == 1) { content[12] = "允许、拒绝"; }
                else if (t.allowRestrictionRemoteActive == 1 && t.denyRestrictionRemoteActive == 0) { content[12] = "允许"; }
                else if (t.allowRestrictionRemoteActive == 0 && t.denyRestrictionRemoteActive == 1) { content[12] = "拒绝"; }
                else if (t.allowRestrictionRemoteActive == 0 && t.denyRestrictionRemoteActive == 0) { content[12] = "无"; }
                #endregion
                //content[13] = "删除";

                dvComm.Rows.Add(content);
            }

            //OPC设置的显示界面

            dvOpc.RowCount = 0;
            foreach (opcServer ts in opclist)
            {
                if (ts.isExist == false)
                {
                    string[] content = new string[8];
                    content[0] = ts.name;
                    for (int i = 1; i < 8; i++)
                    {
                        content[i] = "不存在";
                    }

                    dvOpc.Rows.Add(content);
                    continue;
                }
                else if (ts.LaunchFlag == 0 && ts.AccessFlag == 0)
                {
                    string[] content = new string[8];
                    content[0] = ts.name;
                    for (int i = 1; i < 8; i++)
                    {
                        content[i] = "默认";
                    }

                    dvOpc.Rows.Add(content);
                    continue;
                }
                foreach (opcACL t in opcAclTable)
                {
                    if (t.opcName.Equals(ts.name))
                    {
                        string[] content = new string[8];
                        content[0] = t.opcName;
                        content[1] = t.accountName;
                        #region 显示允许还是拒绝
                        if (ts.AccessFlag == 0) content[2] = "默认";
                        else if (t.allowLocalAccess == 1 && t.denyLocalAccess == 0) content[2] = "允许";
                        else if (t.allowLocalAccess == 1 && t.denyLocalAccess == 1) content[2] = "允许、拒绝";
                        else if (t.allowLocalAccess == 0 && t.denyLocalAccess == 0) content[2] = "无";
                        else if (t.allowLocalAccess == 0 && t.denyLocalAccess == 1) content[2] = "拒绝";

                        if (ts.AccessFlag == 0) content[3] = "默认";
                        else if (t.allowRemoteAccess == 1 && t.denyRemoteAccess == 0) content[3] = "允许";
                        else if (t.allowRemoteAccess == 1 && t.denyRemoteAccess == 1) content[3] = "允许、拒绝";
                        else if (t.allowRemoteAccess == 0 && t.denyRemoteAccess == 0) content[3] = "无";
                        else if (t.allowRemoteAccess == 0 && t.denyRemoteAccess == 1) content[3] = "拒绝";

                        if (ts.LaunchFlag == 0) content[4] = "默认";
                        else if (t.allowLocalLaunch == 1 && t.denyLocalLaunch == 0) content[4] = "允许";
                        else if (t.allowLocalLaunch == 1 && t.denyLocalLaunch == 1) content[4] = "允许、拒绝";
                        else if (t.allowLocalLaunch == 0 && t.denyLocalLaunch == 0) content[4] = "无";
                        else if (t.allowLocalLaunch == 0 && t.denyLocalLaunch == 1) content[4] = "拒绝";

                        if (ts.LaunchFlag == 0) content[5] = "默认";
                        else if (t.allowRemoteLaunch == 1 && t.denyRemoteLaunch == 0) content[5] = "允许";
                        else if (t.allowRemoteLaunch == 1 && t.denyRemoteLaunch == 1) content[5] = "允许、拒绝";
                        else if (t.allowRemoteLaunch == 0 && t.denyRemoteLaunch == 0) content[5] = "无";
                        else if (t.allowRemoteLaunch == 0 && t.denyRemoteLaunch == 1) content[5] = "拒绝";

                        if (ts.LaunchFlag == 0) content[6] = "默认";
                        else if (t.allowLocalActive == 1 && t.denyLocalActive == 0) content[6] = "允许";
                        else if (t.allowLocalActive == 1 && t.denyLocalActive == 1) content[6] = "允许、拒绝";
                        else if (t.allowLocalActive == 0 && t.denyLocalActive == 0) content[6] = "无";
                        else if (t.allowLocalActive == 0 && t.denyLocalActive == 1) content[6] = "拒绝";

                        if (ts.LaunchFlag == 0) content[7] = "默认";
                        else if (t.allowRemoteActive == 1 && t.denyRemoteActive == 0) content[7] = "允许";
                        else if (t.allowRemoteActive == 1 && t.denyRemoteActive == 1) content[7] = "允许、拒绝";
                        else if (t.allowRemoteActive == 0 && t.denyRemoteActive == 0) content[7] = "无";
                        else if (t.allowRemoteActive == 0 && t.denyRemoteActive == 1) content[7] = "拒绝";
                        #endregion
                        dvOpc.Rows.Add(content);
                    }

                }
            }

        }

        private void updateCurrentName(string name) {
            //if (this.InvokeRequired)
            //{
            //    updateCurrentNameCallBack stcb = new updateCurrentNameCallBack(updateCurrentName);
            //    this.Invoke(stcb, name);
            //}
            //else
            //{
                label1.Text = name;
            //}
        }

        //0:用户名密码不能为空
        //1:创建成功
        //2:已经存在
        //-1:失败
        private int addUser(string username,string password) {
            if (username.Equals("")|| password.Equals("")) return 0;
            string[] names = userAndGroup.getAllUser();
            foreach (string name in names) {
                if (name.ToUpper().Equals(username)) return 2;
            }
            return userAndGroup.CreateLocalWindowsAccount(username, password, username, "", true, false);
        }

        //推荐设置
        private void button1_Click(object sender, EventArgs e)
        {
            string newUserAclName = "";
            if (radioCurrent.Checked == radioSpecific.Checked == radioNewAccount.Checked == false)
            {
                MessageBox.Show("请选择一个账号进行OPC通信设置");
                return;
            }
            if (radioNewAccount.Checked)//新建账号
            {
                string name = textBox1.Text;
                string passwd = textBox2.Text;
                int flag = addUser(name, passwd);
                switch (flag)
                {
                    case 0: MessageBox.Show("用户名和密码不能为空"); return;
                    case 2: MessageBox.Show("该用户已经存在"); return;
                    case -1: MessageBox.Show("创建失败"); return;
                    case 1:break;
                    default: return;
                }
                newUserAclName = getSidByName(name);
                if (newUserAclName.Equals("")) return;
            }
            else if (radioSpecific.Checked)//指定账户
            {
                if (comboBox1.SelectedIndex < 0) return;
                string selectedName = comboBox1.SelectedItem.ToString();
                newUserAclName = getSidByName(selectedName);
            }
            else//当前账户
            {
                newUserAclName = currentSID;
            }
            var commSddls = getSecurityCommSDDL(newUserAclName);

            bool flag2 = true;
            //写入注册表
            if (commSddls.Length == 4)
            {
                flag2 &= regUnit.setCommSDDL(commSddls[0], "DefaultAccessPermission");
                flag2 &= regUnit.setCommSDDL(commSddls[1], "MachineAccessRestriction");
                flag2 &= regUnit.setCommSDDL(commSddls[2], "DefaultLaunchPermission");
                flag2 &= regUnit.setCommSDDL(commSddls[3], "MachineLaunchRestriction");
            }

            flag2 &= regUnit.setCommSetting();  //连接 标识 TCP/IP 启用DCOM

            var opcSddlAccess = recommendOpcAccessSDDL(newUserAclName);
            var opcSddlLaunch = recommendOpcLaunchSDDL(newUserAclName);
            foreach (var t in opclist) {
                flag2 &= regUnit.setOpcSDDL(t.uuid, opcSddlLaunch, "LaunchPermission");//启动激活自定义
                flag2 &= regUnit.setOpcSDDL(t.uuid, opcSddlAccess, "AccessPermission");//访问自定义
                flag2 &= regUnit.setOpcSetting(t.uuid);

                //flag2 &= regUnit.deleteAppIDRegKey(t.uuid, "AccessPermission");//访问默认
            }

            update();

            if (flag2 == true)
            {
                MessageBox.Show("设置成功！");
            }
            else
            {
                MessageBox.Show("设置失败！");
            }

        }

        //默认设置
        private void button2_Click(object sender, EventArgs e)
        {
            var sddls = getDefaultCommSDDL();
            bool flag = true;
            //写入注册表
            if (sddls.Length == 4) {
                flag &= regUnit.setCommSDDL(sddls[0], "DefaultAccessPermission");
                flag &= regUnit.setCommSDDL(sddls[1], "MachineAccessRestriction");
                flag &= regUnit.setCommSDDL(sddls[2], "DefaultLaunchPermission");
                flag &= regUnit.setCommSDDL(sddls[3], "MachineLaunchRestriction");
            }
            //删除OPC下的注册表权限设置，恢复默认值
            foreach (var opc in opclist) {
                flag &= regUnit.deleteAppIDRegKey(opc.uuid, "AccessPermission");
                flag &= regUnit.deleteAppIDRegKey(opc.uuid, "LaunchPermission");
                flag &= regUnit.deleteAppIDRegKey(opc.uuid, "ActivateAtStorage");
                flag &= regUnit.deleteAppIDRegKey(opc.uuid, "RunAs");
                flag &= regUnit.deleteAppIDRegKey(opc.uuid, "RemoteServerName");
                flag &= regUnit.deleteAppIDRegKey(opc.uuid, "EndPoints");
                flag &= regUnit.deleteAppIDRegKey(opc.uuid, "AuthenticationLevel");
            }

            update();

            if (flag == true)
            {
                MessageBox.Show("设置为Microsoft默认值成功！");
            }
            else {
                MessageBox.Show("设置为Microsoft默认值成功！");
            }
        }
       
        //自定义设置
        private void button3_Click(object sender, EventArgs e)
        {
            bool everyoneFlag = false;
            for (int i = 0; i < dvComm.RowCount; i++) {
                if (dvComm.Rows[i].Cells[0].Value.ToString().ToUpper().Equals("EVERYONE")) {
                    everyoneFlag = true;
                    if (!dvComm.Rows[i].Cells[3].Value.ToString().Contains("允许")
                        || !dvComm.Rows[i].Cells[9].Value.ToString().Contains("允许")
                        || !dvComm.Rows[i].Cells[11].Value.ToString().Contains("允许")) {
                        DialogResult result = MessageBox.Show("取消Everyone的本地权限可能影响系统正常运行！", "提示信息", MessageBoxButtons.OKCancel);
                        if (result == DialogResult.OK)
                            break;
                        else
                            return;
                        
                    }
                }
            }
            if (everyoneFlag == false) {//没有everyone的行
                DialogResult result = MessageBox.Show("取消Everyone的本地权限可能影响系统正常运行！", "提示信息", MessageBoxButtons.OKCancel);
                if (result == DialogResult.OK)
                { }
                else
                    return;

            }

            bool flag = true;
            string t1 = getCustomDefaultAccessFromView();
            string t2 = getCustomRestrictionAccessFromView();
            string t3 = getCustomDefaultLaunchFromView();
            string t4 = getCustomRestrictionLaunchFromView();
            
            flag &= regUnit.setCommSDDL(t1, "DefaultAccessPermission");
            flag &= regUnit.setCommSDDL(t2, "MachineAccessRestriction");
            flag &= regUnit.setCommSDDL(t3, "DefaultLaunchPermission");
            flag &= regUnit.setCommSDDL(t4, "MachineLaunchRestriction");
            update();
            if (flag) MessageBox.Show("设置成功！");
            else MessageBox.Show("设置失败！");
        }

        //默认情况下 的通用SDDL设置
        public string[] getDefaultCommSDDL()
        {
            string DefaultAccessSDDL = "O:BAG:BAD:(A;;CCDCLC;;;BA)(A;;CCDC;;;SY)(A;;CCDCLC;;;PS)";

            string RestrictionAccessSDDL = "O:BAG:BAD:(A;;CCDC;;;AN)(A;;CCDCLC;;;LU)(A;;CCDCLC;;;WD)(A;;CCDCLC;;;S-1-5-32-562)";

            string DefaultLaunchSDDL = "O:BAG:BAD:(A;;CCDCLCSWRP;;;BA)(A;;CCDCLCSWRP;;;SY)(A;;CCDCLCSWRP;;;IU)";

            string RestrictionLaunchSDDL = "O:BAG:BAD:(A;;CCDCLCSWRP;;;BA)(A;;CCDCLCSWRP;;;LU)(A;;CCDCLCSWRP;;;S-1-5-32-562)(A;;CCDCSW;;;WD)";

            return new string[] { DefaultAccessSDDL, RestrictionAccessSDDL, DefaultLaunchSDDL, RestrictionLaunchSDDL };

        }

        //获得当前配置下的安全配置,关闭全部的远程权限，只有指定账户有远程限制权限
        public string[] getSecurityCommSDDL(string newUserSid) {
            string DefaultAccessSDDL = "O:BAG:BAD:";

            foreach (commACL t in commAclTable)
            {
                if (t.sidOrAclName.ToUpper().Equals(newUserSid.ToUpper())) continue;//指定的账号单独进行设置
                string acl = "";
                if (t.allowDefaultLocalAccess == 1 ) { acl = "CCDC"; }
                if(acl.Length>0)
                    DefaultAccessSDDL += "(A;;" + acl + ";;;" + t.sidOrAclName + ")";
            }

            DefaultAccessSDDL += "(A;;CCDC;;;" + newUserSid + ")";//指定账户的权限


            //不需要拒绝部分的设置

            string RestrictionAccessSDDL = "O:BAG:BAD:";

            foreach (commACL t in commAclTable)
            {
                if (t.sidOrAclName.ToUpper().Equals(newUserSid.ToUpper())) continue;
                if (t.displayName.ToUpper().Equals("EVERYONE")) continue;//everyone账号单独进行设置
                string acl = "";
                if (t.allowRestrictionLocalAccess == 1) acl = "CCDC";
                RestrictionAccessSDDL += "(A;;" + acl + ";;;" + t.sidOrAclName + ")";
            }

            RestrictionAccessSDDL += "(A;;CCDC;;;WD)";//everyone的本地权限，防止无法使用系统
            RestrictionAccessSDDL += "(A;;CCDCLC;;;" + newUserSid + ")";//指定账户的权限


            string DefaultLaunchSDDL = "O:BAG:BAD:";

            foreach (commACL t in commAclTable)
            {
                if (t.sidOrAclName.ToUpper().Equals(newUserSid.ToUpper())) continue;
                string acl = "";
                if (t.allowDefaultLocalLaunch == 1) acl += "DC";
                if (t.allowDefaultLocalActive == 1) acl += "SW";

                if (acl.Length > 0)
                {
                    DefaultLaunchSDDL += "(A;;CC" + acl + ";;;" + t.sidOrAclName + ")";
                }
            }
            DefaultLaunchSDDL += "(A;;CCDCSW;;;" + newUserSid + ")";//指定账户的权限




            string RestrictionLaunchSDDL = "O:BAG:BAD:";

            foreach (commACL t in commAclTable)
            {
                if (t.sidOrAclName.ToUpper().Equals(newUserSid.ToUpper())) continue;
                string acl = "";
                if (t.allowRestrictionLocalLaunch == 1) acl += "DC";
                if (t.allowRestrictionLocalActive == 1) acl += "SW";

                if (acl.Length > 0)
                {
                    RestrictionLaunchSDDL += "(A;;CC" + acl + ";;;" + t.sidOrAclName + ")";
                }
            }

            RestrictionLaunchSDDL += "(A;;CCDCSW;;;" + newUserSid + ")";
            RestrictionLaunchSDDL += "(A;;CCDCLCSWRP;;;" + newUserSid + ")";

            

            return new string[] { DefaultAccessSDDL, RestrictionAccessSDDL, DefaultLaunchSDDL, RestrictionLaunchSDDL };


        }

        //指定账户，SY,IU有权限
        public string recommendOpcLaunchSDDL(string newUserSid) {
            string sddl = "O:BAG:BAD:(A;;CCDCLCSWRP;;;SY)(A;;CCDCLCSWRP;;;IU)";
            sddl += "A;;CCDCLCSWRP;;;" + newUserSid + ")";
            return sddl;
        }

        public string recommendOpcAccessSDDL(string newUserSid)
        {
            string sddl = "O:BAG:BAD:(A;;CCDCLC;;;SY)(A;;CCDCLC;;;IU)";
            sddl += "A;;CCDCLC;;;" + newUserSid + ")";
            return sddl;
        }

        //从datagridview返回sddl string
        private string getCustomDefaultAccessFromView() {
            string DefaultAccessSDDL = "O:BAG:BAD:";
            for (int i = 0; i < dvComm.RowCount; i++) {
                string acl = "";
                string aclname = getSidByName(dvComm.Rows[i].Cells[0].Value.ToString());
                string tmp1 = dvComm.Rows[i].Cells[1].Value.ToString();//DC
                string tmp2 = dvComm.Rows[i].Cells[2].Value.ToString();//LC
                if (tmp1.Contains("允许")  && tmp2.Contains("允许") ) { acl += "(A;;CCDCLC;;;" + aclname + ")"; }
                else if (tmp1.Contains("允许")  && !tmp2.Contains("允许")) { acl += "(A;;CCDC;;;" + aclname + ")"; }
                else if (!tmp1.Contains("允许") && tmp2.Contains("允许")) { acl += "(A;;CCLC;;;" + aclname + ")"; }
                else if (!tmp1.Contains("允许") && !tmp2.Contains("允许")) {  }

                if (tmp1.Contains("拒绝") && tmp2.Contains("拒绝")) { acl += "(D;;CCDCLC;;;" + aclname + ")"; }
                else if (tmp1.Contains("拒绝") && !tmp2.Contains("拒绝")) { acl += "(D;;CCDC;;;" + aclname + ")"; }
                else if (!tmp1.Contains("拒绝") && tmp2.Contains("拒绝")) { acl += "(D;;CCLC;;;" + aclname + ")"; }
                else if (!tmp1.Contains("拒绝") && !tmp2.Contains("拒绝")) { }
                if (acl.Length > 0) DefaultAccessSDDL += acl;

            }
            return DefaultAccessSDDL;
        }

        private string getCustomRestrictionAccessFromView()
        {
            string RestrictionAccessSDDL = "O:BAG:BAD:";
            for (int i = 0; i < dvComm.RowCount; i++)
            {
                string acl = "";
                string aclname = getSidByName(dvComm.Rows[i].Cells[0].Value.ToString());
                string tmp1 = dvComm.Rows[i].Cells[3].Value.ToString();//DC
                string tmp2 = dvComm.Rows[i].Cells[4].Value.ToString();//LC
                if (tmp1.Contains("允许") && tmp2.Contains("允许")) { acl += "(A;;CCDCLC;;;" + aclname + ")"; }
                else if (tmp1.Contains("允许") && !tmp2.Contains("允许")) { acl += "(A;;CCDC;;;" + aclname + ")"; }
                else if (!tmp1.Contains("允许") && tmp2.Contains("允许")) { acl += "(A;;CCLC;;;" + aclname + ")"; }
                else if (!tmp1.Contains("允许") && !tmp2.Contains("允许")) { }

                if (tmp1.Contains("拒绝") && tmp2.Contains("拒绝")) { acl += "(D;;CCDCLC;;;" + aclname + ")"; }
                else if (tmp1.Contains("拒绝") && !tmp2.Contains("拒绝")) { acl += "(D;;CCDC;;;" + aclname + ")"; }
                else if (!tmp1.Contains("拒绝") && tmp2.Contains("拒绝")) { acl += "(D;;CCLC;;;" + aclname + ")"; }
                else if (!tmp1.Contains("拒绝") && !tmp2.Contains("拒绝")) { }
                if (acl.Length > 0) RestrictionAccessSDDL += acl;

            }
            return RestrictionAccessSDDL;
        }

        private string getCustomDefaultLaunchFromView()
        {
            string DefaultLaunchSDDL = "O:BAG:BAD:";
            for (int i = 0; i < dvComm.RowCount; i++)
            {
                string acl = "";
                string aclname = getSidByName(dvComm.Rows[i].Cells[0].Value.ToString());
                string tmp1 = dvComm.Rows[i].Cells[5].Value.ToString();//DC
                string tmp2 = dvComm.Rows[i].Cells[6].Value.ToString();//LC
                string tmp3 = dvComm.Rows[i].Cells[7].Value.ToString();//SW
                string tmp4 = dvComm.Rows[i].Cells[8].Value.ToString();//RP
                string t = "";
                if (tmp1.Contains("允许")) t += "DC";
                if (tmp2.Contains("允许")) t += "LC";
                if (tmp3.Contains("允许")) t += "SW";
                if (tmp4.Contains("允许")) t += "RP";
                if (t.Length > 0) acl += "(A;;CC" + t + ";;;" + aclname + ")";
                t = "";
                if (tmp1.Contains("拒绝")) t += "DC";
                if (tmp2.Contains("拒绝")) t += "LC";
                if (tmp3.Contains("拒绝")) t += "SW";
                if (tmp4.Contains("拒绝")) t += "RP";
                if (t.Length > 0) acl += "(D;;CC" + t + ";;;" + aclname + ")";
                if (acl.Length > 0) DefaultLaunchSDDL += acl;
                
            }
            return DefaultLaunchSDDL;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            radioSpecific.Select();
        }

        private string getCustomRestrictionLaunchFromView()
        {
            string RestrictionLaunchSDDL = "O:BAG:BAD:";
            for (int i = 0; i < dvComm.RowCount; i++)
            {
                string acl = "";
                string aclname = getSidByName(dvComm.Rows[i].Cells[0].Value.ToString());
                string tmp1 = dvComm.Rows[i].Cells[9].Value.ToString();//DC
                string tmp2 = dvComm.Rows[i].Cells[10].Value.ToString();//LC
                string tmp3 = dvComm.Rows[i].Cells[11].Value.ToString();//SW
                string tmp4 = dvComm.Rows[i].Cells[12].Value.ToString();//RP
                string t = "";
                if (tmp1.Contains("允许")) t += "DC";
                if (tmp2.Contains("允许")) t += "LC";
                if (tmp3.Contains("允许")) t += "SW";
                if (tmp4.Contains("允许")) t += "RP";
                if (t.Length > 0) acl += "(A;;CC" + t + ";;;" + aclname + ")";
                t = "";
                if (tmp1.Contains("拒绝")) t += "DC";
                if (tmp2.Contains("拒绝")) t += "LC";
                if (tmp3.Contains("拒绝")) t += "SW";
                if (tmp4.Contains("拒绝")) t += "RP";
                if (t.Length > 0) acl += "(D;;CC" + t + ";;;" + aclname + ")";
                if (acl.Length > 0) RestrictionLaunchSDDL += acl;

            }
            return RestrictionLaunchSDDL;
        }

        //优先返回 LA 而不是s-1-1-500
        public string getSidByName(string name) {
            foreach (mSID t in insideSidList) {
                if (name.ToUpper().Equals(t.name.ToUpper())) {
                    return t.sid;
                }
            }
            string sid = userAndGroup.getSidByName(name);
            if (sid == null || sid.Equals("")) return name;//本地没有读取到该账号。在安全描述符里，内置了一些奇怪的sid账号，这些账号找不到，直接返回
            return sid;
        }


    }
}
















