
using System;
using System.Data;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using CAdapp = Autodesk.AutoCAD.ApplicationServices.Application;
using RibbonControl = Autodesk.Windows.RibbonControl;
/*AutoCAD*/
using AcAdAmComm;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.BoundaryRepresentation;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Internal.DatabaseServices;
using Autodesk.AutoCAD.Customization;
using Autodesk.AutoCAD.Colors;
using Autodesk.Windows;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Collections.Generic;


namespace SESDTools
{
    /// <summary>
    /// 类    名：GlobalVariableController
    /// 功能描述：全局变量类
    /// 创 建 人：SY
    /// 创建时间：2021/07/23 10:43
    /// 描    述：将画面上的数据保存
    /// 修 改 人：
    /// 修改时间：
    /// 修改描述：
    /// 版    本：V1.0
    /// </summary>
    public class GlobalVariableController
    {
        #region 按钮全局变量
        //按钮全局变量
        public static bool[] GlobalButtonFlag = { true, false, false, false, false, false, false, false, true, true, true };
        #endregion

        #region 数据全局变量
        //数据全局变量
        public static DataSet GlobalDataSet = new DataSet();
        #endregion
    }
    public class PCSPartList
    {
        //private string _partno;//Part NO 零件号
        //private string _des;//Description 名称
        //private string _material;//Material 材质
        //private string _qty;//Quantity 数量
        //public PCSPartList(string PartNO, string DES, string Material, string QTY)
        //{
        //    this._partno = PartNO;
        //    this._des = DES;
        //    this._material = Material;
        //    this._qty = QTY;
        //}
        public string PartNO {  get; set; } 
        public string DES { get; set; }
        public string Material { get; set; } 
        public string QTY { get; set; }

    }
    
    /// <summary>
    /// 类    名：HomeController
    /// 功能描述：添加自定义菜单
    /// 创 建 人：SY
    /// 创建时间：2021/07/23 10:43
    /// 描    述：将自定义菜单图标添加到AutoCAD界面上
    /// 修 改 人：
    /// 修改时间：
    /// 修改描述：
    /// 版    本：V1.0
    /// </summary>

    public class HomeController
    {
        private static String FlangeStd; //定义字符串 泵法兰标准/Pump Flange Standard
        private static int LoopC; //定义整型 程序循环次数
        private static Point3d OCPoint = new Point3d(0, 0, 0); //定义3D坐标点 坐标原点
        private static string Savename; //定义字符串 保存文件名称
        private static string CustomerName; //定义字符串 用户名称
        private static string Project; //定义字符串 项目名称
        private static string SulzerNO; //定义字符串 Sulzer订单号
        private static string SDate;//定义字符串 日期
        private static string PI, PO, PD, PBD, WI, FI, FO, CI, CO, II, IO, ID, QI, QO, GBI, GBO;
        private static string RatingPI, RatingPO, RatingPD, RatingPBD, RatingWI, RatingFI, RatingFO, RatingCI, RatingCO, RatingII, RatingIO, RatingID, RatingQI, RatingQO, RatingGBI, RatingGBO;
        private static string FacingPI, FacingPO, FacingPD, FacingPBD, FacingWI, FacingFI, FacingFO, FacingCI, FacingCO, FacingII, FacingIO, FacingID, FacingQI, FacingQO, FacingGBI, FacingGBO;
        private static string MATPUMP, MATFLUSHING, MATCOOLING, MATSTEAMJACKET, MATQUENCH, MATBUFFERGAS;
        //DataRow 4 Z Partlist
        public static DataRow[] newRPartListZ;
        public static DataRow[] newRZMaterial1;
        public static DataRow[] newRZMaterial2;
        public static DataRow[] newRZMaterial3;
        //DataRow 4 PRE PartList
        public static DataRow[] newRPCSPRE1;
        public static DataRow[] newRPCSPRE2;
        public static DataRow[] newRPCSPRE3;
        public static DataRow[] newRPCSPRE4;
        public static DataRow[] newRPCSPRE5;
        #region 生成CAD自定义菜单按钮界面方法
        /// <summary>
        /// 生成CAD自定义菜单按钮界面方法
        /// </summary>
        [CommandMethod("CreateMyCuix")]
        public static void CreateMyCuix()
        {
            //存放自定义文件路径
            string iconFileFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\SESDIcon\\";
            //小图标文件
            string smallButtonIconFile = iconFileFolder + "SmallButtonIcon.bmp";
            //大图标文件
            string largeButtonIcon = iconFileFolder + "LargeButtonIcon.bmp";
            var editor = CAdapp.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                var customizationSection = new CustomizationSection();
                var menuGroup = customizationSection.MenuGroup;
                var newTab = customizationSection.AddNewTab("Sulzer Tools", null);
                //READDATA Button
                var panelReadData = newTab.AddNewPanel("Read Data", null);
                var rowReadData = panelReadData.AddNewRibbonRow();
                //PGA Button
                var panelPGA = newTab.AddNewPanel("Pga", null);
                var rowPGA = panelPGA.AddNewRibbonRow();
                //PPI Button
                var panelPPI = newTab.AddNewPanel("Ppi", null);
                var rowPPI = panelPPI.AddNewRibbonRow();
                //PCS Button
                var panelPCS = newTab.AddNewPanel("Pcs", null);
                var rowPCS = panelPCS.AddNewRibbonRow();
                //PUC Button
                var panelPUC = newTab.AddNewPanel("Puc", null);
                var rowPUC = panelPUC.AddNewRibbonRow();
                //PNT Button
                var panelPNT = newTab.AddNewPanel("Pnt", null);
                var rowPNT = panelPNT.AddNewRibbonRow();
                //PDS Button
                var panelPDS = newTab.AddNewPanel("Pds", null);
                var rowPDS = panelPDS.AddNewRibbonRow();
                //PSL Button
                var panelPSL = newTab.AddNewPanel("Psl", null);
                var rowPSL = panelPSL.AddNewRibbonRow();
                //VER Button
                var panelVER = newTab.AddNewPanel("Ver", null);
                var rowVER = panelVER.AddNewRibbonRow();
                //ADMINI Button
                var panelAdmini = newTab.AddNewPanel("Admin", null);
                var rowAdmini = panelAdmini.AddNewRibbonRow();
                //ADMINIINFOR Button
                var panelAdminiInfor = newTab.AddNewPanel("Admin Infor", null);
                var rowAdminiInfor = panelAdminiInfor.AddNewRibbonRow();
                menuGroup.Name = "SulzerTools";
                //添加自定义功能按钮
                rowReadData.AddNewButton("Data", "Read Data", "SumListForm", "Data",
                    smallButtonIconFile, largeButtonIcon, RibbonButtonStyle.LargeWithText);
                rowPGA.AddNewButton("Pga", "Pga", "PGA", "Pga",
                    smallButtonIconFile, largeButtonIcon, RibbonButtonStyle.LargeWithText);
                rowPPI.AddNewButton("Ppi", "Ppi", "PPI", "Ppi",
                    smallButtonIconFile, largeButtonIcon, RibbonButtonStyle.LargeWithText);
                rowPCS.AddNewButton("Pcs", "Pcs", "PCS", "Pcs",
                    smallButtonIconFile, largeButtonIcon, RibbonButtonStyle.LargeWithText);
                rowPUC.AddNewButton("Puc", "Puc", "PUC", "Puc",
                    smallButtonIconFile, largeButtonIcon, RibbonButtonStyle.LargeWithText);
                rowPNT.AddNewButton("Pnt", "Pnt", "PNT", "Pnt",
                    smallButtonIconFile, largeButtonIcon, RibbonButtonStyle.LargeWithText);
                rowPDS.AddNewButton("Pds", "Pds", "PDS", "Pds",
                    smallButtonIconFile, largeButtonIcon, RibbonButtonStyle.LargeWithText);
                rowPSL.AddNewButton("Psl", "Psl", "PSL", "Psl",
                    smallButtonIconFile, largeButtonIcon, RibbonButtonStyle.LargeWithText);
                rowVER.AddNewButton("Ver", "Ver", "VersionInformationForm", "Ver",
                    smallButtonIconFile, largeButtonIcon, RibbonButtonStyle.LargeWithText);
                rowAdmini.AddNewButton("Admin", "Admin", "PasswordConfirmationForm", "Admin",
                    smallButtonIconFile, largeButtonIcon, RibbonButtonStyle.LargeWithText);
                rowAdminiInfor.AddNewButton("Admin Infor", "Admin Infor", "AdministratorInformationForm", "Admin Infor",
                    smallButtonIconFile, largeButtonIcon, RibbonButtonStyle.LargeWithText);
                //保存新用户界面文件
                var fileName = Path.Combine(iconFileFolder, "SulzerToolCui.cuix");
                File.Delete(fileName);
                customizationSection.SaveAs(fileName);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage(Environment.NewLine + ex.Message);
            }
        }
        #endregion

        #region 按钮初始化
        /// <summary>
        /// 按钮初始化
        /// </summary>
        public void ButtonInitialize()
        {
            for (int i = 0; i < GlobalVariableController.GlobalButtonFlag.Length; i++)
            {
                if (i == 0)
                {
                    GlobalVariableController.GlobalButtonFlag[0] = true;
                }
                else
                {
                    GlobalVariableController.GlobalButtonFlag[i] = false;
                }
                //后3个按钮可亮
                if (i == GlobalVariableController.GlobalButtonFlag.Length - 3)
                {
                    GlobalVariableController.GlobalButtonFlag[i] = true;
                }
                if (i == GlobalVariableController.GlobalButtonFlag.Length - 2)
                {
                    GlobalVariableController.GlobalButtonFlag[i] = true;
                }
                if (i == GlobalVariableController.GlobalButtonFlag.Length - 1)
                {
                    GlobalVariableController.GlobalButtonFlag[i] = true;
                }
            }
        }
        #endregion

        #region 更新CAD自定义菜单按钮
        [CommandMethod("ButtonUpdata")]
        public void ButtonUpdata()
        {
            RibbonControl control = ComponentManager.Ribbon;//获取Ribbon界面
            RibbonTabCollection tabs = control.Tabs;
            RibbonTab tab = new RibbonTab();
            for (int i = 0; i < tabs.Count; i++)
            {
                if (tabs[i].AutomationName == "Sulzer Tools")
                {
                    tab = tabs[i];
                    break;
                }
            }
            //RibbonPanel Panel1 = new RibbonPanel();
            string[] pnlid = { "pnlRead Data", "pnlPga", "pnlPpi", "pnlPcs", "pnlPuc", "pnlPnt", "pnlPds", "pnlPsl", "pnlVer", "pnlAdmin", "pnlAdmin Info" };
            if (pnlid.Length == tab.Panels.Count)
            {
                for (int i = 0; i < pnlid.Length; i++)
                {
                    for (int j = 0; j < tab.Panels.Count; j++)
                    {
                        if (tab.Panels[j].UID == pnlid[i])
                        {
                            tab.Panels[j].IsEnabled = GlobalVariableController.GlobalButtonFlag[i];
                            break;
                        }
                    }
                }
            }
        }
        #endregion

        [CommandMethod("PGA")]
        public void PGA()
        {
            SelectPathForm form = new SelectPathForm("PGA", GlobalVariableController.GlobalDataSet);
            form.ShowDialog();

        }
        [CommandMethod("PPI")]
        public void PPI()
        {
            Document DocPPI = Application.DocumentManager.MdiActiveDocument;
            Editor EdPPI = Application.DocumentManager.MdiActiveDocument.Editor;
            string str = GlobalVariableController.GlobalDataSet.Tables[1].Rows[0].ItemArray[4].ToString();
            //define savepath
            SelectPathForm form = new SelectPathForm("PPI", GlobalVariableController.GlobalDataSet);
            form.ShowDialog();
            LoopC = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows.Count;
            CustomerName = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[8].ToString();
            Project = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[9].ToString();
            SulzerNO = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[12].ToString();
            FlangeStd = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[17].ToString();
            SDate = string.Format("{0:yyyy/MM/dd}",DateTime.Parse( form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[16].ToString()));
            int PumpQTY = 0;
            string FlangeFlag = form.BASICDATASAT.Tables["FLANGESTANDARDTABLE"].Select("FLANGESTANDARD = '" + FlangeStd + "' ")[0].Field<string>("FLAG").ToString();
            PI = PO = PD = PBD = WI = FI = FO = CI = CO = II = IO = ID = QI = QO = GBI = "";
            Dictionary<string, string> DicPPIDes = new Dictionary<string, string>();
            Dictionary<string, string> DicPumpTableName = new Dictionary<string, string>();
            string[] ArrPumptype = { "ZA", "ZE", "ZF", "PRE", "PRELF", "PRESF", "PREK" };
            string[] TableName = { "SD1.1-ZATABLE", "SD1.2-ZETABLE", "SD1.3-ZFTABLE", "SD1.4-PRETABLE", "SD1.5-PRELFTABLE", "SD1.6-PRESFTABLE", "SD1.7-PREKTABLE" };
            string[] ArrItem = { "PI", "PO", "PD", "WI", "PBD", "FI", "FO", "CI", "CO", "II", "IO", "ID", "QI", "QO", "GBI", "GBO" };
            string[] ArrDes ={ "Pump suction nozzle ",  "Pump discharge nozzle ",  "Pump drain ",
                     "Pump Warmup Inlet",  "Baseplated drain",  "Flushing Inlet",
                      "Flushing Outlet",  "Water cooling inlet",  "Water cooling outlet",
                      "Insulation medium inlet ",  "Insulation medium outlet",  "Insulation medium Drain",
                      "Quench Inlet",  "Quench Outlet",  "Gas buffer Inlet",  "Gas buffer Outlet"};
            //string[] ArrDes ={ "泵入口/Pump suction nozzle ",  "泵出口/Pump discharge nozzle ",  "泵排残液接口/Pump drain ",
            //         "泵暖泵入口/Pump Warmup Inlet",  "泵底座排残液接口/Baseplated drain",  "密封冲洗入口/Flushing Inlet",
            //          "密封冲洗出口/Flushing Outlet",  "冷却水管路入口/Water cooling inlet",  "冷却水管路出口/Water cooling outlet",
            //          "保温介质入口/Insulation medium inlet ",  "保温介质出口/Insulation medium outlet",  "保温介质排残/Insulation medium Drain",
            //          "急冷液入口/Quench Inlet",  "急冷液出口/Quench Outlet",  "隔离气进口/ Gas buffer Inlet",  "隔离气出口/ Gas buffer Outlet"};
            for (int CDes = 0; CDes <= 15; CDes++) { DicPPIDes.Add(ArrItem[CDes], ArrDes[CDes]); }
            for (int Ctbname = 0; Ctbname <= TableName.Length - 1; Ctbname++) { DicPumpTableName.Add(ArrPumptype[Ctbname], TableName[Ctbname]); }
            for (int i = 0; i <= LoopC - 1; i++)
            {
                Database workingDB = HostApplicationServices.WorkingDatabase;//定义数据库文件，并指向当前活动文件
                Database db = new Database(false, true);//定义新数据库文件
                SortedDictionary<int, string> DicPPIConnlist = new SortedDictionary<int, string>();
                Dictionary<string, string> DicPPISize = new Dictionary<string, string>();
                Dictionary<string, string> DicPPIRating = new Dictionary<string, string>();
                Dictionary<string, string> DicPPIFacing = new Dictionary<string, string>();
                Dictionary<string, string> DicPPIMaterial = new Dictionary<string, string>();
                Dictionary<string, string> DicIDList = new Dictionary<string, string>();
                //get value from sumarylist
                string BearingType = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColBearing");
                string PDSize = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColPumpDrainNozzleSize");
                Boolean MatFlange =Boolean.Parse(form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i]["ColCompanionFlg"].ToString());
                Boolean Pwarmup = Boolean.Parse(form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColWarmUpPiping").ToString());
                String Pumptype = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<String>("ColType");
                String PREPressure = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<String>("ColPREPumpPressureClass");
                int.TryParse(form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<String>("ColQty"), out PumpQTY);
                string  PumpSize = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColSize").Split('(')[0].Replace(Pumptype, "").ToString();
                string Services = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColService");
                string ItemNO = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColItemNo");
                //Boolean PureMist = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<Boolean>("ColReservedPureLubrication");
                //int APInozzle = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<int>("ColNozzleLoad");
                String API682Ver = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<String>("ColApiEdition");
                string PlanA = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColSealPlanA");
                string PlanB = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColSealPlanB");
                string PlanC = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColSealPlanC");
                string MaterialClass = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColMaterialClass");
                Boolean CaseBrket = Boolean.Parse(form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i]["ColCaseBracket"].ToString());
                Boolean BRHDrain = Boolean.Parse(form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i]["ColBearighousingDrain"].ToString()); 
                int PDValve = int.Parse(form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColPumpDrainValve").ToString());
                string SerialNO = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColSeriesNo");
                string IDPump = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColTagNo");
                string ServiceName = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColService");
                Savename = "PPI" + SulzerNO + ("00" + ItemNO).ToString().Substring(("00" + ItemNO).ToString().Length - 4, 4) + "-01" + "_Ver0";
                RatingGBO = "NPTF With Plug"; GBO = "1/2" + '"';
                //detect cooling piping with single line or more
                int PlanM = 0;
                if (BearingType.Contains("S")) { PlanM = PlanM + 1; }
                if (PlanA.Contains("21") || PlanA.Contains("23")) { PlanM = PlanM + 1; }
                if (PlanB.Contains("52") || PlanB.Contains("53"))
                {
                    PlanM = PlanM + 1;
                    //redifine API682 Plan
                    if (API682Ver.Contains("3TH"))
                    {
                        PlanB = PlanB + "-3rd";
                    }
                    else
                if (API682Ver.Contains("4TH"))
                    {
                        PlanB = PlanB + "-4rd";
                    }
                }
                //detect the blocks involved
                //basic CONN
                for (int iConnlist = 0; iConnlist <= 2; iConnlist++) { DicPPIConnlist.Add(iConnlist + 1, ArrItem[iConnlist]); }
                if (Pwarmup == true) { DicPPIConnlist.Add(5, "PBD"); }
                //other CONN
                if (Pwarmup == true) { DicPPIConnlist.Add(4, "WI"); }
                if (PlanA.Contains("32") || PlanB.Contains("32")) { DicPPIConnlist.Add(6, "FI"); }
                if (PlanB.Contains("54")) { DicPPIConnlist.Add(6, "FI"); DicPPIConnlist.Add(7, "FO"); }
                if (PlanM > 0) { DicPPIConnlist.Add(8, "CI"); DicPPIConnlist.Add(9, "CO"); }
                if (CaseBrket == true) { DicPPIConnlist.Add(10, "II"); DicPPIConnlist.Add(11, "IO"); DicPPIConnlist.Add(12, "ID"); }
                if (PlanC.Contains("62")) { DicPPIConnlist.Add(13, "QI"); DicPPIConnlist.Add(14, "QO"); }
                if (PlanC.Contains("72") || PlanC.Contains("74")) { DicPPIConnlist.Add(15, "GBI"); DicPPIConnlist.Add(16, "GBO"); }
                //Get Connlist matingflange material from BASICDATA
                string expMAT = "MATERIALCLASS = '" + MaterialClass + "' ";
                DataRow[] newRMAT;
                newRMAT = form.BASICDATASAT.Tables["MATFLANGEMATERIALTABLE"].Select(expMAT);
                //Set Connlist matingflange material
                MATPUMP = newRMAT[0].Field<string>("PUMP");
                MATFLUSHING = newRMAT[0].Field<string>("FLUSHING");
                MATCOOLING = newRMAT[0].Field<string>("COOLING");
                MATSTEAMJACKET = newRMAT[0].Field<string>("STEAMJACKET");
                MATQUENCH = newRMAT[0].Field<string>("QUENCH");
                MATBUFFERGAS = newRMAT[0].Field<string>("BUFFERGAS");
                //Set Basic Connection size
                string expSD = "PumpSize = '" +  PumpSize + "' ";
                string TBnamePtype = "";
                if (Pumptype.Contains("ZA")) { TBnamePtype = "ZA"; }
                if (Pumptype.Contains("ZE")) { TBnamePtype = "ZE"; }
                if (Pumptype.Contains("PR")) { TBnamePtype = Pumptype; }
                DataRow[] newRSD;
                newRSD = form.BASICDATASAT.Tables[DicPumpTableName[TBnamePtype]].Select(expSD);
                switch (FlangeFlag)
                {
                    case "NPS":
                        PI = newRSD[0].Field<string>(FlangeFlag + "s");
                        PO = newRSD[0].Field<string>(FlangeFlag + "d");
                        PD = newRSD[0].Field<string>(FlangeFlag + "dr");
                        PI = PD;
                        if (Pumptype.Contains("ZA"))
                        {
                            PBD = "1" + '"';
                        }
                        else
                        {
                            PBD = "2" + '"';
                        }
                        II = IO = ID = QI = QO = GBI = "1/2" + '"';
                        if (FlangeStd.Contains("3406"))
                        {
                            if (MatFlange == true)
                            {
                                RatingPBD = "N/A";
                                if (Pumptype.Contains("ZA"))
                                {
                                    RatingPI = RatingPO = RatingPD = RatingWI = RatingFI = RatingFO = RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = RatingGBI = "PN20" + " SCH40";
                                }
                                if (Pumptype.Contains("ZE"))
                                {
                                    RatingPI = RatingPO = RatingPD = RatingWI = RatingFI = RatingFO = RatingGBI = "PN50" + " SCH80";
                                    RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = "PN20" + " SCH40";
                                }
                                if (Pumptype.Contains("ZF"))
                                {
                                    RatingPI = RatingPO = RatingPD = RatingWI = RatingFI = RatingFO = RatingGBI = "PN110" + " SCH80";
                                    RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = "PN20" + " SCH40";

                                }
                                if (Pumptype.Contains("PRE"))
                                {
                                    if (PREPressure == "PRE-CL300") { RatingPI = RatingPO = RatingPD = RatingWI = "PN50"; }
                                    else if (PREPressure == "PRE-CL600") { RatingPI = RatingPO = RatingPD = RatingWI = "PN110"; }
                                    RatingFI = RatingFO = RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = RatingGBI = "PN20";
                                }
                            }
                            if (MatFlange == false)
                            {
                                RatingPBD = "NPTF";
                                if (Pumptype.Contains("ZA"))
                                {
                                    RatingPI = RatingPO = RatingPD = RatingWI = RatingFI = RatingFO = RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = RatingGBI = "PN20";
                                }
                                if (Pumptype.Contains("ZE"))
                                {
                                    RatingPI = RatingPO = RatingPD = RatingWI = "PN50";
                                    RatingFI = RatingFO = RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = RatingGBI = "PN20";
                                }
                                if (Pumptype.Contains("ZF"))
                                {
                                    RatingPI = RatingPO = RatingPD = RatingWI = "PN110";
                                    RatingFI = RatingFO = RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = RatingGBI = "PN20";
                                }
                                if (Pumptype.Contains("PRE"))
                                {
                                    if (PREPressure == "PRE-CL300") { RatingPI = RatingPO = RatingPD = RatingWI = "PN50"; }
                                    else if (PREPressure == "PRE-CL600") { RatingPI = RatingPO = RatingPD = RatingWI = "PN110"; }
                                    RatingFI = RatingFO = RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = RatingGBI = "PN20";
                                }
                            }
                        }
                        else
                        {
                            if (MatFlange == true)
                            {
                                RatingPBD = "N/A";
                                if (Pumptype.Contains("ZA"))
                                {
                                    RatingPI = RatingPO = RatingPD = RatingWI = RatingFI = RatingFO = RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = RatingGBI = "150#" + " SCH40";
                                }
                                if (Pumptype.Contains("ZE"))
                                {
                                    RatingPI = RatingPO = RatingPD = RatingWI = RatingFI = RatingFO = RatingGBI = "300#" + " SCH80";
                                    RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = "150#" + " SCH40";
                                }
                                if (Pumptype.Contains("ZF"))
                                {
                                    RatingPI = RatingPO = RatingPD = RatingWI = RatingFI = RatingFO = RatingGBI = "600#" + " SCH80";
                                    RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = "150#" + " SCH40";

                                }
                                if (Pumptype.Contains("PRE"))
                                {
                                    if (PREPressure == "PRE-CL300") { RatingPI = RatingPO = RatingPD = RatingWI = "300#"; }
                                    else if (PREPressure == "PRE-CL600") { RatingPI = RatingPO = RatingPD = RatingWI = "600#"; }
                                    RatingFI = RatingFO = RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = RatingGBI = "150#";
                                }
                            }
                            if (MatFlange == false)
                            {
                                RatingPBD = "NPTF";
                                if (Pumptype.Contains("ZA"))
                                {
                                    RatingPI = RatingPO = RatingPD = RatingWI = RatingFI = RatingFO = RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = RatingGBI = "150#";
                                }
                                if (Pumptype.Contains("ZE"))
                                {
                                    RatingPI = RatingPO = RatingPD = RatingWI = "300#";
                                    RatingFI = RatingFO = RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = RatingGBI = "150#";
                                }
                                if (Pumptype.Contains("ZF"))
                                {
                                    RatingPI = RatingPO = RatingPD = RatingWI = "600#";
                                    RatingFI = RatingFO = RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = RatingGBI = "150#";
                                }
                                if (Pumptype.Contains("PRE"))
                                {
                                    if (PREPressure == "PRE-CL300") { RatingPI = RatingPO = RatingPD = RatingWI = "300#"; }
                                    else if (PREPressure == "PRE-CL600") { RatingPI = RatingPO = RatingPD = RatingWI = "600#"; }
                                    RatingFI = RatingFO = RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = RatingGBI = "150#";
                                }
                            }
                        }
                        break;
                    case "DN":
                        PI = "DN" + newRSD[0].Field<string>(FlangeFlag + "s").Replace("DN", "");
                        PO = "DN" + newRSD[0].Field<string>(FlangeFlag + "d").Replace("DN", "");
                        PD = "DN" + newRSD[0].Field<string>(FlangeFlag + "dr").Replace("DN", "");
                        PI = PD;
                        II = IO = ID = QI = QO = GBI = "DN15";
                        if (MatFlange == true)
                        {
                            RatingPBD = "N/A";
                            if (Pumptype.Contains("ZA"))
                            {
                                RatingPI = RatingPO = RatingPD = RatingWI = RatingFI = RatingFO = RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = RatingGBI = "PN25" + " N/A";
                            }
                            if (Pumptype.Contains("ZE"))
                            {
                                RatingPI = RatingPO = RatingPD = RatingWI = RatingFI = RatingFO = RatingGBI = "PN40" + " N/A";
                                RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = "PN25" + " N/A";
                            }
                            if (Pumptype.Contains("ZF"))
                            {
                                RatingPI = RatingPO = RatingPD = RatingWI = RatingFI = RatingFO = RatingGBI = "PN100" + " N/A";
                                RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = "PN25" + " N/A";

                            }
                            if (Pumptype.Contains("PRE"))
                            {
                                if (PREPressure == "PRE-CL300") { RatingPI = RatingPO = RatingPD = RatingWI = "PN40" + " N/A"; }
                                else if (PREPressure == "PRE-CL600") { RatingPI = RatingPO = RatingPD = RatingWI = "PN100" + " N/A"; }
                                        RatingFI = RatingFO = RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = RatingGBI = "PN25" + " N/A";
                            }
                        }
                        if (MatFlange == false)
                        {
                            RatingPBD = "NPTF";
                            if (Pumptype.Contains("ZA"))
                            {
                                RatingPI = RatingPO = RatingPD = RatingWI = RatingFI = RatingFO = RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = RatingGBI = "PN25";
                            }
                            if (Pumptype.Contains("ZE"))
                            {
                                RatingPI = RatingPO = RatingPD = RatingWI = "PN40";
                                RatingFI = RatingFO = RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = RatingGBI = "PN25";
                            }
                            if (Pumptype.Contains("ZF"))
                            {
                                RatingPI = RatingPO = RatingPD = RatingWI = "PN100";
                                RatingFI = RatingFO = RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = RatingGBI = "PN25";
                            }
                            if (Pumptype.Contains("PRE"))
                            {
                                if (PREPressure == "PRE-CL300") { RatingPI = RatingPO = RatingPD = RatingWI = "PN40"; }
                                else if (PREPressure == "PRE-CL600") { RatingPI = RatingPO = RatingPD = RatingWI = "PN100"; }
                                RatingFI = RatingFO = RatingCI = RatingCO = RatingII = RatingIO = RatingID = RatingQI = RatingQO = RatingGBI = "PN25";
                            }
                        }
                        break;
                }

                FacingPBD = FacingGBO = "N/A";
                if (MatFlange == true)
                {
                    FacingPI = FacingPO = FacingPD = FacingWI = FacingFI = FacingFO = FacingCI = FacingCO = FacingII = FacingIO = FacingID = FacingQI = FacingQO = FacingGBI = "RF";

                }
                else
                {
                    FacingPI = FacingPO = "WN RF";
                    FacingPD = FacingWI = FacingFI = FacingFO = FacingCI = FacingCO = FacingII = FacingIO = FacingID = FacingQI = FacingQO = FacingGBI = "SW RF";

                }
                if (Pumptype.Contains("ZA"))
                {
                    PBD = "1" + '"';
                }
                else { PBD = "2" + '"'; }
                GBO = "1/2" + '"';
                //transfer Connlist values
                foreach (string ite in DicPPIConnlist.Values)
                {
                    switch (ite)
                    {
                        case "PI":
                            DicPPISize.Add(ite, PI);
                            DicPPIRating.Add(ite, RatingPI);
                            DicPPIFacing.Add(ite, FacingPI);
                            DicPPIMaterial.Add(ite, MATPUMP);
                            break;
                        case "PO":
                            DicPPISize.Add(ite, PO);
                            DicPPIRating.Add(ite, RatingPO);
                            DicPPIFacing.Add(ite, FacingPO);
                            DicPPIMaterial.Add(ite, MATPUMP);
                            break;
                        case "PD":
                            DicPPISize.Add(ite, PD);
                            DicPPIRating.Add(ite, RatingPD);
                            DicPPIFacing.Add(ite, FacingPD);
                            DicPPIMaterial.Add(ite, MATPUMP);
                            break;
                        case "WI":
                            DicPPISize.Add(ite, WI);
                            DicPPIRating.Add(ite, RatingWI);
                            DicPPIFacing.Add(ite, FacingWI);
                            DicPPIMaterial.Add(ite, MATPUMP);
                            break;
                        case "PBD":
                            DicPPISize.Add(ite, PBD);
                            DicPPIRating.Add(ite, RatingPBD);
                            DicPPIFacing.Add(ite, FacingPBD);
                            DicPPIMaterial.Add(ite, "N/A");
                            break;
                        case "FI":
                            DicPPISize.Add(ite, FI);
                            DicPPIRating.Add(ite, RatingFI);
                            DicPPIFacing.Add(ite, FacingFI);
                            DicPPIMaterial.Add(ite, MATFLUSHING);
                            break;
                        case "FO":
                            DicPPISize.Add(ite, FO);
                            DicPPIRating.Add(ite, RatingFO);
                            DicPPIFacing.Add(ite, FacingFO);
                            DicPPIMaterial.Add(ite, MATFLUSHING);
                            break;
                        case "CI":
                            DicPPISize.Add(ite, CI);
                            DicPPIRating.Add(ite, RatingCI);
                            DicPPIFacing.Add(ite, FacingCI);
                            DicPPIMaterial.Add(ite, MATCOOLING);
                            break;
                        case "CO":
                            DicPPISize.Add(ite, CO);
                            DicPPIRating.Add(ite, RatingCO);
                            DicPPIFacing.Add(ite, FacingCO);
                            DicPPIMaterial.Add(ite, MATCOOLING);
                            break;
                        case "II":
                            DicPPISize.Add(ite, II);
                            DicPPIRating.Add(ite, RatingII);
                            DicPPIFacing.Add(ite, FacingII);
                            DicPPIMaterial.Add(ite, MATSTEAMJACKET);
                            break;
                        case "IO":
                            DicPPISize.Add(ite, IO);
                            DicPPIRating.Add(ite, RatingIO);
                            DicPPIFacing.Add(ite, FacingIO);
                            DicPPIMaterial.Add(ite, MATSTEAMJACKET);
                            break;
                        case "ID":
                            DicPPISize.Add(ite, ID);
                            DicPPIRating.Add(ite, RatingID);
                            DicPPIFacing.Add(ite, FacingID);
                            DicPPIMaterial.Add(ite, MATSTEAMJACKET);
                            break;
                        case "QI":
                            DicPPISize.Add(ite, QI);
                            DicPPIRating.Add(ite, RatingQI);
                            DicPPIFacing.Add(ite, FacingQI);
                            DicPPIMaterial.Add(ite, MATQUENCH);
                            break;
                        case "QO":
                            DicPPISize.Add(ite, QO);
                            DicPPIRating.Add(ite, RatingQO);
                            DicPPIFacing.Add(ite, FacingQO);
                            DicPPIMaterial.Add(ite, MATQUENCH);
                            break;
                        case "GBI":
                            DicPPISize.Add(ite, GBI);
                            DicPPIRating.Add(ite, RatingGBI);
                            DicPPIFacing.Add(ite, FacingGBI);
                            DicPPIMaterial.Add(ite, MATBUFFERGAS);
                            break;
                        case "GBO":
                            DicPPISize.Add(ite, GBO);
                            DicPPIRating.Add(ite, RatingGBO);
                            DicPPIFacing.Add(ite, FacingGBO);
                            DicPPIMaterial.Add(ite, "N/A");
                            break;
                    }
                }
                //transfer IDNOlist value
                string strSerialNO = SerialNO.Replace("/", "");
                string strIDPumpPre = IDPump.Replace("/", "").Substring(0, IDPump.Length - PumpQTY - 2);
                string strIDPumpSuf = IDPump.Replace(strIDPumpPre, "").Replace("/", "").ToString();
                for (int dicIDNO = 0; dicIDNO <= PumpQTY - 1; dicIDNO++)
                {
                    DicIDList.Add(strSerialNO.Substring(dicIDNO + dicIDNO * 5, 6), strIDPumpPre + strIDPumpSuf.Substring(dicIDNO, 1));
                }
                //CAD打开PID.dwg模板
                try
                {
                    db.ReadDwgFile("C:\\temp\\SulzerTools\\Release\\DWGTemplate\\TEMPLATE_PPI.dwg", System.IO.FileShare.ReadWrite, false, "");
                    db.CloseInput(true);
                    HostApplicationServices.WorkingDatabase = db;
                }
                catch (Autodesk.AutoCAD.Runtime.Exception e)
                {
                    EdPPI.WriteMessage("\nUnable to open .dwg file : " + e.StackTrace);
                    return;
                }
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;

                    DBDictionary myLayoutDictionary = (DBDictionary)db.LayoutDictionaryId.GetObject(OpenMode.ForWrite);
                    Layout myLayout;
                    foreach (DBDictionaryEntry entry in myLayoutDictionary)
                    {
                        if (entry.Key == "2") { myLayout = (Layout)tr.GetObject(entry.Value, OpenMode.ForWrite); break; }

                    }
                    LayoutManager ltman = LayoutManager.Current;
                    Application.SetSystemVariable("TILEMODE", 0);
                    EdPPI.SwitchToPaperSpace();
                    ltman.CurrentLayout = "2";
                    BlockTableRecord btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    LayerTable acLyrTbl = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                    //Insert blocks without AttributeDefinition
                    //Insert Pump block
                    ObjectId blkRecIdPump = ObjectId.Null;
                    {
                        blkRecIdPump = bt["Pump"];
                    }
                    using (BlockReference acBlkRefPump = new BlockReference(OCPoint, blkRecIdPump))
                    {
                        btr.AppendEntity(acBlkRefPump);
                        tr.AddNewlyCreatedDBObject(acBlkRefPump, true);
                    }
                    //Insert SteamJacket block
                    if (CaseBrket == true)
                    {
                        ObjectId blkRecIdSteamJacket = ObjectId.Null;
                        {
                            blkRecIdSteamJacket = bt["Jacket"];
                        }
                        using (BlockReference acBlkRefSteamJacket = new BlockReference(OCPoint, blkRecIdSteamJacket))
                        {
                            btr.AppendEntity(acBlkRefSteamJacket);
                            tr.AddNewlyCreatedDBObject(acBlkRefSteamJacket, true);
                        }
                    }
                    //Insert OilDrain block
                    if (BRHDrain == true)
                    {
                        ObjectId blkRecIdOilDrain = ObjectId.Null;
                        {
                            blkRecIdOilDrain = bt["OilDrain"];
                        }
                        using (BlockReference acBlkRefOilDrain = new BlockReference(OCPoint, blkRecIdOilDrain))
                        {
                            btr.AppendEntity(acBlkRefOilDrain);
                            tr.AddNewlyCreatedDBObject(acBlkRefOilDrain, true);
                        }
                    }
                    //Insert PDValve block
                    string BlkPDValve = "";
                    if (Pwarmup == true)
                    {
                        switch (PDValve)
                        {
                            case 0:
                                BlkPDValve = "PDWarmupW" + "O" + "Valve"; break;
                            case 1:
                                BlkPDValve = "PDWarmupW" + "S" + "Valve"; break;
                            case 2:
                                BlkPDValve = "PDWarmupW" + "D" + "Valve"; break;
                        }
                    }
                    else
                    {
                        switch (PDValve)
                        {
                            case 0:
                                BlkPDValve = "StdPDW" + "O" + "Valve"; break;
                            case 1:
                                BlkPDValve = "StdPDW" + "S" + "Valve"; break;
                            case 2:
                                BlkPDValve = "StdPDW" + "D" + "Valve"; break;
                        }
                    }
                    ObjectId blkRecIdPDValve = ObjectId.Null;
                    {
                        blkRecIdPDValve = bt[BlkPDValve];
                    }
                    using (BlockReference acBlkRefPDValve = new BlockReference(OCPoint, blkRecIdPDValve))
                    {
                        btr.AppendEntity(acBlkRefPDValve);
                        tr.AddNewlyCreatedDBObject(acBlkRefPDValve, true);
                    }
                    //Insert Baseplate block
                    ObjectId blkRecIdBaseDrain = ObjectId.Null;
                    {
                        blkRecIdBaseDrain = bt["BaseDrain"];
                    }
                    using (BlockReference acBlkRefBaseDrain = new BlockReference(OCPoint, blkRecIdBaseDrain))
                    {
                        btr.AppendEntity(acBlkRefBaseDrain);
                        tr.AddNewlyCreatedDBObject(acBlkRefBaseDrain, true);
                    }
                    if (BearingType.Contains("S"))
                    {
                        ObjectId blkRecIdBearingCooling = ObjectId.Null;
                        {
                            blkRecIdBearingCooling = bt["BearingCooling"];
                        }
                        using (BlockReference acBlkRefBearingCooling = new BlockReference(OCPoint, blkRecIdBearingCooling))
                        {
                            btr.AppendEntity(acBlkRefBearingCooling);
                            tr.AddNewlyCreatedDBObject(acBlkRefBearingCooling, true);
                        }
                    }


                    //Insert Flushing Plan Blocks
                    if (PlanA != "空白")
                    {
                        ObjectId blkRecIdPlanA = ObjectId.Null;
                        {
                            blkRecIdPlanA = bt[PlanA];
                        }
                        using (BlockReference acBlkRefPlanA = new BlockReference(OCPoint, blkRecIdPlanA))
                        {
                            btr.AppendEntity(acBlkRefPlanA);
                            tr.AddNewlyCreatedDBObject(acBlkRefPlanA, true);
                        }
                    }
                    //else
                    //{
                    //    PlanA = "";
                    //}
                    if (PlanC != "空白")
                    {
                        ObjectId blkRecIdPlanC = ObjectId.Null;
                        {
                            blkRecIdPlanC = bt[PlanC];
                        }
                        using (BlockReference acBlkRefPlanC = new BlockReference(OCPoint, blkRecIdPlanC))
                        {
                            btr.AppendEntity(acBlkRefPlanC);
                            tr.AddNewlyCreatedDBObject(acBlkRefPlanC, true);
                        }
                    }
                    if (PlanB != "空白")
                    {
                        ObjectId blkRecIdPlanB = ObjectId.Null;
                        {
                            blkRecIdPlanB = bt[PlanB];
                        }
                        using (BlockReference acBlkRefPlanB = new BlockReference(OCPoint, blkRecIdPlanB))
                        {
                            btr.AppendEntity(acBlkRefPlanB);
                            tr.AddNewlyCreatedDBObject(acBlkRefPlanB, true);
                        }
                    }
                    if (PlanM > 0)
                    {
                        ObjectId blkRecIdPlanM = ObjectId.Null;
                        {
                            blkRecIdPlanM = bt["PlanM"];
                        }
                        using (BlockReference acBlkRefPlanM = new BlockReference(OCPoint, blkRecIdPlanM))
                        {
                            btr.AppendEntity(acBlkRefPlanM);
                            tr.AddNewlyCreatedDBObject(acBlkRefPlanM, true);
                        }
                    }

                    //图层控制
                    LayerTableRecord acLyrTblRecCoolingSL = tr.GetObject(acLyrTbl["CoolingSL"], OpenMode.ForWrite) as LayerTableRecord;
                    acLyrTblRecCoolingSL.IsOff = true;
                    LayerTableRecord acLyrTblRecCoolingDL = tr.GetObject(acLyrTbl["CoolingDL"], OpenMode.ForWrite) as LayerTableRecord;
                    acLyrTblRecCoolingDL.IsOff = true;
                    LayerTableRecord acLyrTblRecMatFlange = tr.GetObject(acLyrTbl["MatFlange"], OpenMode.ForWrite) as LayerTableRecord;
                    acLyrTblRecMatFlange.IsOff = true;
                    if (PlanM >= 2)
                    {
                        acLyrTblRecCoolingDL.IsOff = false;
                    }
                    else
                    if (PlanM == 1)
                    {
                        acLyrTblRecCoolingSL.IsOff = false;
                    }
                    if (MatFlange == true) { acLyrTblRecMatFlange.IsOff = false; }


                    tr.Commit();
                }

                //Insert block of Connlist and set value
                double x = 510.85;
                double y = 634.53;
                double z = 0.0;
                foreach (string ite in DicPPIConnlist.Values)
                {
                    Point3d PConnList = new Point3d(x, y, z);
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
                        BlockTableRecord btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        ObjectId blkRecIdConnList = ObjectId.Null;
                        {
                            blkRecIdConnList = bt["ConnList"];
                        }
                        BlockReference acBlkRefConnList = new BlockReference(PConnList, blkRecIdConnList);
                        y = y - 13.5;
                        btr.AppendEntity(acBlkRefConnList);
                        tr.AddNewlyCreatedDBObject(acBlkRefConnList, true);
                        BlockTableRecord btrConnList = tr.GetObject(blkRecIdConnList, OpenMode.ForWrite) as BlockTableRecord;
                        tr.Commit();
                        btrConnList.DFUpdate(false, true, false);
                    }
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        // Get Block Table and Block Table Record 
                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord btrMS = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                        foreach (ObjectId id in btrMS)
                        {
                            // Fetch the entity 
                            Entity ent = (Entity)tr.GetObject(id, OpenMode.ForWrite, false);
                            // Test if the entity is a BlockReference 
                            if (ent.GetType() == typeof(BlockReference))
                            {
                                // It's a BlockReference... cast it: 
                                BlockReference blRef = (BlockReference)ent;
                                AttributeCollection attCol = blRef.AttributeCollection;

                                foreach (ObjectId objID in attCol)

                                {
                                    DBObject dbObj = tr.GetObject(objID, OpenMode.ForRead) as DBObject;
                                    AttributeReference attRef = dbObj as AttributeReference;
                                    switch (attRef.Tag)
                                    {
                                        case "ITEM":
                                            attRef.UpgradeOpen();
                                            if (attRef.TextString == "") { attRef.TextString = ite; }
                                            attRef.DowngradeOpen();
                                            break;
                                        case "SERVICEDESCIPTION":
                                            attRef.UpgradeOpen();
                                            if (attRef.TextString == "") { attRef.TextString = DicPPIDes[ite]; }
                                            attRef.DowngradeOpen();
                                            break;
                                        case "CONSIZE":
                                            attRef.UpgradeOpen();
                                            if (attRef.TextString == "") { attRef.TextString = DicPPISize[ite]; }
                                            attRef.DowngradeOpen();
                                            break;
                                        case "RATING":
                                            attRef.UpgradeOpen();
                                            if (attRef.TextString == "") { attRef.TextString = DicPPIRating[ite]; }
                                            attRef.DowngradeOpen();
                                            break;
                                        case "FACING":
                                            attRef.UpgradeOpen();
                                            if (attRef.TextString == "") { attRef.TextString = DicPPIFacing[ite]; }
                                            attRef.DowngradeOpen();
                                            break;
                                        case "MATERIAL":
                                            attRef.UpgradeOpen();
                                            if (attRef.TextString == "") { attRef.TextString = DicPPIMaterial[ite]; }
                                            attRef.DowngradeOpen();
                                            break;
                                    }

                                }
                            }
                        }

                        tr.Commit();
                    }
                }

                //其它属性定义赋值
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    // Get Block Table and Block Table Record 
                    BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btrMS = (BlockTableRecord)trans.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                    foreach (ObjectId id in btrMS)
                    {
                        // Fetch the entity 
                        Entity ent = (Entity)trans.GetObject(id, OpenMode.ForWrite, false);
                        // Test if the entity is a BlockReference 
                        if (ent.GetType() == typeof(BlockReference))
                        {
                            BlockReference blRef = (BlockReference)ent;
                            Autodesk.AutoCAD.DatabaseServices.AttributeCollection attCol = blRef.AttributeCollection;
                            foreach (ObjectId objID in attCol)
                            {
                                DBObject dbObj = trans.GetObject(objID, OpenMode.ForRead) as DBObject;
                                AttributeReference acAttRef = dbObj as AttributeReference;
                                switch (acAttRef.Tag)
                                {
                                    case "TYPE":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = Pumptype;
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "SIZE":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = PumpSize;
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "SEALPLAN":
                                        acAttRef.UpgradeOpen();
                                        string strtem = PlanA + "+" + PlanB.ToString().Replace("PLAN", "").Replace("-3rd", "").Replace("-4rd", "") + "+" + PlanC.ToString().Replace("PLAN", "");
                                        if (strtem.Contains("空白")) { strtem = strtem.Replace("+空白", ""); }
                                        acAttRef.TextString = strtem;
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "SulzerOrderNO":
                                        acAttRef.UpgradeOpen();
                                        string strtem2 = "00" + ItemNO;
                                        acAttRef.TextString = SulzerNO + "-" + strtem2.Substring(strtem2.Length - 4); 
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "ITEMNo":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = Savename.Replace("_Ver0", "");
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "Preparedby":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[13].ToString(); 
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "Reviewedby":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[14].ToString(); 
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "Approvedby":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[15].ToString(); 
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "PDATE":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = SDate;
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "RDATE":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = SDate;
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "ADATE":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = SDate;
                                        acAttRef.DowngradeOpen();
                                        break;
                                }
                            }
                        }
                    }
                    //Change to Layout1
                    Application.SetSystemVariable("TILEMODE", 0);
                    EdPPI.SwitchToPaperSpace();
                    LayoutManager ltman1 = LayoutManager.Current;
                    ltman1.CurrentLayout = "1";
                    //BlockTable bt1 = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btrMS1 = (BlockTableRecord)trans.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                    foreach (ObjectId id1 in btrMS1)
                    {
                        // Fetch the entity 
                        Entity ent1 = (Entity)trans.GetObject(id1, OpenMode.ForWrite, false);
                        // Test if the entity is a BlockReference 
                        if (ent1.GetType() == typeof(BlockReference))
                        {
                            // It's a BlockReference... cast it: 
                            BlockReference blRef1 = (BlockReference)ent1;
                            // Now you can acces attributes via the AttributeCollection 
                            // Do Something with: 
                            Autodesk.AutoCAD.DatabaseServices.AttributeCollection attCol1 = blRef1.AttributeCollection;
                            foreach (ObjectId objID1 in attCol1)
                            {
                                DBObject dbObj1 = trans.GetObject(objID1, OpenMode.ForRead) as DBObject;
                                AttributeReference acAttRef1 = dbObj1 as AttributeReference;
                                switch (acAttRef1.Tag)
                                {
                                    //Contract Information
                                    case "ORDERNO":
                                        acAttRef1.UpgradeOpen();
                                        acAttRef1.TextString = SulzerNO;
                                        acAttRef1.DowngradeOpen();
                                        break;
                                    case "PROJECT":
                                        acAttRef1.UpgradeOpen();
                                        acAttRef1.TextString = Project;
                                        acAttRef1.DowngradeOpen();
                                        break;
                                    case "PURCHASER":
                                        acAttRef1.UpgradeOpen();
                                        acAttRef1.TextString = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[7].ToString(); 
                                        acAttRef1.DowngradeOpen();
                                        break;
                                    case "USER":
                                        acAttRef1.UpgradeOpen();
                                        acAttRef1.TextString = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[8].ToString(); 
                                        acAttRef1.DowngradeOpen();
                                        break;
                                    case "CUST.PO#":
                                        acAttRef1.UpgradeOpen();
                                        acAttRef1.TextString = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[11].ToString();
                                        acAttRef1.DowngradeOpen();
                                        break;
                                    case "DEVICESERVICE":
                                        acAttRef1.UpgradeOpen();
                                        acAttRef1.TextString = ServiceName;
                                        acAttRef1.DowngradeOpen();
                                        break;
                                    case "INSTALLEDAT":
                                        acAttRef1.UpgradeOpen();
                                        acAttRef1.TextString = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[10].ToString();
                                        acAttRef1.DowngradeOpen();
                                        break;
                                    //Title Information
                                    case "ITEMNo":
                                        acAttRef1.UpgradeOpen();
                                        acAttRef1.TextString = Savename.Replace("_Ver0", "");
                                        acAttRef1.DowngradeOpen();
                                        break;
                                    case "SulzerOrderNO":
                                        string strtem1 = "00" + ItemNO;
                                        acAttRef1.UpgradeOpen();
                                        acAttRef1.TextString = SulzerNO + "-" + strtem1.Substring(strtem1.Length - 4);
                                        acAttRef1.DowngradeOpen();
                                        break;
                                    case "TYPE":
                                        acAttRef1.UpgradeOpen();
                                        acAttRef1.TextString = Pumptype;
                                        acAttRef1.DowngradeOpen();
                                        break;
                                    case "SIZE":
                                        acAttRef1.UpgradeOpen();
                                        acAttRef1.TextString = PumpSize;
                                        acAttRef1.DowngradeOpen();
                                        break;
                                    case "SEALPLAN":
                                        acAttRef1.UpgradeOpen();
                                        string strtem = PlanA + "+" + PlanB.ToString().Replace("PLAN", "").Replace("-3rd", "").Replace("-4rd", "") + "+" + PlanC.ToString().Replace("PLAN", "");
                                        if (strtem.Contains("空白")) { strtem = strtem.Replace("+空白", ""); }
                                        acAttRef1.TextString = strtem; acAttRef1.DowngradeOpen();
                                        break;
                                    case "Preparedby":
                                        acAttRef1.UpgradeOpen();
                                        acAttRef1.TextString = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[13].ToString(); 
                                        acAttRef1.DowngradeOpen();
                                        break;
                                    case "Reviewedby":
                                        acAttRef1.UpgradeOpen();
                                        acAttRef1.TextString = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[14].ToString();
                                        acAttRef1.DowngradeOpen();
                                        break;
                                    case "Approvedby":
                                        acAttRef1.UpgradeOpen();
                                        acAttRef1.TextString = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[15].ToString(); 
                                        acAttRef1.DowngradeOpen();
                                        break;
                                    case "DATEP":
                                        acAttRef1.UpgradeOpen();
                                        acAttRef1.TextString = SDate;
                                        acAttRef1.DowngradeOpen();
                                        break;
                                    case "DATER":
                                        acAttRef1.UpgradeOpen();
                                        acAttRef1.TextString = SDate;
                                        acAttRef1.DowngradeOpen();
                                        break;
                                    case "DATEA":
                                        acAttRef1.UpgradeOpen();
                                        acAttRef1.TextString = SDate;
                                        acAttRef1.DowngradeOpen();
                                        break;
                                }
                            }
                        }
                    }
                    trans.Commit();
                }
                //Insert block of IDNO and set value
                double x1 = 994.05;
                double y1 = 681.84;
                double z1 = 0.0;
                foreach (var itetemp in DicIDList)
                {
                    Point3d PIDNO = new Point3d(x1, y1, z1);
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTable bttem = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
                        BlockTableRecord btrMStem = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        ObjectId blkRecIdIDNO = ObjectId.Null;
                        {
                            blkRecIdIDNO = bttem["IDNO"];
                        }
                        BlockReference acBlkRefIDNO = new BlockReference(PIDNO, blkRecIdIDNO);
                        y1 = y1 - 12;
                        btrMStem.AppendEntity(acBlkRefIDNO);
                        tr.AddNewlyCreatedDBObject(acBlkRefIDNO, true);
                        BlockTableRecord btrIDNO = tr.GetObject(blkRecIdIDNO, OpenMode.ForWrite) as BlockTableRecord;
                        tr.Commit();
                        btrIDNO.DFUpdate(false, true, false);
                    }
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        // Get Block Table and Block Table Record 
                        BlockTable bttem = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord btrMStem = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                        foreach (ObjectId objid in btrMStem)
                        {

                            Entity enttem = (Entity)tr.GetObject(objid, OpenMode.ForWrite, false);

                            if (enttem.GetType() == typeof(BlockReference))
                            {

                                BlockReference blRef = (BlockReference)enttem;
                                AttributeCollection attCol = blRef.AttributeCollection;

                                foreach (ObjectId objID in attCol)

                                {
                                    DBObject dbObj = tr.GetObject(objID, OpenMode.ForRead) as DBObject;
                                    AttributeReference attRef = dbObj as AttributeReference;
                                    switch (attRef.Tag)
                                    {
                                        case "SERIALNO":
                                            attRef.UpgradeOpen();
                                            if (attRef.TextString == "") { attRef.TextString = itetemp.Key.ToString(); }
                                            attRef.DowngradeOpen();
                                            break;
                                        case "PUMP":
                                            attRef.UpgradeOpen();
                                            if (attRef.TextString == "") { attRef.TextString = itetemp.Value.ToString(); }
                                            attRef.DowngradeOpen();
                                            break;
                                            //case "DRIVER":
                                            //    attRef.UpgradeOpen();
                                            //    if (attRef.TextString == "") { attRef.TextString = Pumptype + "-" + (10 + 10 * ii).ToString(); }
                                            //    attRef.DowngradeOpen();
                                            //    break;
                                    }

                                }
                            }
                        }

                        tr.Commit();
                    }
                }
                HostApplicationServices.WorkingDatabase = workingDB;//切换当前激活的数据库文件
                db.SaveAs(form.SelectedPath + "\\" + Savename + ".dwg", DwgVersion.AC1024);//另存dwg文件
                workingDB.Dispose();
                db.Dispose();
                //begin to plot2PDF 开始打印PDF
                Database workingDB4plot = HostApplicationServices.WorkingDatabase;
                Database db4plot = new Database(false, true);//定义新数据库文件
                try
                {
                    db4plot.ReadDwgFile(form.SelectedPath + "\\" + Savename + ".dwg", System.IO.FileShare.ReadWrite, false, "");
                    db4plot.CloseInput(true);
                    HostApplicationServices.WorkingDatabase = db4plot;
                }
                catch (Autodesk.AutoCAD.Runtime.Exception e)
                {
                    EdPPI.WriteMessage("\nUnable to open .dwg file : " + e.StackTrace);
                    return;
                }
                short bgp = (short)Application.GetSystemVariable("BACKGROUNDPLOT");
                try
                {
                    Application.SetSystemVariable("BACKGROUNDPLOT", 0);

                    using (Transaction tr4plot = db4plot.TransactionManager.StartTransaction())
                    {
                        List<Layout> layouts = new List<Layout>();
                        DBDictionary layoutDict = (DBDictionary)db4plot.LayoutDictionaryId.GetObject(OpenMode.ForRead);
                        foreach (DBDictionaryEntry entry in layoutDict)
                        {
                            if (entry.Key != "Model")
                            {
                                layouts.Add((Layout)tr4plot.GetObject(entry.Value, OpenMode.ForRead));
                            }
                        }
                        layouts.Sort((l1, l2) => l1.TabOrder.CompareTo(l2.TabOrder));

                        string filename = Path.ChangeExtension(form.SelectedPath + "\\" + Savename, "pdf");

                        MultiSheetsPdf plotter = new MultiSheetsPdf(filename, layouts);
                        plotter.Publish();


                        tr4plot.Commit();

                    }
                }
                catch (System.Exception e)
                {
                    //Editor ed = AcAp.DocumentManager.MdiActiveDocument.Editor;
                    EdPPI.WriteMessage("\nError: {0}\n{1}", e.Message, e.StackTrace);
                }
                finally
                {
                    Application.SetSystemVariable("BACKGROUNDPLOT", bgp);
                }

                HostApplicationServices.WorkingDatabase = workingDB4plot;//切换当前激活的数据库文件
                workingDB4plot.Dispose();
                db4plot.Dispose();


            }
            MessageBox.Show("Congratulations!" + "\r\n" + "PPI documents completed!");
        }
        [CommandMethod("PCS")]
        public void PCS()
        {
            Document DocPCS = Application.DocumentManager.MdiActiveDocument;
            Editor EdPCS = Application.DocumentManager.MdiActiveDocument.Editor;
            //define savepath
            SelectPathForm form = new SelectPathForm("PCS", GlobalVariableController.GlobalDataSet);
            form.ShowDialog();
            LoopC = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows.Count;
            //Basic information
            CustomerName = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[8].ToString();
            Project = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[9].ToString();
            SulzerNO = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[12].ToString();
            FlangeStd = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[17].ToString();
            SDate = string.Format("{0:yyyy/MM/dd}", DateTime.Parse(form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[16].ToString()));
            int PumpQTY = 0;
            String BlkZPureMistDView = "ZDViewPureOilMist";
            String BlkPREPureMistDView = "PREDVPUREMIST";
            String BlkPRESpacerSleeveDView = "DViewPREDRYSHAFT";
            String BlkZAOFrWring = "ZAORearWearing";
            string BlkZEGDView = "ZXDView";
            String BlkZ7630DView = "ZYDView";
            //PRE BearingSize Flag Dictionary
            string[] PREBringSizeSmall = { "0A", "0B", "1A", "2A" };
            Dictionary<string, string> DicPREBringSmall = new Dictionary<string, string>();
            for(int iPREBringSizeSmall = 0; iPREBringSizeSmall <= PREBringSizeSmall.Length - 1; iPREBringSizeSmall++) { DicPREBringSmall.Add(PREBringSizeSmall[iPREBringSizeSmall], ""); }
            for (int i = 0; i <= LoopC - 1; i++)
            {
                Database workingDB = HostApplicationServices.WorkingDatabase;//定义数据库文件，并指向当前活动文件
                Database db = new Database(false, true);//定义新数据库文件
                //Define conditions related with Pump Assembly Block
                string PumpType = "";
                String PumpTypeZ = "";
                string PumpSize = "";
                string BearingSize = "";
                string BearingSizeStr = "";
                string BearingSize4ZPartNO = "";
                String BearingType = "";
                Boolean PureMistFlag = false;
                string PureMistStr = "";
                String API610Ver = "";
                Boolean InducerFlag = false;
                String InducerStr = "";
                Boolean BalHoleFlag = false;
                String BalHoleStr = "";
                String BalZRearRingStr = "";
                String ShaftType = "";
                Boolean PLAN02 = false;
                Boolean StmJacket = false;
                string BearingCage = "";
                //define Material MOD value from SumListTable
                String ShaftMOD = "";
                String VoluteMOD = "";
                string ImpellerMOD = "";
                string WearingMOD = "";
                String CoverCLPlateMOD = "";
                String THBushMOD = "";
                String KeyMOD = "";
                String GasketMOD = "";
                String SpacerSleeveMOD = "";
                String ImpellerNutMOD = "";
                String InducerMOD = "";
                String IndSleeveMOD = "";
                String WearPlateMOD = "";
                //Define Shared Block Name
                string BlkPump = "";
                //only 4 Z Serise Pump
                String BlkZSpacerSleeveDView = "";
                //only 4 PRE Serise Pump
                string BlkPREImpeller = "";
                string BlkPRECover = "";
                string BlkPREBringHouse = "";
                string BlkPREWetShaft = "";
                String BlkPRECoverCool = "";
                //define ID List
                int.TryParse(form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<String>("ColQty"), out PumpQTY);
                string SerialNO = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColSeriesNo");
                string IDPump = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColTagNo");
                string ServiceName = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColService");
                //Set MOD value from SumlistTable
                ShaftType = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i]["ColDryWetShaft"].ToString();
                string MaterialClass = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColMaterialClass");
                if (MaterialClass.Contains("MOD")) { MaterialClass = "MOD"; }
                if (ShaftType == "Wet") { ShaftMOD = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColWetInput"); }
                if (MaterialClass == "MOD")
                {
                    VoluteMOD = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColVoluteCasing");
                    ImpellerMOD = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColImpeller");
                    WearingMOD = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColStationaryWearRing");
                    CoverCLPlateMOD = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColCoverCoolingPlate");
                    THBushMOD = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColThroatBush");
                    KeyMOD = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColImpellerKey").ToString();
                    GasketMOD = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColGasket");
                    SpacerSleeveMOD = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColSpacerSleeve");
                    ImpellerNutMOD = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColImpellerNut");
                    InducerMOD = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColInducer");
                    IndSleeveMOD = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColInducerSleeve");
                    WearPlateMOD = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColWearPlate");
                }
                //set Other Values
                string ItemNO = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColItemNo");
                Savename = "PCS" + SulzerNO + ("00" + ItemNO).ToString().Substring(("00" + ItemNO).ToString().Length - 4, 4) + "-01" + "_Ver0";
                //Define IDList AttributeDefinition Dictionary 
                Dictionary<string, string> DicIDList = new Dictionary<string, string>();
                List<string> LstBlkName = new List<string>();
                //Get Value&Information 
                PumpType = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<String>("ColType");
                if (PumpType == "ZEG") { PumpTypeZ = "ZE"; } else { PumpTypeZ = PumpType; }
                if (PumpType.Substring(0, 1) == "Z")
                {
                    PumpSize = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColSize").Split('(')[0].Replace(PumpType, "").ToString();
                    API610Ver = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColSize").Split('(')[1].Split(')')[0].Replace("th", "");
                    BearingSizeStr = PumpSize.Split('-')[1].Substring(0, 1);
                    if (int.Parse(PumpSize.Split('-')[1].Substring(0, 1)) >= 6) { BearingSize = "LKLarge"; BearingSize4ZPartNO = "(LK6)"; } else { BearingSize = "LKSmall"; }
                }
                else
                {
                    PumpSize = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColSize");
                    string expSize = "SIZE = '" + PumpSize + "' ";
                    DataRow[] newRBrSize= new DataRow[0]; ;
                    newRBrSize = form.BASICDATASAT.Tables["PRE-BEARINGSIZETABLE"].Select(expSize);
                    BearingSize = newRBrSize[0].Field<string>(PumpType);
                    BearingSizeStr = BearingSize.Substring(0, 1);
                    if (DicPREBringSmall.ContainsKey(BearingSize)) { BearingSize = "SMALL";  } else { BearingSize = "LARGE"; }
                    InducerFlag = Boolean.Parse(form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i]["ColPREInducerYesNo"].ToString());
                    if (InducerFlag == false) { InducerStr = "WOINDUCER"; } else { InducerStr = "WIINDUCER"; }
                }
                BearingType = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColBearing");
                PureMistFlag = Boolean.Parse(form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i]["ColReservedPureLubrication"].ToString());
                if (PureMistFlag == true) { PureMistStr = "WIPUMIST"; } else { PureMistStr = "WOPUMIST"; }
                BalHoleFlag = Boolean.Parse(form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i]["ColImpellerBalanceHole"].ToString());
                if (BalHoleFlag == true) { BalHoleStr = "WI"; BalZRearRingStr = "WO"; } else { BalHoleStr = "WO"; BalZRearRingStr = "WI"; }
                if (PumpType == "PREK" || PumpType == "PRESF") { BlkPREImpeller = "IMPL-" + PumpType; } else { BlkPREImpeller = "IMPL-" + PumpType + "-" + BalHoleStr + "BALHOLE"; }
                
                BlkZSpacerSleeveDView = "SpaSleeve" + ShaftType + "Shaft";
                if (form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i].Field<string>("ColSealPlanA") == "PLAN02") { PLAN02 = true; }
                StmJacket = Boolean.Parse(form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i]["ColCaseBracket"].ToString());
                BearingCage = form.SUMLISTDATASAT.Tables["DataTableBody"].Rows[i]["ColBearingCageMaterial"].ToString();
                if (PLAN02 == true || StmJacket == true) { BlkPRECoverCool = "WICOOL"; } else { BlkPRECoverCool = "WOCOOL"; }
                //define Varies 4 PartList
                string expZPNList1 = "";
                string expZPNList2 = "";
                string expZPNList3 = "";
                string expZMaterial1 = "";
                string expZMaterial2 = "";
                string expZMaterial3 = "";
                string expPRECode1 = "";
                string expPRECode2 = "";
                string expPRECode3 = "";
                String PRECODE = "";
                DataRow[] newRPRECODE=new DataRow[0];
                string expPREBEARINGCAGE = "";
                DataRow[] newRPREBEARINGCAGE= new DataRow[0]; ;
                string expZBEARINGCAGE = "";
                DataRow[] newRZBEARINGCAGE= new DataRow[0]; ;
                string expPCSPRE1 = "";
                string expPCSPRE2 = "";
                string expPCSPRE3 = "";
                string expPCSPRE4 = "";
                string expPCSPRE5 = "";
                //Get PartList Value from BASICDATA
                //only 4 Z Serise Pump
                if(PumpType.Substring(0,1)=="Z")
                {
                    //Lookup PCSZPARTNO Table
                    if (PumpType.Substring(0, 2) == "ZE")
                    {
                        expZPNList1 = "ITEM = '" + PumpTypeZ + API610Ver+"th" + BearingSize4ZPartNO + "' AND ";
                    }
                    else
                    {
                        expZPNList1 = "ITEM = '" + PumpTypeZ + BearingSize4ZPartNO + "' AND ";
                    }
                    expZPNList2 = "BALHOLE = '" + BalHoleStr + "' AND ";
                    expZPNList3 = "BRTYPE = '" + BearingType + "' ";
                    newRPartListZ = form.BASICDATASAT.Tables["PCSZPARTNOTABLE"].Select(expZPNList1 + expZPNList2 + expZPNList3);
                    //Lookup PCSZMATERIAL Table
                    expZMaterial1 = "ITEM = '" + "NAME" + "' ";
                    expZMaterial2 = "ITEM = '" + "QTY" + "' ";
                    expZMaterial3 = "ITEM = '" + MaterialClass + "' ";
                    newRZMaterial1 = form.BASICDATASAT.Tables["PCSZMATERIALTABLE"].Select(expZMaterial1);
                    newRZMaterial2 = form.BASICDATASAT.Tables["PCSZMATERIALTABLE"].Select(expZMaterial2);
                    newRZMaterial3 = form.BASICDATASAT.Tables["PCSZMATERIALTABLE"].Select(expZMaterial3);
                    //Lookup ZBEARINGCAGE Table
                    expZBEARINGCAGE = "LK = '" + BearingSizeStr + "' ";
                    newRZBEARINGCAGE = form.BASICDATASAT.Tables["ZBEARINGCAGETABLE"].Select(expZBEARINGCAGE);
                    //Modify Shaft Material 
                    if (ShaftType == "Wet") { newRZMaterial3[0].SetField<String>("SHAFT", ShaftMOD); newRZMaterial3[0].SetField<String>("KEY1", "S31803"); }
                    else { newRZMaterial3[0].SetField<String>("SHAFT", "30Cr13"); newRZMaterial3[0].SetField<String>("KEY1", "2Cr13"); }
                    //set Bearing Material
                    newRZMaterial3[0].SetField<String>("COLROLLERBRING", newRZBEARINGCAGE[0].Field<string>("ROLLER") + BearingCage);
                    if (BearingType.Substring(0, 1) == "Q")
                    { newRZMaterial3[0].SetField<String>("ANGCONTBALBRING", newRZBEARINGCAGE[0].Field<string>("ANGULAR") + BearingCage); }
                    else
                    { newRZMaterial3[0].SetField<String>("TAPERROLBRING", newRZBEARINGCAGE[0].Field<string>("TAPER")); }
                    //Modify MOD Material
                    if (MaterialClass == "MOD")
                    {
                        newRZMaterial3[0].SetField<String>("VOLUTE", VoluteMOD);
                        newRZMaterial3[0].SetField<String>("CASINGCOVER", VoluteMOD);
                        newRZMaterial3[0].SetField<String>("IMPELLER", ImpellerMOD);
                        newRZMaterial3[0].SetField<String>("FSTWRING", WearingMOD);
                        newRZMaterial3[0].SetField<String>("RSTWRING", WearingMOD);
                        newRZMaterial3[0].SetField<String>("FIMWRING", WearingMOD);
                        newRZMaterial3[0].SetField<String>("RIMWRING", WearingMOD);
                        newRZMaterial3[0].SetField<String>("THRBUSH", THBushMOD);
                        newRZMaterial3[0].SetField<String>("KEY1", KeyMOD);
                        newRZMaterial3[0].SetField<String>("GASKET1", GasketMOD);
                        newRZMaterial3[0].SetField<String>("SPASLEEVE", SpacerSleeveMOD);
                        newRZMaterial3[0].SetField<String>("IMPELLERNUT1", ImpellerNutMOD);
                        newRZMaterial3[0].SetField<String>("IMPELLERNUT2", ImpellerNutMOD);
                        newRZMaterial3[0].SetField<String>("FWEARPLATE", WearPlateMOD);
                        newRZMaterial3[0].SetField<String>("RWEARPLATE", WearPlateMOD);
                    }
                }
                else
                {
                    //only 4 PRE Pump
                    //Lookup PRECODE Table
                    expPRECode1 = "BEARINGTYPE = '" + BearingType + "' AND ";
                    expPRECode2 = "INDUCER = '" + InducerStr + "' AND ";
                    expPRECode3 = "PUMPTYPEBALHOLE = '" + PumpType +  BalHoleStr + "' ";
                    newRPRECODE = form.BASICDATASAT.Tables["PRECODETABLE"].Select(expPRECode1 + expPRECode2 + expPRECode3);
                    PRECODE = newRPRECODE[0].Field<string>(BearingSize + BlkPRECoverCool);
                    //Lookup PREBEARINGCAGE Table
                    expPREBEARINGCAGE = "MEC = '" + BearingSizeStr + "' ";
                    newRPREBEARINGCAGE = form.BASICDATASAT.Tables["PREBEARINGCAGETABLE"].Select(expPREBEARINGCAGE);
                    //Lookup PCSPRE Table
                    expPCSPRE1 = "ITEM = '" + "NAME" + "' ";
                    expPCSPRE2 = "ITEM = '" + "QTY" + "' ";
                    expPCSPRE3 = "ITEM = '" + MaterialClass + "' ";
                    expPCSPRE4 = "ITEM = '" + "PartNO" + "' ";
                    expPCSPRE5 = "ITEM = '" + PRECODE + "' ";
                    if (PRECODE != "")
                    {
                        newRPCSPRE1 = form.BASICDATASAT.Tables["PCSPRETABLE"].Select(expPCSPRE1);
                        newRPCSPRE2 = form.BASICDATASAT.Tables["PCSPRETABLE"].Select(expPCSPRE2);
                        newRPCSPRE3 = form.BASICDATASAT.Tables["PCSPRETABLE"].Select(expPCSPRE3);
                        newRPCSPRE4 = form.BASICDATASAT.Tables["PCSPRETABLE"].Select(expPCSPRE4);
                        newRPCSPRE5 = form.BASICDATASAT.Tables["PCSPRETABLE"].Select(expPCSPRE5);
                        //Modify Shaft Material 
                        if (ShaftType == "Wet") { newRPCSPRE3[0].SetField<String>("Shaft", ShaftMOD); } else { newRPCSPRE3[0].SetField<String>("SHAFT", "30Cr13"); }
                        //set Bearing Material
                        newRPCSPRE3[0].SetField<String>("COLROLLERBRING", newRPREBEARINGCAGE[0].Field<string>("ROLLER") + BearingCage);
                        newRPCSPRE3[0].SetField<String>("ANGCONTBALBRING", newRPREBEARINGCAGE[0].Field<string>("ANGULAR") + BearingCage);
                        //Modify MOD Material
                        if (MaterialClass == "MOD")
                        {
                            newRPCSPRE3[0].SetField<String>("VoluteCasing", VoluteMOD);
                            newRPCSPRE3[0].SetField<String>("CasingCover", VoluteMOD);
                            newRPCSPRE3[0].SetField<String>("Impeller", ImpellerMOD);
                            newRPCSPRE3[0].SetField<String>("StationaryWearRingF", WearingMOD);
                            newRPCSPRE3[0].SetField<String>("StationaryWearRingR", WearingMOD);
                            newRPCSPRE3[0].SetField<String>("ImpellerWearRingF", WearingMOD);
                            newRPCSPRE3[0].SetField<String>("ImpellerWearRingR", WearingMOD);
                            newRPCSPRE3[0].SetField<String>("Cover-CoolPlate", CoverCLPlateMOD);
                            newRPCSPRE3[0].SetField<String>("ThroatBush", THBushMOD);
                            newRPCSPRE3[0].SetField<String>("Key1", KeyMOD);
                            newRPCSPRE3[0].SetField<String>("Key2", KeyMOD);
                            newRPCSPRE3[0].SetField<String>("Gasket1", GasketMOD);
                            newRPCSPRE3[0].SetField<String>("SpacerSleeve", SpacerSleeveMOD);
                            newRPCSPRE3[0].SetField<String>("ImpellerNut", ImpellerNutMOD);
                            newRPCSPRE3[0].SetField<String>("Inducer", InducerMOD);
                            newRPCSPRE3[0].SetField<String>("InducerSleeve", IndSleeveMOD);
                            newRPCSPRE3[0].SetField<String>("FrontWearPlate", WearPlateMOD);
                        }
                    }
                }
                //transfer values 2 PCSPartList class properties
                List<PCSPartList> PartList = new List<PCSPartList>();
                if(PumpType.Substring(0,1)=="Z")
                {
                    for (int iPartList = 3; iPartList <= newRPartListZ[0].ItemArray.Length  - 1; iPartList++)
                    {
                        if (newRPartListZ[0][iPartList].ToString() != "")
                        {
                            PartList.Add(new PCSPartList()
                            {
                                PartNO = newRPartListZ[0][iPartList].ToString(),
                                DES = newRZMaterial1[0][iPartList].ToString(),
                                Material = newRZMaterial3[0][iPartList].ToString(),
                                QTY = newRZMaterial2[0][iPartList].ToString()
                            });
                        }
                    }
                }
                else
                {
                    for (int iPartList = 2; iPartList <= newRPCSPRE5[0].ItemArray.Length - 1; iPartList++)
                    {
                        if (newRPCSPRE5[0][iPartList].ToString() != "")
                        {
                            PartList.Add(new PCSPartList()
                            {
                                PartNO = newRPCSPRE4[0][iPartList].ToString(),
                                DES = newRPCSPRE1[0][iPartList].ToString(),
                                Material = newRPCSPRE3[0][iPartList].ToString(),
                                QTY = newRPCSPRE2[0][iPartList].ToString()
                            });
                        }
                    }
                }
                //transfer IDNOlist value
                string strSerialNO = SerialNO.Replace("/", "");
                string strIDPumpPre = IDPump.Replace("/", "").Substring(0, IDPump.Length - PumpQTY - 2);
                string strIDPumpSuf = IDPump.Replace(strIDPumpPre, "").Replace("/", "").ToString();
                for (int dicIDNO = 0; dicIDNO <= PumpQTY - 1; dicIDNO++)
                {
                    DicIDList.Add(strSerialNO.Substring(dicIDNO + dicIDNO * 5, 6), strIDPumpPre + strIDPumpSuf.Substring(dicIDNO, 1));
                }
                
                //Collect Basic Blocks Name
                //Only 4 Z Serise Pump
                if (PumpType.Substring(0, 1) == "Z")
                {
                    
                    BlkPump = PumpTypeZ + "-" + API610Ver + "TH-" + BearingType + "-" + BearingSize + "-" + BalZRearRingStr;
                    LstBlkName.Add(BlkPump);
                    if (PureMistFlag == true) { LstBlkName.Add(BlkZPureMistDView);}
                    LstBlkName.Add(BlkZSpacerSleeveDView);
                    if (PumpType.Contains("ZAO")) { LstBlkName.Add(BlkZAOFrWring);}
                    if( PumpType.Contains("ZE") &&(StmJacket == true ||PLAN02==true)) { LstBlkName.Add(BlkZEGDView);}
                    if (PumpSize == "400-7630") { LstBlkName.Add(BlkZ7630DView); }
                }
                //Only 4 PRE Serise Pump
                else
                {
                    BlkPump = "VOLT-" +  PumpType + "-" + InducerStr;
                    LstBlkName.Add(BlkPump);
                    LstBlkName.Add(BlkPREImpeller); 
                    if("BR"+ BearingSize=="BRLARGE" && PumpType=="PRESF")
                    {
                        BlkPRECover = "COVE-PRESF-BRLARGE-WOBAHOL-" + BlkPRECoverCool;
                    }
                    else if(BalHoleFlag==true)
                    {
                        BlkPRECover = "COVE-PRE-BR" + BearingSize + "-" + BalHoleStr + "BAHOL-" + BlkPRECoverCool;
                    }
                    else
                    {
                        BlkPRECover = "COVE-PREX-BR" + BearingSize + "-" + BalHoleStr + "BAHOL-" + BlkPRECoverCool;
                    }
                    LstBlkName.Add(BlkPRECover);
                    BlkPREBringHouse = "BEAH-" + BearingType + "-" + PureMistStr;
                    LstBlkName.Add(BlkPREBringHouse);
                    BlkPREWetShaft = "SHAFT-PREX-" + BearingSize;
                    LstBlkName.Add(BlkPREWetShaft);
                    if (PureMistFlag == true) { LstBlkName.Add(BlkPREPureMistDView); }
                    if (ShaftType.Substring(0, 1) == "D") { LstBlkName.Add(BlkPRESpacerSleeveDView); }

                }
                //Open PCS dwg Database
                try
                {
                    db.ReadDwgFile("C:\\temp\\SulzerTools\\Release\\DWGTemplate\\TEMPLATE_PCS.dwg", System.IO.FileShare.ReadWrite, false, "");
                    db.CloseInput(true);
                    HostApplicationServices.WorkingDatabase = db;
                }
                catch (Autodesk.AutoCAD.Runtime.Exception e)
                {
                    EdPCS.WriteMessage("\nUnable to open .dwg file : " + e.StackTrace);
                    return;
                }
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;

                    DBDictionary myLayoutDictionary = (DBDictionary)db.LayoutDictionaryId.GetObject(OpenMode.ForWrite);
                    Layout myLayout;
                    foreach (DBDictionaryEntry entry in myLayoutDictionary)
                    {
                        if (entry.Key == "2") { myLayout = (Layout)tr.GetObject(entry.Value, OpenMode.ForWrite); break; }

                    }
                    LayoutManager ltman = LayoutManager.Current;
                    Application.SetSystemVariable("TILEMODE", 0);
                    EdPCS.SwitchToPaperSpace();
                    ltman.CurrentLayout = "2";
                    BlockTableRecord btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    LayerTable acLyrTbl = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                    //Insert blocks without AttributeDefinition
                    for (int BlkLoop = 0;BlkLoop<= LstBlkName.Count-1;BlkLoop++)
                    {
                        ObjectId blkRecIdBasic = ObjectId.Null;
                        {
                            blkRecIdBasic = bt[LstBlkName[BlkLoop]];
                        }
                        using (BlockReference acBlkRefBasic = new BlockReference(OCPoint, blkRecIdBasic))
                        {
                            btr.AppendEntity(acBlkRefBasic);
                            tr.AddNewlyCreatedDBObject(acBlkRefBasic, true);
                        }
                    }
                    //control Layers
                    LayerTableRecord acLyrTblRecZEG = tr.GetObject(acLyrTbl["ZEG"], OpenMode.ForWrite) as LayerTableRecord;
                    acLyrTblRecZEG.IsOff = true;
                    LayerTableRecord acLyrTblRecZ400_7630 = tr.GetObject(acLyrTbl["Z400_7630"], OpenMode.ForWrite) as LayerTableRecord;
                    acLyrTblRecZ400_7630.IsOff = true;
                    LayerTableRecord acLyrTblRecPREDRYSHAFT = tr.GetObject(acLyrTbl["PREDRYSHAFT"], OpenMode.ForWrite) as LayerTableRecord;
                    acLyrTblRecPREDRYSHAFT.IsOff = true;
                    LayerTableRecord acLyrTblRecPREJACKET = tr.GetObject(acLyrTbl["PREJACKET"], OpenMode.ForWrite) as LayerTableRecord;
                    acLyrTblRecPREJACKET.IsOff = true;
                    if (PumpType.Contains("ZEG")) { acLyrTblRecZEG.IsOff = false; }
                    if (PumpSize == "400-7630") { acLyrTblRecZ400_7630.IsOff = false; }
                    if (ShaftType.Substring(0, 1) == "D") { acLyrTblRecPREDRYSHAFT.IsOff = false; }
                    if ( StmJacket == true) { acLyrTblRecPREJACKET.IsOff = false; }
                    tr.Commit();
                }
                //Insert Partlist and set value  
                double x = -450.00;
                double y = 9699.74;
                double z = 0.0;
                foreach (PCSPartList Plst in PartList)
                {
                    Point3d PConnList = new Point3d(x, y, z);
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
                        BlockTableRecord btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        ObjectId blkRecIdConnList = ObjectId.Null;
                        {
                            blkRecIdConnList = bt["PcsTitleList1"];
                        }
                        BlockReference acBlkRefConnList = new BlockReference(PConnList, blkRecIdConnList);
                        y = y - 190;
                        btr.AppendEntity(acBlkRefConnList);
                        tr.AddNewlyCreatedDBObject(acBlkRefConnList, true);
                        BlockTableRecord btrConnList = tr.GetObject(blkRecIdConnList, OpenMode.ForWrite) as BlockTableRecord;
                        tr.Commit();
                        btrConnList.DFUpdate(false, true, false);
                    }
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        // Get Block Table and Block Table Record 
                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord btrMS = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                        foreach (ObjectId id in btrMS)
                        {
                            Entity ent = (Entity)tr.GetObject(id, OpenMode.ForWrite, false);
                            if (ent.GetType() == typeof(BlockReference))
                            {
                                BlockReference blRef = (BlockReference)ent;
                                AttributeCollection attCol = blRef.AttributeCollection;

                                foreach (ObjectId objID in attCol)

                                {
                                    DBObject dbObj = tr.GetObject(objID, OpenMode.ForRead) as DBObject;
                                    AttributeReference attRef = dbObj as AttributeReference;
                                    switch (attRef.Tag)
                                    {
                                        case "PN1":
                                            attRef.UpgradeOpen();
                                            if (attRef.TextString == "") { attRef.TextString = Plst.PartNO; }
                                            attRef.DowngradeOpen();
                                            break;
                                        case "DES1":
                                            attRef.UpgradeOpen();
                                            if (attRef.TextString == "") { attRef.TextString = Plst.DES ; }
                                            attRef.DowngradeOpen();
                                            break;
                                        case "MAT1":
                                            attRef.UpgradeOpen();
                                            if (attRef.TextString == "") { attRef.TextString = Plst.Material ; }
                                            attRef.DowngradeOpen();
                                            break;
                                        case "QTY1":
                                            attRef.UpgradeOpen();
                                            if (attRef.TextString == "") { attRef.TextString = Plst.QTY; }
                                            attRef.DowngradeOpen();
                                            break;
                                    }

                                }
                            }
                        }

                        tr.Commit();
                    }
                }
                //Insert block of IDNO and set value
                double x1 = 10606.67;
                double y1 = 8200.44;
                double z1 = 0.0;
                foreach (var itetemp in DicIDList)
                {
                    Point3d PIDNO = new Point3d(x1, y1, z1);
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTable bttem = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
                        BlockTableRecord btrMStem = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        ObjectId blkRecIdIDNO = ObjectId.Null;
                        {
                            blkRecIdIDNO = bttem["PcsTitleList2"];
                        }
                        BlockReference acBlkRefIDNO = new BlockReference(PIDNO, blkRecIdIDNO);
                        y1 = y1 - 148;
                        btrMStem.AppendEntity(acBlkRefIDNO);
                        tr.AddNewlyCreatedDBObject(acBlkRefIDNO, true);
                        BlockTableRecord btrIDNO = tr.GetObject(blkRecIdIDNO, OpenMode.ForWrite) as BlockTableRecord;
                        tr.Commit();
                        btrIDNO.DFUpdate(false, true, false);
                    }
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        // Get Block Table and Block Table Record 
                        BlockTable bttem = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord btrMStem = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                        foreach (ObjectId objid in btrMStem)
                        {

                            Entity enttem = (Entity)tr.GetObject(objid, OpenMode.ForWrite, false);

                            if (enttem.GetType() == typeof(BlockReference))
                            {

                                BlockReference blRef = (BlockReference)enttem;
                                AttributeCollection attCol = blRef.AttributeCollection;

                                foreach (ObjectId objID in attCol)

                                {
                                    DBObject dbObj = tr.GetObject(objID, OpenMode.ForRead) as DBObject;
                                    AttributeReference attRef = dbObj as AttributeReference;
                                    switch (attRef.Tag)
                                    {
                                        case "SN1":
                                            attRef.UpgradeOpen();
                                            if (attRef.TextString == "") { attRef.TextString = itetemp.Key.ToString(); }
                                            attRef.DowngradeOpen();
                                            break;
                                        case "PNO1":
                                            attRef.UpgradeOpen();
                                            if (attRef.TextString == "") { attRef.TextString = itetemp.Value.ToString(); }
                                            attRef.DowngradeOpen();
                                            break;
                                            //case "DRIVER":
                                            //    attRef.UpgradeOpen();
                                            //    if (attRef.TextString == "") { attRef.TextString = Pumptype + "-" + (10 + 10 * ii).ToString(); }
                                            //    attRef.DowngradeOpen();
                                            //    break;
                                    }

                                }
                            }
                        }

                        tr.Commit();
                    }
                }
                //set other AttributeDefinition value
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    // Get Block Table and Block Table Record 
                    BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btrMS = (BlockTableRecord)trans.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
                    foreach (ObjectId id in btrMS)
                    {
                        // Fetch the entity 
                        Entity ent = (Entity)trans.GetObject(id, OpenMode.ForWrite, false);
                        // Test if the entity is a BlockReference 
                        if (ent.GetType() == typeof(BlockReference))
                        {
                            BlockReference blRef = (BlockReference)ent;
                            Autodesk.AutoCAD.DatabaseServices.AttributeCollection attCol = blRef.AttributeCollection;
                            foreach (ObjectId objID in attCol)
                            {
                                DBObject dbObj = trans.GetObject(objID, OpenMode.ForRead) as DBObject;
                                AttributeReference acAttRef = dbObj as AttributeReference;
                                switch (acAttRef.Tag)
                                {
                                    case "PUMPTYPE":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = PumpType ;
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "PUMPSIZE":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = PumpSize;
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "DRAWINGNO":
                                        acAttRef.UpgradeOpen();
                                        string strtem2 = "00" + ItemNO;
                                        acAttRef.TextString = SulzerNO + "-" + strtem2.Substring(strtem2.Length - 4);
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "DRAWINGNAME":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = Savename.Replace("_Ver0", "");
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "DRAWN":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[13].ToString();
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "CHECKED":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[14].ToString();
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "APPROVED":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[15].ToString();
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "DTIME":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = SDate;
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "CTIME":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = SDate;
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "ATIME":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = SDate;
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "ORDERNO":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = SulzerNO;
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "PROJECT":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = Project;
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "PURCHASER":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[7].ToString();
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "USER":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[8].ToString();
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "CUST.PO#":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[11].ToString();
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "DEVICESERVICE":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = ServiceName;
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "PUMPOEM":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = "Sulzer";
                                        acAttRef.DowngradeOpen();
                                        break;
                                    case "INSTALLEDAT":
                                        acAttRef.UpgradeOpen();
                                        acAttRef.TextString = form.SUMLISTDATASAT.Tables["DataTableHead"].Rows[0].ItemArray[10].ToString();
                                        acAttRef.DowngradeOpen();
                                        break;
                                }
                            }
                        }
                    }

                    trans.Commit();
                }

                //Purge database
                using (Transaction Tx =db.TransactionManager.StartTransaction())
                {
                    BlockTable table = Tx.GetObject(db.BlockTableId,OpenMode.ForRead) as BlockTable;
                    ObjectIdCollection blockIds = new ObjectIdCollection();
                    do
                    {
                        blockIds.Clear();
                        foreach (ObjectId id in table) { blockIds.Add(id); }
                        db.Purge(blockIds);
                        foreach (ObjectId id in blockIds)
                        {
                            DBObject obj = Tx.GetObject(id, OpenMode.ForWrite);
                            obj.Erase();
                        }
                    } while (blockIds.Count != 0);
                    Tx.Commit();
                }
                //Output dwg file
                HostApplicationServices.WorkingDatabase = workingDB;//切换当前激活的数据库文件
                db.SaveAs(form.SelectedPath + "\\" + Savename + ".dwg", DwgVersion.AC1024);//另存dwg文件
                workingDB.Dispose();
                db.Dispose();
                //begin to plot2PDF 开始打印PDF
                Database workingDB4plot = HostApplicationServices.WorkingDatabase;
                Database db4plot = new Database(false, true);//定义新数据库文件
                try
                {
                    db4plot.ReadDwgFile(form.SelectedPath + "\\" + Savename + ".dwg", System.IO.FileShare.ReadWrite, false, "");
                    db4plot.CloseInput(true);
                    HostApplicationServices.WorkingDatabase = db4plot;
                }
                catch (Autodesk.AutoCAD.Runtime.Exception e)
                {
                    EdPCS.WriteMessage("\nUnable to open .dwg file : " + e.StackTrace);
                    return;
                }
                short bgp = (short)Application.GetSystemVariable("BACKGROUNDPLOT");
                try
                {
                    Application.SetSystemVariable("BACKGROUNDPLOT", 0);

                    using (Transaction tr4plot = db4plot.TransactionManager.StartTransaction())
                    {
                        List<Layout> layouts = new List<Layout>();
                        DBDictionary layoutDict = (DBDictionary)db4plot.LayoutDictionaryId.GetObject(OpenMode.ForRead);
                        foreach (DBDictionaryEntry entry in layoutDict)
                        {
                            if (entry.Key != "Model")
                            {
                                layouts.Add((Layout)tr4plot.GetObject(entry.Value, OpenMode.ForRead));
                            }
                        }
                        layouts.Sort((l1, l2) => l1.TabOrder.CompareTo(l2.TabOrder));

                        string filename = Path.ChangeExtension(form.SelectedPath + "\\" + Savename, "pdf");

                        MultiSheetsPdf plotter = new MultiSheetsPdf(filename, layouts);
                        plotter.Publish();


                        tr4plot.Commit();

                    }
                }
                catch (System.Exception e)
                {
                    //Editor ed = AcAp.DocumentManager.MdiActiveDocument.Editor;
                    EdPCS.WriteMessage("\nError: {0}\n{1}", e.Message, e.StackTrace);
                }
                finally
                {
                    Application.SetSystemVariable("BACKGROUNDPLOT", bgp);
                }

                HostApplicationServices.WorkingDatabase = workingDB4plot;//切换当前激活的数据库文件
                workingDB4plot.Dispose();
                db4plot.Dispose();


            }
            MessageBox.Show("Congratulations!" + "\r\n" + "PCS documents completed.");



        }
        [CommandMethod("PUC")]
        public void PUC()
        {
            SelectPathForm form = new SelectPathForm("PUC", GlobalVariableController.GlobalDataSet);
            form.ShowDialog();

        }
        [CommandMethod("PNT")]
        public void PNT()
        {
            SelectPathForm form = new SelectPathForm("PNT", GlobalVariableController.GlobalDataSet);
            form.ShowDialog();

        }
        [CommandMethod("PDS")]
        public void PDS()
        {
            SelectPathForm form = new SelectPathForm("PDS", GlobalVariableController.GlobalDataSet);
            form.ShowDialog();

        }
        [CommandMethod("PSL")]
        public void PSL()
        {
            SelectPathForm form = new SelectPathForm("PSL", GlobalVariableController.GlobalDataSet);
            form.ShowDialog();

        }

        #region 修改数据画面
        /// <summary>
        /// 修改数据画面作成
        /// </summary>
        [CommandMethod("SumListForm")]
        public void SumListForm()
        {
            try
            {
                ////软件是否可用，计算机名和用户名CHECK
                //string nodeNameComputerName = "computerName";
                //string strMachineName = System.Environment.MachineName;
                //string strUser = System.Environment.UserName;
                //string strDgvComputerName = XMLHelper.ReadSecureInterfaceSettings(nodeNameComputerName);
                //string[] rowDataComputerName = strDgvComputerName.Split(';');
                //bool bolUseFlag = false;
                //for (int i = 0; i < rowDataComputerName.Length; i++)
                //{
                //    string strRow = rowDataComputerName[i];
                //    if (strRow.Length == 0)
                //    {
                //        break;
                //    }
                //    string[] cellDataComputerName = strRow.Split(',');
                //    //计算机名和用户名CHECK
                //    if (strUser.Equals(cellDataComputerName[0]) && strMachineName.Equals(cellDataComputerName[1]))
                //    {
                //        bolUseFlag = true;
                //    }
                //}
                //if(!bolUseFlag)
                //{
                //    //Log写入
                //    LogHelper.GetSingleObj().Error("修改数据画面认证", "没有授权，无法使用");
                //    MessageBox.Show("It cannot be used without authorization!", "ERROR");
                //    return;
                //}
                ButtonUpdata();
                SumListForm form = new SumListForm();
                form.ShowDialog();
                ButtonUpdata();
            }
            catch (System.Exception ex)
            {
                //Log写入
                LogHelper.GetSingleObj().Error("修改数据画面", ex.Message);
                return;
            }
        }
        #endregion

        #region Authority Management frame
        /// <summary>
        /// Authority Management start-up
        /// </summary>
        [CommandMethod("AdministratorForm")]
        public void AdministratorForm()
        {
            AdministratorForm form = new AdministratorForm();
            form.ShowDialog();
        }
        #endregion

        #region Administrator Information frame
        /// <summary>
        /// Administrator Information start-up
        /// </summary>
        [CommandMethod("AdministratorInformationForm")]
        public void AdministratorInformationForm()
        {
            AdministratorInformationForm form = new AdministratorInformationForm();
            form.ShowDialog();
        }
        #endregion

        #region Password Confirmation frame
        /// <summary>
        /// Password Confirmation start-up
        /// </summary>
        [CommandMethod("PasswordConfirmationForm")]
        public void PasswordConfirmationForm()
        {
            PasswordConfirmationForm form = new PasswordConfirmationForm();
            form.ShowDialog();
        }
        #endregion

        #region Version InformationForm frame
        /// <summary>
        /// Version InformationForm start-up
        /// </summary>
        [CommandMethod("VersionInformationForm")]
        public void VersionInformationForm()
        {
            VersionInformationForm form = new VersionInformationForm();
            form.ShowDialog();
        }
        #endregion

    }
}
