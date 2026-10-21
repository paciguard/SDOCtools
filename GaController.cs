using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Customization;
/*AutoCAD*/
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;
using Microsoft.Office.Interop.Excel;
using CAdapp = Autodesk.AutoCAD.ApplicationServices.Application;
using Excel = Microsoft.Office.Interop.Excel;

namespace SESDTools
{
    /// <summary>
    /// 类    名：GaController
    /// 功能描述：制作外形图功能类
    /// 创 建 人：SY
    /// 创建时间：2021/07/29 10:43
    /// 描    述：制作外形图
    /// 修 改 人：
    /// 修改时间：
    /// 修改描述：
    /// 版    本：V1.0
    /// </summary>
    public class GaController
    {
        #region GA功能
        /// <summary>
        /// 外形图GA功能
        /// </summary>
        [CommandMethod("GA")]
        public void GA()
        {

        }
        #endregion
    }
}
