using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace WordAddin_Right_Click_Menu
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            RemoveAddedMenuItems();
            AddRightClickMenuItems();
            this.Application.WindowBeforeRightClick+=new Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(Application_WindowBeforeRightClick);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void AddRightClickMenuItems()
        {
            Office.CommandBarButton AddBtn = null;
            AddBtn = (Office.CommandBarButton)Application.CommandBars["Text"].Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, true);
            AddBtn.Tag = "share";
            AddBtn.Caption = "分享";
            AddBtn.Enabled = false;
            AddBtn.Click+=new Office._CommandBarButtonEvents_ClickEventHandler(AddBtn_Click);
        }
        private void RemoveAddedMenuItems()
        {
            //获取
            Office.CommandBarButton btn = null;
            do 
            {
                if(btn!=null)
                    btn.Delete(true);
                btn = (Office.CommandBarButton)Application.CommandBars["Text"].FindControl(Office.MsoControlType.msoControlButton, missing, "share", missing, false);
            } while (btn!=null);
        }
        private void Application_WindowBeforeRightClick(Word.Selection Sel , ref bool Cancel)
        {
            //如果sel的文本不为空,则激活按钮
            if (!string.IsNullOrWhiteSpace(Sel.Range.Text))
            {
                Office.CommandBarButton AddBtn = (Office.CommandBarButton)Application.CommandBars["Text"].FindControl(Office.MsoControlType.msoControlButton,missing,"share",missing,missing);
                AddBtn.Enabled = true;
            }
        }
        private void AddBtn_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            MessageBox.Show("菜单按钮响应");
        }
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
