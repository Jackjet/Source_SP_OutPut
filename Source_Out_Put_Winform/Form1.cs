using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Source_Out_Put_Winform
{
    public partial class Form1 : Form
    {
      

        public Form1()
        {
            InitializeComponent();

            //DBManage.sqlConnection = "Server=192.168.1.49;Database = YanQingYiZhongDB;Uid=sa;Pwd =sa123??;";

            //string sql = string.Format("insert into  [dbo].[PortalTreeData](Name,Display,IsDelete,CreateTime,Creator,PId,BeforeUrl,BeforeAfter,AfterUrl,SortId,EnName ) values('{0}',0,0,getDate(),'',0,'',0,'',0,'')", );
            //string error = "";


          
          

            //DBManage.Transaction(sql, out error);
        }
    }
}
