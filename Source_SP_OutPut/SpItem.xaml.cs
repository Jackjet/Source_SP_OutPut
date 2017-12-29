using ConferenceCommon.WebHelper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Source_SP_OutPut
{
    /// <summary>
    /// SpItem.xaml 的交互逻辑
    /// </summary>
    public partial class SpItem : UserControl
    {
        public WebCredentialManage web = null;

        public SpItem(string userName,string password)
        {
            InitializeComponent();

            web = new WebCredentialManage(this.webBrowser, userName, password);
        }
    }
}
