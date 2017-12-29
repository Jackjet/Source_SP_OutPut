using ConferenceCommon.FileDownAndUp;
using ConferenceCommon.SharePointHelper;
using ConferenceCommon.TimerHelper;
using ConferenceCommon.WebHelper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using spClient = Microsoft.SharePoint.Client;
using appsettings = System.Configuration.ConfigurationManager;
using ConferenceCommon.LogHelper;


namespace Source_SP_OutPut
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        #region 字段

        static ClientContextManage client = new ClientContextManage();


        static string webSite = appsettings.AppSettings["webSite"];

        static string beforeImageSite = appsettings.AppSettings["beforeImageSite"];

        string userName = appsettings.AppSettings["userName"];
        string password = appsettings.AppSettings["password"];
        string domain = appsettings.AppSettings["domain"];

        int imageDownloadSpeed = Convert.ToInt32(appsettings.AppSettings["imageDownloadSpeed"]);
        int dicForeachSpeed = Convert.ToInt32(appsettings.AppSettings["dicForeachSpeed"]);
        string htmlBody = appsettings.AppSettings["htmlBody"];


        bool NeedInsertDB = Convert.ToBoolean(appsettings.AppSettings["NeedInsertDB"]);

        string rootPart = appsettings.AppSettings["rootPart"];

        TextBlock txtBefore = null;

        int nextCount = 0;

        Dictionary<int, string> rootDic = new Dictionary<int, string>();

        #endregion

        #region 构造函数

        public MainWindow()
        {
            InitializeComponent();



            try
            {

                LogManage.LogInit();
                //DBManage.sqlConnection = "Server=117.106.85.18;Database = YanQingYiZhongDB;Uid=sa;Pwd =sa@2016";
                //  DBManage.sqlConnection = "Server=192.168.1.72;Database = DaYuYiXiaoDB;Uid=sa;Pwd =sa123??;";
                DBManage.sqlConnection = System.Configuration.ConfigurationManager.ConnectionStrings["contr"].ToString();


                //cmb_Config();


                spClient.ListCollection listCollection = client.GetAllLists(webSite, userName, password, domain);
                foreach (var item in listCollection)
                {
                    if (item.Title == "123" || item.Title == "表单模板" || item.Title == "网站页面" || item.Title == "网站资产" ||
                     item.Title == "微源" || item.Title == "文档" || item.Title == "样式库" || item.Title == "左侧导航" || item.Title == "appdata"
                        || item.Title == "fpdatasources" || item.Title == "TaxonomyHiddenList" || item.Title == "Web 部件库" || item.Title == "wfpub" || item.Title == "左侧导航"
                        || item.Title == "列表模板库" || item.Title == "母版页样式库" || item.Title == "内容类型发布错误日志"
                        || item.Title == "项目策略项列表" || item.Title == "解决方案库" || item.Title == "已转换表单" || item.Title == "组合外观"
                        )
                    {

                        continue;
                    }
                    TextBlock txt = new TextBlock() { Text = item.Title };
                    txt.MouseLeftButtonDown += txt_MouseLeftButtonDown;
                    txt.MouseRightButtonDown += txt_MouseRightButtonDown;
                    txt.Tag = item;
                    listBox.Items.Add(txt);
                }
            }
            catch (Exception ex)
            {

                LogManage.WriteLog(this.GetType(), ex);
            }

            //foreach (var item in listBox.Items)
            //{
            //    Thread.Sleep(5000);
            //    this.txt_MouseLeftButtonDown(item, null);
            //}
            //cmb_Config();
        }
        #endregion


        void txt_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            nextCount = this.listBox.Items.IndexOf((sender as TextBlock));
        }

        void txt_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            BeginOutPut(sender, null);
        }

        #region 开始导出

        public void BeginOutPut(object sender, Action callback)
        {
            TextBlock txt = sender as TextBlock;
            if (txtBefore != null)
            {
                txtBefore.Background = new SolidColorBrush(Colors.White);
            }


            txt.Background = new SolidColorBrush(Colors.SkyBlue);

            txtBefore = txt;
            if (txt.Tag != null)
            {
                spClient.List list = txt.Tag as spClient.List;


                #region 执行过程

                cm cm = GetCm(list.Title);
                //if (cm != null)
                if (cm != null)
                {
                    List<Dictionary<string, object>> diclist = client.ClientGetDic(webSite, list.Title, "");
                    if (diclist != null && diclist.Count > 0)
                    {
                        this.listB2.Items.Clear();
                        int count = 0;
                        foreach (var item in diclist)
                        {
                            SpItem spItem = new SpItem(userName, password);
                            if (item.ContainsKey("Title"))
                            {
                                string contentType = "AdvertImgContent";


                                string txtD = Convert.ToString(list.Title);
                                spItem.txt.Text = txtD;


                                //if (txtD == "党员风采" || txtD == "多彩活动" || txtD == "行政领导" || txtD == "语文教师" || txtD == "数学教师" || txtD == "英语教师"
                                //   || txtD == "科任教师" || txtD == "幼儿园教师" || txtD == "后勤职员" || txtD == "教师风采")
                                //{
                                //    contentType = "SchoolStyle";
                                //}
                                //else
                                //{
                                //    contentType = "AdvertImgContent";
                                //}

                                //FileDownLoad(list, contentType);


                                spItem.web.Navicate(Environment.CurrentDirectory + "\\HTMLPage1.html");
                                TimerJob.StartRun(new Action(() =>
                                {
                                    string fileList = "";
                                    string fileNames = "";
                                    int id = Convert.ToInt32(item["ID"]);
                                    if (rootDic.ContainsKey(id))
                                    {
                                        string oldfileList = rootDic[id];

                                        string[] dddd = oldfileList.Split(new char[] { ',' });

                                        foreach (var d in dddd)
                                        {
                                            if (!string.IsNullOrEmpty(d))
                                            {
                                                string resalFileName = System.IO.Path.GetFileName(d);

                                                fileNames += resalFileName + ",";

                                                fileList += "/" + rootPart + "/Attatchment/" + contentType + "/" + resalFileName + ",";
                                            }
                                        }
                                    }


                                    if (item.ContainsKey(htmlBody) && !string.IsNullOrEmpty(Convert.ToString(item[htmlBody])))
                                    {
                                        string html = Convert.ToString(item[htmlBody]);

                                        string[] imgUrils = GetHtmlImageUrlList(html);




                                        html = ImageDownload(html, contentType, imgUrils, new Action(() =>
                                        {
                                            //try
                                            //{
                                            //    spItem.webBrowser.Document.Body.InnerHtml = html;
                                            //}
                                            //catch (Exception)
                                            //{
                                            //}

                                            string uriArray = "";
                                            for (int i = 0; i < imgUrils.Count(); i++)
                                            {
                                                if (html.Contains(imgUrils[i]))
                                                {
                                                    string img_uri = System.IO.Path.GetFileName(imgUrils[i]).Replace("%", "_");

                                                    string fln = "/" + rootPart + "/Attatchment/" + contentType + "/" + img_uri;
                                                    uriArray += fln + ",";
                                                    html = html.Replace(imgUrils[i], fln);
                                                }
                                            }


                                            #region 插入数据库

                                            if (contentType == "AdvertImgContent")
                                            {
                                                this.Insert(item, html, cm.ID, uriArray, fileNames, fileList);
                                            }
                                            else if (contentType == "SchoolStyle")
                                            {
                                                this.Insert2(item, html, cm.ID, uriArray, fileNames, fileList);

                                            }
                                            count++;
                                            if (count == diclist.Count)
                                            {
                                                if (callback != null)
                                                {
                                                    this.Dispatcher.BeginInvoke(new Action(() =>
                                                    {
                                                        if (checkBox.IsChecked == true)
                                                        {
                                                            nextCount++;
                                                            try
                                                            {
                                                                BeginOutPut(this.listBox.Items[nextCount], () =>
                                                                {
                                                                });
                                                            }
                                                            catch (Exception)
                                                            {
                                                            }

                                                        }
                                                    }));

                                                }
                                                //MessageBox.Show("完成导出");
                                            }
                                            #endregion
                                        }


                                        ));
                                    }
                                    else
                                    {
                                        #region 插入数据库

                                        if (contentType == "AdvertImgContent")
                                        {
                                            this.Insert(item, "", cm.ID, "", fileNames, fileList);
                                        }
                                        else if (contentType == "SchoolStyle")
                                        {
                                            this.Insert2(item, "", cm.ID, "", fileNames, fileList);
                                        }
                                        count++;
                                        if (count == diclist.Count)
                                        {
                                            if (callback != null)
                                            {
                                                this.Dispatcher.BeginInvoke(new Action(() =>
                                                {
                                                    if (checkBox.IsChecked == true)
                                                    {
                                                        nextCount++;
                                                        try
                                                        {
                                                            BeginOutPut(this.listBox.Items[nextCount], () =>
                                                            {
                                                            });
                                                        }
                                                        catch (Exception)
                                                        {
                                                        }

                                                    }
                                                }));

                                            }
                                            //MessageBox.Show("完成导出");
                                        }
                                        #endregion
                                    }




                                }), 200);

                                listB2.Items.Add(spItem);
                            }
                        }
                    }

                    else
                    {
                        if (checkBox.IsChecked == true)
                        {
                            nextCount++;
                            BeginOutPut(this.listBox.Items[nextCount], () =>
                            {
                            });
                        }
                    }
                }
                else
                {
                    if (checkBox.IsChecked == true)
                    {
                        nextCount++;
                        BeginOutPut(this.listBox.Items[nextCount], () =>
                        {
                        });
                    }
                }

                #endregion

            }
        }

        #endregion

        //private int BeginSingle(Action callback, spClient.List list, cm cm, List<Dictionary<string, object>> diclist, int count, Dictionary<string, object> item)
        //{
        //    return count;
        //}


        #region 图片下载

        private string ImageDownload(string html, string contentType, string[] imgUrils, Action callback)
        {
            int imgDownload_count = 0;

            if (checkBoxNeedDownloadImage.IsChecked == true)
            {

                for (int i = 0; i < imgUrils.Count(); i++)
                {
                    if (html.Contains(imgUrils[i]))
                    {
                        html = html.Replace(imgUrils[i], beforeImageSite + imgUrils[i].Replace("%", "_"));
                    }
                    Thread.Sleep(imageDownloadSpeed);

                    string imguri = System.IO.Path.GetFileName(imgUrils[i]);
                    imguri = imguri.Replace("%", "_");
                    WebClientManage webClient = new WebClientManage();

                    if (imguri == "20169129358470.jpg" || imguri == "201691293451187.jpg" || imguri == "2016913112928749.jpg")
                    {

                    }
                    webClient.FileDown(beforeImageSite + imgUrils[i], Environment.CurrentDirectory + "\\" + contentType + "\\" + imguri, userName, password, domain, new Action<int>((r) =>
                    {
                    }), new Action<Exception, bool>((errir, isSuccessed) =>
                    {
                        imgDownload_count++;
                        if (isSuccessed && imgDownload_count == imgUrils.Count())
                        {
                            callback();
                        }
                    }));
                }
                if (imgUrils.Count() == 0)
                {
                    callback();
                }
            }
            else
            {
                callback();
            }
            return html;
        }

        #endregion

        #region 文件下载

        /// <summary>
        /// 下载附件
        /// </summary>
        /// <param name="list"></param>
        /// <param name="contentType"></param>
        private void FileDownLoad(spClient.List list, string contentType)
        {
            bool result = client.LoadMethod(list.RootFolder);
            if (!result) return;
            client.LoadMethod(list.RootFolder.Folders);

            foreach (var ccp in list.RootFolder.Folders)
            {
                client.LoadMethod(ccp);
                if (ccp.Name.Equals("Attachments"))
                {
                    rootDic.Clear();
                    client.LoadMethod(ccp.Folders);

                    foreach (var itemChild in ccp.Folders)
                    {
                        var t = itemChild.Name;
                        client.LoadMethod(itemChild.Files);
                        string files = null;
                        foreach (var file in itemChild.Files)
                        {
                            client.LoadMethod(file);

                            Thread.Sleep(100);
                            WebClientManage webClient = new WebClientManage();
                            webClient.FileDown(beforeImageSite + file.ServerRelativeUrl,
                                Environment.CurrentDirectory + "\\" + contentType + "\\" +
                                System.IO.Path.GetFileName(file.ServerRelativeUrl), userName, password, domain, new Action<int>((r) =>
                            {
                            }), new Action<Exception, bool>((errir, isSuccessed) =>
                            {
                            }));

                            files += file.ServerRelativeUrl + ",";
                        }
                        if (!string.IsNullOrEmpty(files))
                        {
                            rootDic.Add(Convert.ToInt32(t), files);
                        }
                    }
                    break;
                }
            }
        }

        #endregion

        #region 表现html元素

        /// <summary> 
        /// 取得HTML中所有图片的 URL。 
        /// </summary> 
        /// <param name="sHtmlText">HTML代码</param> 
        /// <returns>图片的URL列表</returns> 
        public string[] GetHtmlImageUrlList(string sHtmlText)
        {
            // 定义正则表达式用来匹配 img 标签 
            Regex regImg = new Regex(@"<img\b[^<>]*?\bsrc[\s\t\r\n]*=[\s\t\r\n]*[""']?[\s\t\r\n]*(?<imgUrl>[^\s\t\r\n""'<>]*)[^<>]*?/?[\s\t\r\n]*>", RegexOptions.IgnoreCase);

            // 搜索匹配的字符串 
            MatchCollection matches = regImg.Matches(sHtmlText);
            int i = 0;
            string[] sUrlList = new string[matches.Count];

            // 取得匹配项列表 
            foreach (Match match in matches)
                sUrlList[i++] = match.Groups["imgUrl"].Value;
            return sUrlList;
        }

        #endregion

        #region 获取数据库已有的菜单项进行匹配

        public cm GetCm(string title)
        {
            cm cm = null;
            string sql = string.Format("select * from PortalTreeData where  PortalTreeData.Name ='{0}'", title);
            string error = "";
            List<cm> cmList = DBManage.ExcuteEntity<cm>(sql, System.Data.CommandType.Text, out error);
            if (cmList.Count > 0)
            {
                cm = cmList[0];
            }
            return cm;
        }

        #endregion

        #region 数据库插入数据

        public void Insert(Dictionary<string, object> dicList, string html, int menuID, string imgUri, string filename, string filePath)
        {
            if (NeedInsertDB)
            {
                int modeType = 0;
                if (dicList.ContainsKey("_x6a21__x7248_"))
                {
                    if (dicList["_x6a21__x7248_"] != null)
                    {
                        modeType = 1;
                    }
                }


                Microsoft.SharePoint.Client.FieldUserValue vvv = dicList["Author"] as Microsoft.SharePoint.Client.FieldUserValue;
                string creator = vvv.LookupValue;
                html = html.Replace("'", "''");
                string sql = string.Format("insert into Advertising(MenuId,[Description], CreativeHTML,CreateTime,ImageUrl,ClickNum,ModelType,IsDelete,Creator,FileName,FilePath,isPush)  values({0},'{1}','{2}','{3}','{4}',{5},{6},{7},'{8}','{9}','{10}',{11})",
                     menuID, dicList["Title"], html, dicList["Created"], imgUri, dicList["Count"], modeType, 0, creator, filename, filePath, 1);
                string error = "";

                DBManage.Transaction(sql, out error);
            }

        }

        public void Insert2(Dictionary<string, object> dicList, string html, int menuID, string imgUri, string filename, string filePath)
        {
            if (NeedInsertDB)
            {
                string sql1 = string.Format("select * from dbo.SchoolStyle where dbo.SchoolStyle.Description ='{0}'", dicList["Title"]);
                string error1 = "";
                List<cm> cmList = DBManage.ExcuteEntity<cm>(sql1, System.Data.CommandType.Text, out error1);
                if (cmList.Count < 1)
                {
                    int modeType = 0;
                    if (dicList.ContainsKey("_x6a21__x7248_"))
                    {
                        if (dicList["_x6a21__x7248_"] != null)
                        {
                            modeType = 1;
                        }
                    }
                    Microsoft.SharePoint.Client.FieldUserValue vvv = dicList["Author"] as Microsoft.SharePoint.Client.FieldUserValue;
                    string creator = vvv.LookupValue;
                    html = html.Replace("'", "''");
                    string sql = string.Format("insert into SchoolStyle(MenuId,Title,[Description],CreateTime,ImageUrl,ClickNum,ModelType,IsDelete,Creator,FileName,FilePath)  values({0},'{1}','{2}','{3}','{4}',{5},{6},{7},'{8}','{9}','{10}')",
                       menuID, dicList["Title"], html, dicList["Created"], imgUri, dicList["Count"], modeType, 0, creator, filename, filePath);
                    string error = "";

                    DBManage.Transaction(sql, out error);
                }
            }
        }

        #endregion

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            BeginOutPut(this.listBox.Items[nextCount], () =>
            {
            });
        }

        #region 配置菜单并导入到数据库

        public void cmb_Config()
        {
            List<string> list = new List<string>(){
                "首页",

                "学校概况",
                "学校简介",
                "现任领导",
                "师资队伍",
                "学校荣誉",
                "校貌照片",

                "管理机构",
                "支部班子",
                "行政班子",
                "工会班子",

                "党务工作",
                "制度建设",
                "党建论文",
                "专题活动",
                "组织活动",
                "学习型组织",

                "工会工作",
                "教代会",
                "工会之家",
                "文艺活动",

                "团委工作",
                "学校党校",
                "学生会",
                "组织活动",

                "德育工作",
                "规章制度",
                "德育队伍",
                "家教协会",
                "教育活动",
                "德育论文",
                "紫金杯",
                "体卫工作",
                "社会实践",
                "视频照片",

                "教学科研",
                "教学组织",
                "教学制度",
                "教学论文",
                "师资培训",
                "教研活动",
                "精品教案",
                "专业活动",
                "教师风采",
                "视频照片",

                "招生就业",
                "专业介绍",
                "招生简章",
                "实习工作",
                "优秀毕业生",
                 "高等学历",

                "数字校园",
                "企业内网",               
       
          };







            foreach (var item in list)
            {
                InsertCmb(item);
            }





        }

        public void InsertCmb(string _displayName)
        {
            string sql = string.Format("insert into  [dbo].[PortalTreeData](Name,Display,IsDelete,CreateTime,Creator,PId,BeforeUrl,BeforeAfter,AfterUrl,SortId,EnName,DisplayCount,DisplayType ) values('{0}',0,0,getDate(),'',0,'/YQZJ/SitePages/BeforeItemList.aspx?',2,'/admin/AfterList.aspx?id=',0,'',8,'时间')", _displayName);
            string error = "";
            DBManage.Transaction(sql, out error);

        }

        #endregion
    }

}

public class cm
{
    public int ID { get; set; }

    public string Name { get; set; }
}



