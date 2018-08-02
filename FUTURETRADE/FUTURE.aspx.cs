using HtmlAgilityPack;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.Web.Services;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace FUTURETRADE
{
    //頁面回傳
    public class QueryInfo
    {
        public string ErrMsg { get; set; }
        public string Top10ErrMsg { get; set; }
        public List<ProdInfo> ListProdInfo { get; set; }
        public string ProdInfoJson { get; set; }
        public List<Top10ProdInfo> ListTop10ProdInfo { get; set; }
        public List<QueryData> ListQueryData { get; set; }
    }
    //Top10
    public class Top10ProdInfo
    {
        public string ProdNo { get; set; }      //商品編號
        public string ProdName { get; set; }        //商品名稱
        public string Buy { get; set; }      //買價
        public string Sell { get; set; }        //賣價
        public string UpDown { get; set; }      //漲跌
        public string UpDownPercent { get; set; }      //漲跌幅
        public string Trade { get; set; }       //成交量
        public string DataDate { get; set; }       //資料時間
    }
    //Top10 欄位顏色
    public class Top10CellColor
    {
        public string BgColor { get; set; }     //背景
        public string Color { get; set; }      //顏色
    }
    //保證金資料
    public class ProdInfo
    {
        public string ProdNo { get; set; }      //商品編號
        public string ProdName { get; set; }        //商品名稱
        public string ProdId { get; set; }      //商品英文ID
        public string ProdType { get; set; }        //商品類別(大,小型)
        public string Margin { get; set; }      //原始保證金(% or 金額)
        public string Qty { get; set; }     //股數(口)
        public string QueryDate { get; set; }       //查詢日期時間
    }
    //查詢資料
    public class QueryData
    {
        public string ProdId;       //商品代碼
        public string ProdName;     //商品名稱
        public string ProdNo;       //商品代號
        public string MonthTrade;       //近月成交價
        public string Qty;      //股數/口
        public string Margin;       //原始保證金比例
        public string TradeMargin;      //交易需要保證金
        public string QueryDate;        //查詢時間
        public string ProdType;     //商品分類 0/1/2
    }
    //檔案架構
    public class FileData
    {
        public FileStream Stream;
        public int Length;
        public byte[] ByteData;
    }

    public partial class FUTURE : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                //更新時間設定值
                this.hdTimer.Value = ConfigurationManager.AppSettings["Timer"].ToString();
                this.hdStartTimer.Value = ConfigurationManager.AppSettings["StartTimer"].ToString();
                this.hdEndTimer.Value = ConfigurationManager.AppSettings["EndTimer"].ToString();

                if (!Directory.Exists(ConfigurationManager.AppSettings["MarginPath"].ToString()))
                {
                    Directory.CreateDirectory(ConfigurationManager.AppSettings["MarginPath"].ToString());
                }

                if (!File.Exists(ConfigurationManager.AppSettings["MarginPath"].ToString() + ConfigurationManager.AppSettings["MarginTop10TxtName"].ToString()))
                {
                    File.Create(ConfigurationManager.AppSettings["MarginPath"].ToString() + ConfigurationManager.AppSettings["MarginTop10TxtName"].ToString()).Close();
                }

            }
        }

        [WebMethod(Description = "台灣期交所股票類Top10")]
        public static QueryInfo QueryTop10ProdInfo(string ProdInfo)
        {
            QueryInfo _queryInfo = new QueryInfo();
            List<Top10ProdInfo> _listTop10ProdInfo = new List<Top10ProdInfo>();

            string errMsg = "";
            string Top10ErrMsg = "";

            try
            {
                _listTop10ProdInfo = CreateTop10Trade(ProdInfo, ref Top10ErrMsg);
            }
            catch (Exception ex)
            {
                _listTop10ProdInfo = null;
                errMsg = ex.Message;
            }

            _queryInfo.ErrMsg = errMsg;
            _queryInfo.Top10ErrMsg = Top10ErrMsg;
            _queryInfo.ListTop10ProdInfo = _listTop10ProdInfo;

            return _queryInfo;

        }
        [WebMethod(Description = "保證金取得")]
        public static QueryInfo QueryMargin(string Prod)
        {
            QueryInfo _queryInfo = new QueryInfo();
            List<ProdInfo> _listProdInfo = new List<ProdInfo>();

            string errMsg = "";

            try
            {
                //讀檔取得
                //_listProdInfo = CreateMarginingListForTxt(Prod);
                //連網取得
                _listProdInfo = CreateMarginingList(Prod);
            }
            catch (Exception ex)
            {
                _listProdInfo = null;
                errMsg = ex.Message;
            }

            _queryInfo.ErrMsg = errMsg;
            _queryInfo.ListProdInfo = _listProdInfo;
            _queryInfo.ProdInfoJson = JsonConvert.SerializeObject(_listProdInfo);
            return _queryInfo;

        }
        [WebMethod(Description = "查詢資料取得")]
        public static QueryInfo QueryData(string Prod)
        {
            QueryInfo _queryInfo = new QueryInfo();
            List<QueryData> _listQueryData = new List<QueryData>();

            string errMsg = "";

            try
            {
                _listQueryData = CreateQueryData(Prod);
            }
            catch (Exception ex)
            {
                _listQueryData = null;
                errMsg = ex.Message;
            }

            _queryInfo.ErrMsg = errMsg;
            _queryInfo.ListQueryData = _listQueryData;
            return _queryInfo;

        }
        //建立股票期貨Top 10
        protected static List<Top10ProdInfo> CreateTop10Trade(string ProdInfo, ref string Top10ErrMsg)
        {
            string htmlUrl = "http://info512.taifex.com.tw/Future/FusaQuote_Norl_Top1.aspx";
            string htmlPage = "";

            List<Top10ProdInfo> _listTop10ProdInfo = new List<Top10ProdInfo>();
            Top10ProdInfo _Top10ProdInfo = null;
            List<Top10CellColor> _listTop10CellColor = null;
            Top10CellColor _Top10CellColor = null;

            //保證金資料
            List<ProdInfo> _listProdInfo = new List<ProdInfo>();
            _listProdInfo = JsonConvert.DeserializeObject<List<ProdInfo>>(ProdInfo);

            HtmlDocument doc = new HtmlDocument();
            HtmlDocument docTable = null;
            HtmlDocument docTR = null;

            HtmlNodeCollection divNode = null;
            HtmlNodeCollection tableNode = null;
            HtmlNodeCollection trNode = null;

            string[] sourceRow = null;
            List<string> RowColor = null;
            string strProdNo = "";
            double dUpDown = 0;
            double dPrice = 0;

            string fileFolder = ConfigurationManager.AppSettings["MarginPath"].ToString();
            string fileName = ConfigurationManager.AppSettings["MarginPath"].ToString() + ConfigurationManager.AppSettings["MarginTop10TxtName"].ToString();

            //get 台灣期交所 熱門股票期貨
            using (WebClient webClient = new WebClient())
            {
                webClient.Encoding = Encoding.Default;
                htmlPage = webClient.DownloadString(htmlUrl);
            }

            doc.LoadHtml(htmlPage);
            divNode = doc.DocumentNode.SelectNodes("//div[@id='divDG']");

            docTable = new HtmlDocument();
            docTable.LoadHtml(divNode[0].InnerHtml);
            tableNode = docTable.DocumentNode.SelectNodes("//table");

            docTR = new HtmlDocument();
            docTR.LoadHtml(tableNode[0].InnerHtml.ToString());
            trNode = docTR.DocumentNode.SelectNodes("//tr");

            List<string> colorDetail;
            var rowsColor = new List<List<string>>();

            //建立資料列TEXT
            //抬頭欄位
            //商品 買價 買量 賣價 賣量 成交價 漲跌 振幅％ 成交量 開盤 最高 最低 參考價 時間
            //  0       1       2       3       4       5          6      7           8         9      10     11     12        13   (陣列數)
            //=======================================
            IEnumerable<string[]> rows = trNode.Skip(1).Select(rTR => rTR
                 .Elements("td")
                 .Select(td => td.InnerText.Trim())
                 .ToArray());

            //建立資料列HTML,以利取得顏色
            for (int i = 1; i < trNode.Count; i++)
            {
                var a = trNode[i];
                var b = a.ChildNodes;
                colorDetail = new List<string>();
                for (int j = 1; j < b.Count; j++)
                {
                    var c = b[j];
                    var d = (c.OuterHtml.IndexOf("bgcolor") > -1) ? c.OuterHtml.Substring(c.OuterHtml.IndexOf("bgcolor") + 9, 7) : "";
                    var e = (c.InnerHtml.IndexOf("color") > -1) ? c.InnerHtml.Substring(c.InnerHtml.IndexOf("color") + 7, 7) : "#000000";
                    colorDetail.Add(d + "|" + e);
                }
                rowsColor.Add(colorDetail);
            }

            //文字敍述
            string BeforecloseTime = DateTime.Now.AddDays(-1).ToString("yyyy/MM/dd") + " 13:45:00";
            BeforecloseTime = "前一營業日  13:45:00";

            string NowcloseTime = DateTime.Now.ToString("yyyy/MM/dd") + " 13:45:00";

            if (DateTime.Now >= DateTime.Parse(DateTime.Now.ToString("yyyy/MM/dd") + " 00:00:00") &&
                DateTime.Now <= DateTime.Parse(DateTime.Now.ToString("yyyy/MM/dd") + "  " + ConfigurationManager.AppSettings["StartTimer"].ToString()))
            {
                Top10ErrMsg = "今日尚未開盤，資料時間：" + BeforecloseTime;
            }

            if (DateTime.Now <= DateTime.Parse(DateTime.Now.AddDays(1).ToString("yyyy/MM/dd") + " 00:00:00") &&
                DateTime.Now >= DateTime.Parse(DateTime.Now.ToString("yyyy/MM/dd") + "  " + ConfigurationManager.AppSettings["EndTimer"].ToString()))
            {
                Top10ErrMsg = "今日已收盤，資料時間：" + NowcloseTime;
            }

            //收盤時間寫入文字檔
            if (DateTime.Now >= DateTime.Parse(DateTime.Now.ToString("yyyy/MM/dd") + " 13:44:50") &&
                DateTime.Now <= DateTime.Parse(DateTime.Now.ToString("yyyy/MM/dd") + " 13:45:30"))
            {
                using (FileStream fsStream = new FileStream(fileName, FileMode.Create))
                using (BinaryWriter writer = new BinaryWriter(fsStream, Encoding.Default))
                {
                    writer.Write(htmlPage);
                }
            }

            if (rows.Count() == 0)
            {
                if (DateTime.Now >= DateTime.Parse(DateTime.Now.ToString("yyyy/MM/dd") + " " + ConfigurationManager.AppSettings["StartTimer"].ToString()) &&
                    DateTime.Now <= DateTime.Parse(DateTime.Now.ToString("yyyy/MM/dd") + " " + ConfigurationManager.AppSettings["StartRangeTimer"].ToString()))
                {
                    Top10ErrMsg = "今日尚未開盤";
                    return null;
                }
                else
                {
                    if (DateTime.Now >= DateTime.Parse(DateTime.Now.ToString("yyyy/MM/dd") + " " + ConfigurationManager.AppSettings["StartTimer"].ToString()) &&
                        DateTime.Now <= DateTime.Parse(DateTime.Now.ToString("yyyy/MM/dd") + " " + ConfigurationManager.AppSettings["EndTimer"].ToString()))
                    {
                        if (ConfigurationManager.AppSettings["EmergencyFlag"].ToString().ToUpper() == "TRUE")
                        {
                            Top10ErrMsg = ConfigurationManager.AppSettings["EmergencyMsg"].ToString();
                            return null;
                        }
                        else
                        {
                            Top10ErrMsg = "";
                            return null;
                        }
                    }
                    rowsColor = new List<List<string>>();

                    //read file
                    htmlPage = ReadTopFile(fileName);
                    if (htmlPage == null || htmlPage == "") { throw new Exception("尚無資料!"); }

                    doc.LoadHtml(htmlPage);
                    divNode = doc.DocumentNode.SelectNodes("//div[@id='divDG']");

                    docTable = new HtmlDocument();
                    docTable.LoadHtml(divNode[0].InnerHtml);
                    tableNode = docTable.DocumentNode.SelectNodes("//table");

                    docTR = new HtmlDocument();
                    docTR.LoadHtml(tableNode[0].InnerHtml.ToString());
                    trNode = docTR.DocumentNode.SelectNodes("//tr");

                    //建立資料列TEXT
                    //抬頭欄位
                    //商品 買價 買量 賣價 賣量 成交價 漲跌 振幅％ 成交量 開盤 最高 最低 參考價 時間
                    //  0       1       2       3       4       5          6      7           8         9      10     11     12        13   (陣列數)
                    //=======================================
                    IEnumerable<string[]> rowsRest = trNode.Skip(1).Select(rTR => rTR
                         .Elements("td")
                         .Select(td => td.InnerText.Trim())
                         .ToArray());

                    //建立資料列HTML,以利取得顏色
                    for (int i = 1; i < trNode.Count; i++)
                    {
                        var a = trNode[i];
                        var b = a.ChildNodes;
                        colorDetail = new List<string>();
                        for (int j = 1; j < b.Count; j++)
                        {
                            var c = b[j];
                            var d = (c.OuterHtml.IndexOf("bgcolor") > -1) ? c.OuterHtml.Substring(c.OuterHtml.IndexOf("bgcolor") + 9, 7) : "";
                            var e = (c.InnerHtml.IndexOf("color") > -1) ? c.InnerHtml.Substring(c.InnerHtml.IndexOf("color") + 7, 7) : "#000000";
                            colorDetail.Add(d + "|" + e);
                        }
                        rowsColor.Add(colorDetail);
                    }
                    rows = rowsRest;
                }
            }


            for (int i = 0; i < rows.Count(); i++)
            {
                sourceRow = rows.ToList()[i];
                RowColor = rowsColor[i];

                //取得顏色
                _listTop10CellColor = new List<Top10CellColor>();

                foreach (string color in RowColor)
                {
                    _Top10CellColor = new Top10CellColor();
                    _Top10CellColor.BgColor = color.Split('|')[0];
                    _Top10CellColor.Color = color.Split('|')[1];
                    _listTop10CellColor.Add(_Top10CellColor);
                }

                dUpDown = double.Parse(sourceRow[6].ToString().Trim());      //漲跌
                dPrice = double.Parse(sourceRow[12].ToString().Trim());      //參考價

                //search 商品代號
                foreach (ProdInfo p in _listProdInfo)
                {
                    if (sourceRow[0].ToString().Trim().IndexOf(p.ProdName) > -1)
                    {
                        strProdNo = p.ProdNo;
                        break;
                    }
                }

                _Top10ProdInfo = new Top10ProdInfo
                {
                    ProdNo = strProdNo,     //商品代號
                    ProdName = sourceRow[0].ToString().Trim(),      //商品名稱
                    Buy = sourceRow[1].ToString().Trim(),       //買價
                    Sell = sourceRow[3].ToString().Trim(),      //賣價
                    UpDown = sourceRow[6].ToString().Trim(),        //漲跌
                    //漲跌幅公式 漲跌/參考價
                    UpDownPercent = (Math.Round((dUpDown / dPrice), 4) * 100).ToString(),
                    Trade = sourceRow[8].ToString().Trim(),     //成交量
                    DataDate = sourceRow[13].ToString().Trim(),     //時間
                };

                //加入顏色,對應ProdInfo 資料取得的陣列數
                _Top10ProdInfo.ProdNo += "|" + _listTop10CellColor[0].Color.ToString() + "|" + _listTop10CellColor[0].BgColor.ToString();
                _Top10ProdInfo.ProdName += "|" + _listTop10CellColor[0].Color.ToString() + "|" + _listTop10CellColor[0].BgColor.ToString();
                _Top10ProdInfo.Buy += "|" + _listTop10CellColor[1].Color.ToString() + "|" +  _listTop10CellColor[1].BgColor.ToString();
                _Top10ProdInfo.Sell += "|" + _listTop10CellColor[3].Color.ToString() + "|" + _listTop10CellColor[3].BgColor.ToString();
                _Top10ProdInfo.UpDown += "|" + _listTop10CellColor[6].Color.ToString() + "|" + _listTop10CellColor[6].BgColor.ToString();
                _Top10ProdInfo.UpDownPercent += "|" + _listTop10CellColor[6].Color.ToString() + "|" + _listTop10CellColor[6].BgColor.ToString();        //依漲跌顏色為主
                _Top10ProdInfo.Trade += "|" + _listTop10CellColor[8].Color.ToString() + "|" + _listTop10CellColor[8].BgColor.ToString();
                _Top10ProdInfo.DataDate += "|" + _listTop10CellColor[13].Color.ToString() + "|" + _listTop10CellColor[13].BgColor.ToString();

                _listTop10ProdInfo.Add(_Top10ProdInfo);
            }

            return _listTop10ProdInfo;
        }
        //讀檔建立保證金網頁資料 To List
        protected static List<ProdInfo> CreateMarginingListForTxt(string Prod)
        {
            string FilePath = ConfigurationManager.AppSettings["MarginPath"].ToString();
            string FileName = ConfigurationManager.AppSettings["MarginTxtName"].ToString();
            List<ProdInfo> _listProdData = new List<ProdInfo>();
            List<ProdInfo> _listProdInfo = new List<ProdInfo>();
            ProdInfo _ProdInfo = null;
            _listProdData = ReadFile(FilePath + FileName);
            if (_listProdData != null)
            {
                if (Prod != "")
                {
                    foreach (ProdInfo p in _listProdData)
                    {
                        //若輸入是代號,英文則須完全符合,若為名稱則是否存在名稱資料中
                        //數字,英文全符合
                        if (p.ProdId.ToUpper() == Prod.ToUpper() && p.ProdId.Length == Prod.Length)
                        {
                            _ProdInfo = new ProdInfo();
                            _ProdInfo = p;
                        }
                        else if (p.ProdNo == Prod && p.ProdNo.Length == Prod.Length)
                        {
                            _ProdInfo = new ProdInfo();
                            _ProdInfo = p;
                        }
                        else if (p.ProdName.IndexOf(Prod) > -1)
                        {
                            _ProdInfo = new ProdInfo();
                            _ProdInfo = p;
                        }
                        if (_ProdInfo != null)
                        {
                            _listProdInfo.Add(_ProdInfo);
                            _ProdInfo = null;
                        }
                    }

                }
                else
                {
                    _listProdInfo = _listProdData;
                }
            }
            else
            {
                _listProdInfo = null;
            }
            return _listProdInfo;
        }
        //連網建立保證金網頁資料 To List
        protected static List<ProdInfo> CreateMarginingList(string Prod)
        {
            string htmlPage = "";
            List<ProdInfo> _listProdInfo = new List<ProdInfo>();
            ProdInfo _ProdInfo = null;

            HtmlDocument doc = new HtmlDocument();
            HtmlDocument docTable = null;
            HtmlDocument docTR = null;

            HtmlNodeCollection divNode = null;
            HtmlNodeCollection tableNode = null;
            HtmlNodeCollection trNode = null;
            int intTable = 0;

            string[] soRow = null;      //針對特別表格原始資料列用

            //get 台灣期交所 html Page string
            using (WebClient webClient = new WebClient())
            {
                webClient.Encoding = Encoding.UTF8;
                htmlPage = webClient.DownloadString("http://www.taifex.com.tw/chinese/5/stockmargining.asp");
            }

            //台灣期交所,保證金-股票 架構  div[class=section] 
            doc.LoadHtml(htmlPage);
            divNode = doc.DocumentNode.SelectNodes("//div[@class='section']");
            //台灣期交所,保證金-股票 分類為三大div 0. html tag 相關訊息 1.股票期貨契約保證金 2.股票選擇權契約保證金     
            docTable = new HtmlDocument();
            docTable.LoadHtml(divNode[1].InnerHtml);
            tableNode = docTable.DocumentNode.SelectNodes("//table");
            foreach (HtmlNode table in tableNode)
            {
                intTable += 1;
                docTR = new HtmlDocument();
                docTR.LoadHtml(table.InnerHtml.ToString().Replace("\r\n", ""));
                trNode = docTR.DocumentNode.SelectNodes("//tr");
                //建立資料
                //第一行為 tr/th header 部份
                IEnumerable<string[]> rows = trNode.Skip(1).Select(rTR => rTR
                     .Elements("td")
                     .Select(td => td.InnerText.Trim())
                     .ToArray());
                for (int i = 0; i < rows.Count(); i++)
                {
                    if (Prod == "")
                    {
                        soRow = rows.ToList()[i];
                    }
                    else
                    {
                        soRow = null;
                        int xIndex = 0;
                        string sContact = "";
                        //查詢時只取一筆
                        //取得查詢商品ROW
                        foreach (string x in rows.ToList()[i])
                        {
                            //若輸入是代號,英文則須完全符合,若為名稱則是否存在名稱資料中
                            //數字,英文全符合
                            //限定位置,以利效能
                            sContact = (xIndex == 1 || xIndex == 2 || xIndex == 3) ? x : "";
                            if (sContact != "")
                            {
                                if ((x.ToUpper() == Prod.ToUpper() || x == Prod) && x.Length == Prod.Length) { soRow = rows.ToList()[i]; }
                                //中文
                                if (x.IndexOf(Prod) > -1) { soRow = rows.ToList()[i]; }
                                if (soRow != null)
                                {
                                    xIndex = 0;
                                    break;
                                }
                            }
                            xIndex += 1;
                        }
                    }

                    if (soRow != null)
                    {
                        _ProdInfo = new ProdInfo
                        {
                            ProdId = soRow[1].ToString(),
                            ProdNo = soRow[2].ToString(),
                            ProdName = soRow[3].ToString(),
                            Margin = (intTable == 1) ? soRow[8].ToString() : soRow[7].ToString(),
                            QueryDate = DateTime.Now.ToString("HH:mm:ss")
                        };

                        if (intTable == 1)
                        {
                            if (_ProdInfo.ProdName.IndexOf("小型") > -1)
                            {
                                _ProdInfo.ProdType = "1";
                                _ProdInfo.Qty = "100";
                            }
                            else
                            {
                                _ProdInfo.ProdType = "0";
                                _ProdInfo.Qty = "2000";
                            }
                        }
                        else
                        {
                            _ProdInfo.ProdType = "2";
                            _ProdInfo.Qty = "10000";
                        }

                        _listProdInfo.Add(_ProdInfo);
                    }
                }
            }

            return _listProdInfo;
        }
        //建立查詢資料
        protected static List<QueryData> CreateQueryData(string Prod)
        {
            List<QueryData> _listQueryData = new List<QueryData>();
            QueryData _QueryData = null;

            //保證金資料
            List<ProdInfo> _listProdInfo = new List<ProdInfo>();
            //連網取得
            _listProdInfo = CreateMarginingList(Prod);

            //查保證金資料
            if (_listProdInfo != null && _listProdInfo.Count > 0)
            {
                foreach (ProdInfo p in _listProdInfo)
                {
                    _QueryData = new QueryData();
                    _QueryData.ProdId = p.ProdId;
                    _QueryData.ProdName = p.ProdName;
                    _QueryData.ProdNo = p.ProdNo;
                    _QueryData.ProdType = p.ProdType;
                    _QueryData.Qty = p.Qty;
                    _QueryData.QueryDate = DateTime.Now.ToString("HH:mm:ss");
                    //判斷ProType
                    switch (p.ProdType)
                    {
                        case "0":
                        case "1":
                            _QueryData.Margin = p.Margin;       //原始保證金比例
                            _QueryData.TradeMargin = "";        //近月成交價*股數*原始保證金比例
                            break;
                        case "2":
                            _QueryData.Margin = "";
                            _QueryData.TradeMargin = p.Margin;      //原始保證金
                            break;
                    }

                    _listQueryData.Add(_QueryData);
                }
            }
            //成交價查詢
            _listQueryData = CreateTradeList(_listQueryData);
            return _listQueryData;
        }
        //建立股票期貨成交價 To List
        protected static List<QueryData> CreateTradeList(List<QueryData> ListQueryData)
        {
            string htmlUrl = "";
            string htmlPage = "";
            string queryValue = "";
            string strNowYear = "";
            string strNowMonth = "";

            List<QueryData> _listQueryData = new List<QueryData>();

            HtmlDocument doc = new HtmlDocument();
            HtmlDocument docTR = null;

            HtmlNodeCollection tableNode = null;
            HtmlNodeCollection trNode = null;

            _listQueryData = ListQueryData;

            foreach (QueryData q in _listQueryData)
            {
                //取得URL
                if (q.ProdType == "2")
                {
                    htmlUrl = "http://info512.taifex.com.tw/Future/FusaQuote_Norl.aspx?_Commodity=" + q.ProdId + "&_Category=6";
                }
                else
                {
                    htmlUrl = "http://info512.taifex.com.tw/Future/FusaQuote_Norl.aspx?_Commodity=" + q.ProdId + "&_Category=2";
                }
                //近月查詢,依目前台灣期貨交易所,年月格式 為月年(1碼) 註:若2020 會如何表示?
                //若為當月16號之後,則月份+1(期交所規則,到期月好像是16號後當月資料就消失了)
                strNowYear = DateTime.Now.ToString("yyyy");
                if (int.Parse(DateTime.Now.ToString("dd")) > 16)
                {
                    strNowMonth = DateTime.Now.AddMonths(1).ToString("MM");
                }
                else
                {
                    strNowMonth = DateTime.Now.ToString("MM");
                }
                queryValue = strNowMonth + (strNowYear.Substring(strNowYear.Length - 1, 1));
                queryValue = q.ProdName + queryValue;

                //get 台灣期交所 即時交易頁
                using (WebClient webClient = new WebClient())
                {
                    webClient.Encoding = Encoding.UTF8;
                    htmlPage = webClient.DownloadString(htmlUrl);
                }
                doc.LoadHtml(htmlPage);
                tableNode = doc.DocumentNode.SelectNodes("//table[@class='custDataGrid']");
                foreach (HtmlNode table in tableNode)
                {
                    docTR = new HtmlDocument();
                    docTR.LoadHtml(table.InnerHtml.ToString().Replace("\r\n", ""));
                    trNode = docTR.DocumentNode.SelectNodes("//tr");

                    //建立資料列
                    //=======================================
                    IEnumerable<string[]> rows = trNode.Skip(1).Select(rTR => rTR
                         .Elements("td")
                         .Select(td => td.InnerText.Trim())
                         .ToArray());
                    foreach (string[] row in rows)
                    {
                        if (row[0] == queryValue)
                        {
                            q.MonthTrade = row[6].ToString();
                            break;
                        }
                    }
                }

                //計算Type 0,1 交易需要保證金
                double dMonthTrade = 0;
                double dMargin = 0;
                double dQty = 0;

                if (q.ProdType != "2")
                {
                    if (q.MonthTrade != null) { double.TryParse(q.MonthTrade, out dMonthTrade); }
                    if (q.Margin != null) { double.TryParse(q.Margin.Replace("%", ""), out dMargin); }
                    if (q.Qty != null) { double.TryParse(q.Qty, out dQty); }

                    q.TradeMargin = (Math.Round((dMonthTrade * dQty * (dMargin / 100)), 0)).ToString("N0");
                }
            }
            return _listQueryData;
        }
        //讀取Top1-檔案
        protected static string ReadTopFile(string filePath)
        {
            string data = "";

            byte[] buffer;
            FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 1024, true);
            buffer = new byte[(int)fileStream.Length];            // create buffer

            FileData _fileData = new FileData();
            _fileData.Stream = fileStream;
            _fileData.Length = (int)fileStream.Length;
            _fileData.ByteData = buffer;

            IAsyncResult result = fileStream.BeginRead(buffer, 0, _fileData.Length, new AsyncCallback(CallbackRead), _fileData);

            if (result != null)
            {
                _fileData = new FileData();
                _fileData = (FileData)result.AsyncState;
                data = Encoding.Default.GetString(_fileData.ByteData, 0, _fileData.Length);
            }

            return data;
        }
        //讀取檔案
        protected static List<ProdInfo> ReadFile(string filePath)
        {
            string data = "";
            string[] ProdInfoData = { };
            List<ProdInfo> _listProdInfo = new List<ProdInfo>();

            byte[] buffer;
            FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 1024, true);
            buffer = new byte[(int)fileStream.Length];            // create buffer

            FileData _fileData = new FileData();
            _fileData.Stream = fileStream;
            _fileData.Length = (int)fileStream.Length;
            _fileData.ByteData = buffer;

            IAsyncResult result = fileStream.BeginRead(buffer, 0, _fileData.Length, new AsyncCallback(CallbackRead), _fileData);
            if (result != null)
            {
                _fileData = new FileData();
                _fileData = (FileData)result.AsyncState;
                data = Encoding.UTF8.GetString(_fileData.ByteData, 0, _fileData.Length);

                //建置LIST
                if (data != "" && data.Length > 0)
                {
                    ProdInfoData = data.Replace("\r\n", "|").Split('|');
                    //ProdInfoData : [0]ProdType [1]ProdId [2]ProdNo [3]ProdName [4]Qty
                    for (int i = 0; i < ProdInfoData.Length; i++)
                    {
                        string[] pData = ProdInfoData[i].Split(',');
                        ProdInfo _ProdInfo = new ProdInfo
                        {
                            ProdType = pData[0].ToString(),
                            ProdId = pData[1].ToString(),
                            ProdNo = pData[2].ToString(),
                            ProdName = pData[3].ToString(),
                            Qty = pData[4].ToString()
                        };
                        _listProdInfo.Add(_ProdInfo);
                    }
                }
            }
            else
            {
                _listProdInfo = null;
            }
            return _listProdInfo;
        }
        //讀取檔案Use Callback 多線程讀取,防止檔案同時開啟會io.exception
        protected static void CallbackRead(IAsyncResult result)
        {
            FileData fileData = (FileData)result.AsyncState;
            int length = fileData.Stream.EndRead(result);
            fileData.Stream.Close();
            if (length != fileData.Length)
            {
                throw new Exception("Stream is not complete!");
            }
        }
    }
}