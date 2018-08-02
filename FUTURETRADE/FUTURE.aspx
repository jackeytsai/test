<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="FUTURE.aspx.cs" Inherits="FUTURETRADE.FUTURE" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link rel="Stylesheet" type="text/css" href="CSS/MainStyle.css" />
    <script type="text/javascript" src="JS/jquery-1.12.4.js"></script>
    <title></title>
    <style type="text/css">
        .container {
            width: 100%;
            height: 24px;
            border: solid 0px #000;
            position: relative;
        }

        .bottom-right {
            right: 0;
            bottom: 0;
            position: absolute;
        }

        .bottom-left {
            left: 0;
            bottom: 0;
            position: absolute;
        }
    </style>
    <script language="javascript" type="text/javascript">

        var jsonProdInfo = null;
        var t = 0;
        var startTime = null;
        var endTime = null;
        var preEndTime = null;

        $(document).ready(function () {
            //設定更新時間
            t = $("#hdTimer").val();
            $("#TimerSet").html(t);

            preEndTime = new Date().getDate();

            //取回Margin保證金資料
            QueryMargin();
            //等待Margin取回資料,1秒後執行TOP10
            setTimeout('QueryTop10()', 999);
            //每秒執行一次,showTime()
            setInterval("showTime()", 1000);

            //查詢
            $("#btnQuery").click(function () {
                //清空
                $("#spanProdName").html("");
                $('#tableQuery').find('tbody').empty();

                var tParam = fh_GenParam();
                var Prod = $.trim($("#txtProd").val());

                if (Prod == "") {
                    $("#spanQueryErrMsg").html("商品代號/名稱未輸入");
                    //alert("商品代號/名稱未輸入");
                    return false;
                }
                //取回資料
                QueryData(Prod);
            })
        })
        //顯示倒數秒收
        function showTime() {
            t -= 1;
            $("#cTimerDown").html(t);

            if (t == 0) {
                QueryTop10();
            }
        }
        //取回Margin保證金資料
        function QueryMargin() {
            $("#spanTop10ErrMsg").html("");
            var tParam = fh_GenParam();
            var m_WhereParm = { Prod: "" };
            var m_objJSON = fh_CallWebMethod("FUTURE.aspx", "QueryMargin?uri=" + tParam, m_WhereParm, false);
            //錯誤顯示
            if (m_objJSON.ErrMsg != "") {
                $("#spanTop10ErrMsg").html(m_objJSON.ErrMsg);
                return false;
            } else {
                if (m_objJSON.ListProdInfo != null) {
                    jsonProdInfo = m_objJSON.ProdInfoJson;
                } else {
                    $("#spanTop10ErrMsg").html("查無保證金資料");
                }
            }
        }
        //取回Top10資料
        function QueryTop10() {
            $("#spanTop10ErrMsg").html("");

            var tParam = fh_GenParam();
            //若無保證金資料,則重取
            if (jsonProdInfo == null) { QueryMargin(); }
            var m_WhereParm = { ProdInfo: jsonProdInfo };
            var m_objJSON = fh_CallWebMethod("FUTURE.aspx", "QueryTop10ProdInfo?uri=" + tParam, m_WhereParm, false, "");
            //錯誤顯示
            if (m_objJSON.ErrMsg != "") {
                $("#spanTop10ErrMsg").html(m_objJSON.ErrMsg);
                t = $("#hdTimer").val();
                return false;
            } else {
                if (m_objJSON.ListTop10ProdInfo != null && m_objJSON.ListTop10ProdInfo.length > 0) {
                    //建置頁面
                    try {
                        Render_Top10(m_objJSON.ListTop10ProdInfo);
                        //重置更新時間
                        t = $("#hdTimer").val();
                        if (m_objJSON.Top10ErrMsg != "") {
                            $("#spanTop10ErrMsg").html(m_objJSON.Top10ErrMsg);
                        }
                    } catch (mExc) {
                        //重置更新時間
                        t = $("#hdTimer").val();
                        $("#spanTop10ErrMsg").html(mExc.description);
                    }
                } else {
                    //重置畫面
                     $('#tableTop10').find('tbody').empty();
                    //重置更新時間
                    t = $("#hdTimer").val();
                    if (m_objJSON.Top10ErrMsg != "") {
                        $("#spanTop10ErrMsg").html(m_objJSON.Top10ErrMsg);
                    } else {
                        $("#spanTop10ErrMsg").html("目前尚無資料!");
                    }
                }
            }
        }
        //建置Top10HTML 資料列
        function Render_Top10(QueryData) {
            var _Blood = "";
            var _index = 1;
            var strCss = "";

            _blood = '<tbody>';

            $.each(QueryData, function (date, item) {
                if ((_index % 2) == 0) {
                    _blood += '<tr class=\' custDataGridRow\'  bgcolor=\' #DADBF7\'>';
                } else {
                    _blood += '<tr class=\' custDataGridRow\'  bgcolor=\' White\'>';
                }
                //股票代號
                if (item.ProdNo == undefined || item.ProdNo == null) { _blood += '<td class="bu13" align="left">&nbsp;</td>'; }
                else if (item.ProdNo.split('|')[0].length < 1) { _blood += '<td class="bu13" align="left">&nbsp;</td>'; }
                else { _blood += '<td class="bu13" align="left"><font color="' + item.ProdNo.split('|')[1] + '">' + item.ProdNo.split('|')[0] + '<font></td>'; }
                //商品名稱 
                if (item.ProdName == undefined || item.ProdName == null) { _blood += '<td class="bu13" align="left">&nbsp;</td>'; }
                else if (item.ProdName.split('|')[0].length < 1) { _blood += '<td class="bu13" align="left">&nbsp;</td>'; }
                else { _blood += '<td class="bu13" align="left"><font color="' + item.ProdName.split('|')[1] + '">' + item.ProdName.split('|')[0] + '<font></td>'; }
                //買價
                if (item.Buy == undefined || item.Buy == null) { _blood += '<td>&nbsp;</td>'; }
                else if (item.Buy.split('|')[0].length < 1) { _blood += '<td>&nbsp;</td>'; }
                else { _blood += '<td bgcolor="' + item.Buy.split('|')[2] + '"><font color="' + item.Buy.split('|')[1] + '">' + item.Buy.split('|')[0] + '<font></td>'; }
                //賣價
                if (item.Sell == undefined || item.Sell == null) { _blood += '<td>&nbsp;</td>'; }
                else if (item.Sell.split('|')[0].length < 1) { _blood += '<td>&nbsp;</td>'; }
                else { _blood += '<td bgcolor="' + item.Sell.split('|')[2] + '"><font color="' + item.Sell.split('|')[1] + '">' + item.Sell.split('|')[0] + '<font></td>'; }
                //漲跌
                if (item.UpDown == undefined || item.UpDown == null) { _blood += '<td>&nbsp;</td>'; }
                else if (item.UpDown.split('|')[0].length < 1) { _blood += '<td>&nbsp;</td>'; }
                else { _blood += '<td bgcolor="' + item.UpDown.split('|')[2] + '"><font color="' + item.UpDown.split('|')[1] + '">' + item.UpDown.split('|')[0] + '<font></td>'; }
                //漲跌幅
                if (item.UpDownPercent == undefined || item.UpDownPercent == null) { _blood += '<td>&nbsp;</td>'; }
                else if (item.UpDownPercent.split('|')[0].length < 1) { _blood += '<td>&nbsp;</td>'; }
                else { _blood += '<td bgcolor="' + item.UpDownPercent.split('|')[2] + '"><font color="' + item.UpDownPercent.split('|')[1] + '">' + item.UpDownPercent.split('|')[0] + '<font></td>'; }
                //成交量
                if (item.Trade == undefined || item.Trade == null) { _blood += '<td>&nbsp;</td>'; }
                else if (item.Trade.split('|')[0].length < 1) { _blood += '<td>&nbsp;</td>'; }
                else { _blood += '<td bgcolor="' + item.Trade.split('|')[2] + '"><font color="' + item.Trade.split('|')[1] + '">' + item.Trade.split('|')[0] + '<font></td>'; }
                //資料時間
                if (item.DataDate == undefined || item.DataDate == null) { _blood += '<td>&nbsp;</td>'; }
                else if (item.DataDate.split('|')[0].length < 1) { _blood += '<td>&nbsp;</td>'; }
                else { _blood += '<td bgcolor="' + item.DataDate.split('|')[2] + '"><font color="' + item.DataDate.split('|')[1] + '">' + item.DataDate.split('|')[0] + '<font></td>'; }

                _blood += "</tr>";
                _index += 1;
            })
            _blood += "</tbody>";
            $('#tableTop10').find('tbody').empty();
            $('#tableTop10').append(_blood);
        }
        //取回查詢資料
        function QueryData(qProd) {
            $("#spanQueryErrMsg").html("");
            var tParam = fh_GenParam();
            var m_WhereParm = { Prod: qProd };
            var m_objJSON = fh_CallWebMethod("FUTURE.aspx", "QueryData?uri=" + tParam, m_WhereParm, false, "query");
            //錯誤顯示
            if (m_objJSON.ErrMsg != "") {
                $("#spanQueryErrMsg").html(m_objJSON.ErrMsg);
                return false;
            } else {
                if (m_objJSON.ListQueryData != null && m_objJSON.ListQueryData.length > 0) {
                    //建置頁面
                    try {
                        Render_Query(m_objJSON.ListQueryData);
                    } catch (mExc) {
                        $("#spanQueryErrMsg").html(mExc.description);
                    }
                } else {
                    $("#spanQueryErrMsg").html("查無此商品!");
                }
            }
        }
        //建置查詢HTML 資料列
        function Render_Query(QueryData) {
            var _Blood = "";
            var _ProdName = "";

            _blood = '<tbody>';

            $.each(QueryData, function (date, item) {
                //商品全名
                //_ProdName += (_ProdName == "") ? item.ProdNo + item.ProdName + "(" + item.ProdId + ")" : "/" + item.ProdNo + item.ProdName + "(" + item.ProdId + ")";

                _blood += '<tr>';
                //商品名稱
                if (item.ProdName == undefined || item.ProdName == null) { _blood += '<td>&nbsp;</td>'; }
                else if (item.ProdName.length < 1) { _blood += '<td>&nbsp;</td>'; }
                else { _blood += '<td align="right">' + item.ProdName + '</td>'; }
                //近月成交價
                if (item.MonthTrade == undefined || item.MonthTrade == null) { _blood += '<td>&nbsp;</td>'; }
                else if (item.MonthTrade.length < 1) { _blood += '<td>&nbsp;</td>'; }
                else { _blood += '<td align="right">' + item.MonthTrade + '</td>'; }
                //股數/口
                if (item.Qty == undefined || item.Qty == null) { _blood += '<td>&nbsp;</td>'; }
                else if (item.Qty.length < 1) { _blood += '<td>&nbsp;</td>'; }
                else { _blood += '<td align="right">' + item.Qty + '</td>'; }
                //原始保證金比例
                if (item.Margin == undefined || item.Margin == null) { _blood += '<td>&nbsp;</td>'; }
                else if (item.Margin.length < 1) { _blood += '<td>&nbsp;</td>'; }
                else { _blood += '<td align="right">' + item.Margin + '</td>'; }
                //交易需要保證金
                if (item.TradeMargin == undefined || item.TradeMargin == null) { _blood += '<td>&nbsp;</td>'; }
                else if (item.TradeMargin.length < 1) { _blood += '<td>&nbsp;</td>'; }
                else { _blood += '<td align="right">' + item.TradeMargin + '</td>'; }
                //時間
                if (item.QueryDate == undefined || item.QueryDate == null) { _blood += '<td>&nbsp;</td>'; }
                else if (item.QueryDate.length < 1) { _blood += '<td>&nbsp;</td>'; }
                else { _blood += '<td align="right">' + item.QueryDate + '</td>'; }

                _blood += "</tr>";
            })
            _blood += "</tbody>";
            //$("#spanProdName").html("");
            //$("#spanProdName").html(_ProdName);

            $('#tableQuery').find('tbody').empty();
            $('#tableQuery').append(_blood);
        }

        function fh_GenParam() {
            var t = "";
            var sdate = new Date();
            t = sdate.getFullYear() + "" + (sdate.getMonth() + 1) + "" + sdate.getDate() + "" + sdate.getHours() + "" + sdate.getMinutes() + "" + sdate.getSeconds() + "" + sdate.getMilliseconds();
            return t;
        }

        function fh_CallWebMethod(m_Source, m_Method, m_WhereParm, m_async, m_query) {
            try {
                var WhereParam = "";
                if (m_WhereParm != "") {
                    for (var Param in m_WhereParm) {
                        if (typeof (m_WhereParm[Param]) == "object") {
                            var pWhereParam = Param + ": [  ";
                            for (var P in m_WhereParm[Param]) {
                                pWhereParam += " '" + m_WhereParm[Param][P] + "' , ";
                            }
                            pWhereParam = pWhereParam.substr(0, pWhereParam.lastIndexOf(","));
                            pWhereParam += "], ";
                            WhereParam += pWhereParam;
                        }
                        else {
                            WhereParam += (WhereParam == "" ? Param + ":'" + m_WhereParm[Param] + "'" : " ," + Param + ":'" + m_WhereParm[Param] + "'");
                        }
                    }
                }

                var ReturnObj = { ErMsg: "" };
                $.ajax({
                    type: "POST",
                    async: m_async,
                    url: m_Source + "/" + m_Method,
                    data: "{" + WhereParam + "}",
                    contentType: "application/json; charset=utf-8",
                    dataType: "json",
                    error: function (xmlHttpRequest, error) {
                        var m_ErrMessage = xmlHttpRequest.responseJSON.Message;
                        //重置更新時間
                        if (m_query == "") { t = $("#hdTimer").val(); }
                        //alert(m_ErrMessage);
                    },
                    success: function (data) {
                        ReturnObj = data.d
                    }
                });
                return ReturnObj;
            }
            catch (e) {
                //重置更新時間
                if (m_query == "") { t = $("#hdTimer").val(); }
                //alert(e.Message);
            }
        }
        function iframeLoaded() {
            var iFrameID = document.getElementById('iFutrue');
            if (iFrameID) {
                // here you can make the height, I delete it first, then I make it again
                iFrameID.height = "";
                iFrameID.height = iFrameID.contentWindow.document.body.scrollHeight + "px";
            }
        }

        function checkUpdateTime() {
            var flag = true;
            var today = new Date();

            startTime = today.getFullYear().toString() + "/" + (today.getMonth() + 1).toString() + "/" + today.getDate().toString() + " " + $("#hdStartTimer").val();
            endTime = today.getFullYear().toString() + "/" + (today.getMonth() + 1).toString() + "/" + today.getDate().toString() + " " + $("#hdEndTimer").val();

            if (Date.parse(today).valueOf() < Date.parse(startTime).valueOf()) {
                flag = false;
            } else {
                if (Date.parse(today).valueOf() > Date.parse(endTime).valueOf()) {
                    flag = false;
                }
            }

            return flag;
        }
    </script>

</head>
<body class="backgroundStyle">
    <div style="text-align: left">
        <font style="font-family: 微軟正黑體; font-size: 21px; font-weight: bold;">股票期貨 TOP10 成交量即時報價</font>
    </div>
    <div class="container">
        <div class="bottom-right">
            隔&nbsp;<span id="TimerSet"></span>&nbsp;秒自動更新,尚餘&nbsp;<span id="cTimerDown"></span>&nbsp;秒更新
        </div>
        <div class="bottom-left">
        </div>
    </div>
    <div id="divDG" style="height: 305px; width: 100%; overflow: auto;">
        <table class="custDataGrid" cellspacing="0" cellpadding="4" border="0" id="tableTop10">
            <thead>
                <tr class="custDataGridRow" bgcolor="#104E8D">
                    <td align="left" nowrap="nowrap"><font color="White"><b><font color="White">股票代號</font></b></font></td>
                    <td align="left" nowrap="nowrap"><font color="White"><b><font color="White">商品</font></b></font></td>
                    <td align="right" nowrap="nowrap"><font color="White"><b><font color="White">買價</font></b></font></td>
                    <td align="right" nowrap="nowrap"><font color="White"><b><font color="White">賣價</font></b></font></td>
                    <td align="right" nowrap="nowrap"><font color="White"><b><font color="White">漲跌</font></b></font></td>
                    <td align="right" nowrap="nowrap"><font color="White"><b><font color="White">漲跌幅%</font></b></font></td>
                    <td align="right" nowrap="nowrap"><font color="White"><b><font color="White">成交量</font></b></font></td>
                    <td align="right" nowrap="nowrap"><font color="White"><b><font color="White">時間</font></b></font></td>
                </tr>
            </thead>
            <tbody></tbody>
        </table>
        <p><span id="spanTop10ErrMsg" style="color: red; font-weight: bold;"></span></p>
    </div>
    <div class="container">
        <div class="bottom-right">
            *此報價僅供參考，正確報價請依照期貨交易所公告為主
        </div>
        <div class="bottom-left">
        </div>
    </div>
    <br />
    <div id="divQuery">
        <p><font style="font-family: 微軟正黑體; font-size: 21px; font-weight: bold;">股票期貨保證金試算</font></p>
        <br />
        <p>
            請輸入股票代號/名稱：<input id="txtProd" name="txtProd" type="text" />&nbsp;&nbsp;
            <input type="button" id="btnQuery" name="btnQuery" value="查詢" style="height: 22px;" />
        </p>
        <p><span id="spanQueryErrMsg" style="color: red; font-weight: bold;"></span></p>
        <table class="tableStyle" cellspacing="0" cellpadding="4" border="0" id="tableQuery">
            <thead>
                <tr bgcolor="#104E8D">
                    <td align="left" nowrap="nowrap"><font color="White"><b><font color="White">商品名稱</font></b></font></td>
                    <td align="right" nowrap="nowrap"><font color="White"><b><font color="White">近月成交價</font></b></font></td>
                    <td align="right" nowrap="nowrap"><font color="White"><b><font color="White">股數/口</font></b></font></td>
                    <td align="right" nowrap="nowrap"><font color="White"><b><font color="White">原始保證金比例</font></b></font></td>
                    <td align="right" nowrap="nowrap"><font color="White"><b><font color="White">交易需要保證金</font></b></font></td>
                    <td align="right" nowrap="nowrap"><font color="White"><b><font color="White">時間</font></b></font></td>
                </tr>
            </thead>
            <tbody></tbody>
        </table>
    </div>
    <form id="form1" runat="server">
        <asp:HiddenField ID="hdTimer" runat="server" />
        <asp:HiddenField ID="hdStartTimer" runat="server" />
        <asp:HiddenField ID="hdEndTimer" runat="server" />
    </form>
</body>
</html>


