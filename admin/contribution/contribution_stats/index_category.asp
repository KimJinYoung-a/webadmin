<%@Language="VBScript" CODEPAGE="65001" %>
<% option explicit %>
<%
Response.CharSet="utf-8" 
Response.codepage="65001"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/lib/classes/contribution/contributionCls.asp"--> 
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<% 
    Dim vcate, syear, smonth, eyear, emonth, i
    vcate = request("vcate")
   	syear     = requestcheckvar(request("sY"),4)
	smonth     = requestcheckvar(request("sM"),2)
   	eyear     = requestcheckvar(request("eY"),4)
	emonth     = requestcheckvar(request("eM"),2)

    if syear ="" then  syear = Cstr(Year( dateadd("m",-12,date()) ))
    if smonth ="" then smonth = Cstr(Month( dateadd("m",-12,date()) ))
    If day(now()) >= 17 Then
        if eyear ="" then  eyear = Cstr(Year( dateadd("m",-1,date()) ))
        if emonth ="" then emonth = Cstr(Month( dateadd("m",-1,date()) ))
    Else
        if eyear ="" then  eyear = Cstr(Year( dateadd("m",-2,date()) ))
        if emonth ="" then emonth = Cstr(Month( dateadd("m",-2,date()) ))            
    End If
%>
<link rel="stylesheet" href="/css/reset.css">
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko_utf8.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="https://code.jquery.com/jquery-3.5.1.min.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>
<style>

</style>
<script type="text/javascript">


    $(function () {
        <% If vcate <> "" Then %>
            GetContributionList();
        <% End If %>
    });

    function SearchForm(frm) {
        frm.submit();
    }

    function GetContributionList() {
        $.ajax({
            url: '/admin/contribution/contribution_stats/jsondata_category.asp?mode=totalcontribution&vcate=<%=vcate%>&sY=<%=syear%>&sM=<%=smonth%>&eY=<%=eyear%>&eM=<%=emonth%>',
            type: 'get',
            data: $('#contributionSearchFrm').serialize(),
            success: function (data) {
                var totalCount = 0;

                if (data == '') {
                    $("#subList").empty().html('<tr align="center" bgcolor="#fff"><td colspan="15">조건에 맞는 데이터가 없습니다.</td></tr>');
                } else {
                    var gnbHtml = "";
                    var it;
                    data = JSON.parse(data);
                    $.each(data, function (index, element) {
						for (it=0;it < element.length;it++)
						{
                            gnbHtml += "<tr align='center' bgcolor='#F7FFE6'>";                        
                            //gnbHtml += "<tr align='center' bgcolor='#FFEDED'>";                            
                            gnbHtml += "<td>" + element[it].YYYYMM + "</td>";
                            gnbHtml += "<td>" + element[it].totalPurchase + "</td>";
                            gnbHtml += "<td>" + element[it].totalPurchaseIncome + "</td>";
                            gnbHtml += "<td>" + element[it].bonusCoupon + "</td>";
                            gnbHtml += "<td>" + element[it].handllingAmount + "</td>";
                            gnbHtml += "<td>" + element[it].handllingAmountIncome + "</td>";
                            gnbHtml += "<td>" + element[it].productQuantity + "</td>";
                            gnbHtml += "<td>" + element[it].numberOfOrders + "</td>";
                            gnbHtml += "<td>" + element[it].variableCost1 + "</td>";
                            gnbHtml += "<td>" + element[it].variableCost2 + "</td>";
                            gnbHtml += "<td>" + element[it].contributionProfit1 + "</td>";
                            gnbHtml += "<td>" + element[it].contributionProfit2 + "</td>";
                            gnbHtml += "<td>" + element[it].totalPurchaseRate + "</td>";
                            gnbHtml += "<td>" + element[it].bonusCouponRate + "</td>";
                            gnbHtml += "<td>" + element[it].handllingAmountRate + "</td>";
                            gnbHtml += "<td>" + element[it].variableCostRate + "</td>";
                            gnbHtml += "<td>" + element[it].contributionProfitRate + "</td>";
                            gnbHtml += "</tr>";                            
						}
                    });
                    $("#subList").empty().html(gnbHtml);
                }
            }
        });
    }
</script>
</head>
<body>
<div class="cont">
    <div class="card-body">
        <h4 class="card-titile" style="padding:10px;border-bottom:1px solid #e4e4e4;">검색 조건</h4>
        <form name="frm" method="get" action="" class="mb-0">
            <div  style="width:280px; display:inline-block;">
                <div  >
                    <span style="font-size:12px;padding:10px;"> 날짜 : </span>
                    <select name="sY">
                        <%for i=year(now()) to 2020 step -1%>
                            <option value="<%=i%>" <%if Cint(sYear) = cint(i) then%>selected<%end if%>><%=i%></option>
                        <%next%>
                    </select>
                    <select name="sM">
                        <%for i=1 to 12%>
                            <option value="<%=i%>" <%if cInt(sMonth) = cInt(i) then%>selected<%end if%>><%=i%></option>
                        <%next%>
                    </select>
                    ~
                    <select name="eY">
                        <%for i=year(now()) to 2020 step -1%>
                            <option value="<%=i%>" <%if Cint(eYear) = cint(i) then%>selected<%end if%>><%=i%></option>
                        <%next%>
                    </select>
                    <select name="eM">
                        <%for i=1 to 12%>
                            <option value="<%=i%>" <%if cInt(eMonth) = cInt(i) then%>selected<%end if%>><%=i%></option>
                        <%next%>
                    </select>
                </div> 
            </div>
            <div style="display:inline-block;width:150px;">
                <select name="vcate" id="vcate" class="select">
                    <option value="101" <% If vcate="101" Then %>selected<% End If %>>디자인문구</option>
                    <option value="102" <% If vcate="102" Then %>selected<% End If %>>디지털/핸드폰</option>
                    <option value="104" <% If vcate="104" Then %>selected<% End If %>>토이/취미</option>
                    <option value="124" <% If vcate="124" Then %>selected<% End If %>>디자인가전</option>
                    <option value="121" <% If vcate="121" Then %>selected<% End If %>>가구/수납</option>
                    <option value="122" <% If vcate="122" Then %>selected<% End If %>>데코/조명</option>
                    <option value="120" <% If vcate="120" Then %>selected<% End If %>>패브릭/생활</option>
                    <option value="112" <% If vcate="112" Then %>selected<% End If %>>키친</option>
                    <option value="119" <% If vcate="119" Then %>selected<% End If %>>푸드</option>
                    <option value="117" <% If vcate="117" Then %>selected<% End If %>>패션의류</option>
                    <option value="116" <% If vcate="116" Then %>selected<% End If %>>패션잡화</option>
                    <option value="125" <% If vcate="125" Then %>selected<% End If %>>주얼리/시계</option>
                    <option value="118" <% If vcate="118" Then %>selected<% End If %>>뷰티</option>
                    <option value="115" <% If vcate="115" Then %>selected<% End If %>>베이비/키즈</option>
                    <option value="110" <% If vcate="110" Then %>selected<% End If %>>Cat & Dog</option>
                </select>
            </div>
            <div style="display:inline-block;">
                <button type="button"  class="button_s" style="width:100px;height:30px;" onclick="SearchForm(this.form);">검색</button>
            </div>
        </form>
    </div>
    <div class="pad20">
        <div class="list02">
            <table cellpadding="0" cellspacing="0" border="0" class="a" width="100%">
                <tr height="50">
                </tr>
            </table>
        </div>
    </div>    
    <% If vcate <> "" Then %>
        <div class="pad20">
            <div class="list02">
                <form name="frmList" method="POST" action="">
                    <input type="hidden" name="mode" value="sub">
                    <table align="left" cellpadding="5" cellspacing="1" class="a" bgcolor="#999" width="100%">
                        <!--tbody>
                            <tr bgcolor="#FFF">
                                <td colspan="15">
                                    <table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
                                        <tbody>
                                            <tr>
                                                <td align="left">총 <span id="gnbTotalCount">0</span> 건</td>
                                                <td align="right">
                                                    <input type="button" value="정렬수정" onclick="goSortArrayEdit();" class="button" />&nbsp;
                                                    <input type="button" value="GNB 등록" onclick="goGnbWriteEdit('');" class="button">
                                                </td>
                                            </tr>
                                        </tbody>
                                    </table>
                                </td>
                            </tr>
                        </tbody-->
                        <colgroup>
                            <col width="80">
                            <col width="80">
                            <col width="80">
                            <col width="80">
                            <col width="80">
                            <col width="80">
                            <col width="80">
                            <col width="80">
                            <col width="80">
                            <col width="80">
                            <col width="80">
                            <col width="80">
                            <col width="80">
                            <col width="80">
                            <col width="80">
                            <col width="80">
                            <col width="80">
                        </colgroup>
                        <tbody>
                            <tr align="center" style="background:#E6E6E6;">
                                <td>년월</td>
                                <td>구매총액</td>
                                <td>구매총액수익</td>
                                <td>보너스쿠폰</td>
                                <td>취급액</td>
                                <td>취급액수익</td>
                                <td>상품수량</td>
                                <td>주문건수</td>
                                <td>변동비1<br/>(물류,수수료)</td>
                                <td>변동비2<br/>(판촉비)</td>
                                <td>공헌이익1</td>
                                <td>공헌이익2</td>
                                <td>구매총액수익율</td>
                                <td>보너스쿠폰율</td>
                                <td>취급액수익율</td>
                                <td>변동비율</td>
                                <td>공헌이익율</td>
                            </tr>
                        </tbody>
                        <tbody id="subList">
                        </tbody>
                    </table>
                    <input name="__RequestVerificationToken" type="hidden" value="CfDJ8O2TkOoLyxFPmywhWCbibd8CJsp-OXurdzFgId9unP_ZRfcTctHu9Dwz48pAfXWUsWb1xY0l-Cs9H5vx1Y9_p2jfNwt5XxGo1x3KFSjW_0tUGEy54ITYH6a2IhC-p1uegc8awQLSW7Xh8YJDFMN1Zv869aNPLIhXHKktFFWNafK_MNSXHNqe1M3-PvcSAMGrWA" />
                </form>
            </div>
        </div>
        <div class="pad20">
            <div class="list02">
                <table cellpadding="0" cellspacing="0" border="0" class="a" width="100%">
                    <tr height="100">
                    </tr>
                </table>
            </div>
        </div>
        <div class="pad20">
            <div class="list02">
                <table cellpadding="0" cellspacing="0" border="0" class="a" width="100%">
                    <tr>
                        <td>
                            <!--table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
                                <tr>
                                    <td width="100%" align="center" style="padding:7 0 7 0"><font size="3">공헌이익율구조</font></td>
                                </tr>
                            </table-->
                            <div id="chartdiv2" align="center"></div>
                            <script type="text/javascript">	
                            $('document').ready(function(){
                                $.ajax({
                                    url: '/admin/contribution/contribution_stats/jsondata_category.asp?mode=ContributionMarginStructure&vcate=<%=vcate%>&sY=<%=syear%>&sM=<%=smonth%>&eY=<%=eyear%>&eM=<%=emonth%>',
                                    type: 'get',
                                    data: $('#contributionSearchFrm').serialize(),
                                    success: function (data) {
                                        data = JSON.parse(data);
                                        FusionCharts.ready(function(){
                                            var myChart2 = new FusionCharts({
                                                "type": "msline",
                                                "width":"950",
                                                "height":"550",
                                                "dataFormat": "json"
                                            });
                                            myChart2.setJSONData(data);
                                            myChart2.render("chartdiv2");
                                        });
                                    }
                                });
                            });
                            </script>
                        </td>                
                    </tr>
                </table>
            </div>
        </div>
        <div class="pad20">
            <div class="list02">
                <table cellpadding="0" cellspacing="0" border="0" class="a" width="100%">
                    <tr height="100">
                    </tr>
                </table>
            </div>
        </div>
        <div class="pad20">
            <div class="list02">
                <table cellpadding="0" cellspacing="0" border="0" class="a" width="100%">
                    <tr>
                        <td>
                            <!--table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
                                <tr>
                                    <td width="100%" align="center" style="padding:7 0 7 0"><font size="3">구매총액</font></td>
                                </tr>
                            </table-->
                            <div id="chartdiv3" align="center"></div>
                            <script type="text/javascript">	
                            $('document').ready(function(){
                                $.ajax({
                                    url: '/admin/contribution/contribution_stats/jsondata_category.asp?mode=TotalPurchase&vcate=<%=vcate%>&sY=<%=syear%>&sM=<%=smonth%>&eY=<%=eyear%>&eM=<%=emonth%>',
                                    type: 'get',
                                    data: $('#contributionSearchFrm').serialize(),
                                    success: function (data) {
                                        data = JSON.parse(data);
                                        FusionCharts.ready(function(){
                                            var myChart3 = new FusionCharts({
                                                "type": "msline",
                                                "width":"950",
                                                "height":"550",
                                                "dataFormat": "json"
                                            });
                                            myChart3.setJSONData(data);
                                            myChart3.render("chartdiv3");
                                        });
                                    }
                                });
                            });
                            </script>
                        </td>                
                    </tr>
                </table>
            </div>
        </div>    
        <div class="pad20">
            <div class="list02">
                <table cellpadding="0" cellspacing="0" border="0" class="a" width="100%">
                    <tr height="100">
                    </tr>
                </table>
            </div>
        </div>
        <div class="pad20">
            <div class="list02">
                <table cellpadding="0" cellspacing="0" border="0" class="a" width="100%">
                    <tr>
                        <td>
                            <!--table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
                                <tr>
                                    <td width="100%" align="center" style="padding:7 0 7 0"><font size="3">취급액</font></td>
                                </tr>
                            </table-->
                            <div id="chartdiv4" align="center"></div>
                            <script type="text/javascript">	
                            $('document').ready(function(){
                                $.ajax({
                                    url: '/admin/contribution/contribution_stats/jsondata_category.asp?mode=HandlingAmount&vcate=<%=vcate%>&sY=<%=syear%>&sM=<%=smonth%>&eY=<%=eyear%>&eM=<%=emonth%>',
                                    type: 'get',
                                    data: $('#contributionSearchFrm').serialize(),
                                    success: function (data) {
                                        data = JSON.parse(data);
                                        FusionCharts.ready(function(){
                                            var myChart4 = new FusionCharts({
                                                "type": "msline",
                                                "width":"950",
                                                "height":"550",
                                                "dataFormat": "json"
                                            });
                                            myChart4.setJSONData(data);
                                            myChart4.render("chartdiv4");
                                        });
                                    }
                                });
                            });
                            </script>
                        </td>                
                    </tr>
                </table>
            </div>
        </div>
        <div class="pad20">
            <div class="list02">
                <table cellpadding="0" cellspacing="0" border="0" class="a" width="100%">
                    <tr height="100">
                    </tr>
                </table>
            </div>
        </div>
        <div class="pad20">
            <div class="list02">
                <table cellpadding="0" cellspacing="0" border="0" class="a" width="100%">
                    <tr>
                        <td>
                            <!--table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
                                <tr>
                                    <td width="100%" align="center" style="padding:7 0 7 0"><font size="3">변동비</font></td>
                                </tr>
                            </table-->
                            <div id="chartdiv5" align="center"></div>
                            <script type="text/javascript">	
                            $('document').ready(function(){
                                $.ajax({
                                    url: '/admin/contribution/contribution_stats/jsondata_category.asp?mode=VariableCost&vcate=<%=vcate%>&sY=<%=syear%>&sM=<%=smonth%>&eY=<%=eyear%>&eM=<%=emonth%>',
                                    type: 'get',
                                    data: $('#contributionSearchFrm').serialize(),
                                    success: function (data) {
                                        data = JSON.parse(data);
                                        FusionCharts.ready(function(){
                                            var myChart5 = new FusionCharts({
                                                "type": "msline",
                                                "width":"950",
                                                "height":"550",
                                                "dataFormat": "json"
                                            });
                                            myChart5.setJSONData(data);
                                            myChart5.render("chartdiv5");
                                        });
                                    }
                                });
                            });
                            </script>
                        </td>                
                    </tr>
                </table>
            </div>
        </div>
        <div class="pad20">
            <div class="list02">
                <table cellpadding="0" cellspacing="0" border="0" class="a" width="100%">
                    <tr height="100">
                    </tr>
                </table>
            </div>
        </div>
        <div class="pad20">
            <div class="list02">
                <table cellpadding="0" cellspacing="0" border="0" class="a" width="100%">
                    <tr>
                        <td>
                            <!--table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
                                <tr>
                                    <td width="100%" align="center" style="padding:7 0 7 0"><font size="3">공헌이익</font></td>
                                </tr>
                            </table-->
                            <div id="chartdiv6" align="center"></div>
                            <script type="text/javascript">	
                            $('document').ready(function(){
                                $.ajax({
                                    url: '/admin/contribution/contribution_stats/jsondata_category.asp?mode=ContributionProfit&vcate=<%=vcate%>&sY=<%=syear%>&sM=<%=smonth%>&eY=<%=eyear%>&eM=<%=emonth%>',
                                    type: 'get',
                                    data: $('#contributionSearchFrm').serialize(),
                                    success: function (data) {
                                        data = JSON.parse(data);
                                        FusionCharts.ready(function(){
                                            var myChart6 = new FusionCharts({
                                                "type": "msline",
                                                "width":"950",
                                                "height":"550",
                                                "dataFormat": "json"
                                            });
                                            myChart6.setJSONData(data);
                                            myChart6.render("chartdiv6");
                                        });
                                    }
                                });
                            });
                            </script>
                        </td>                
                    </tr>
                </table>
            </div>
        </div>
    <% Else %>
        <div class="pad20">
            <div class="list02">
                <table cellpadding="0" cellspacing="0" border="0" class="a" width="100%">
                    <tr height="300">
                        <td>
                        </td>                
                    </tr>
                    <tr>
                        <td>
                            <table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
                                <tr>
                                    <td width="100%" align="center" style="padding:7 0 7 0"><font size="3">카테고리를 선택해주세요.</font></td>
                                </tr>
                            </table>
                        </td>                
                    </tr>
                </table>
            </div>
        </div>
    <% End If %>
</div>
<form name="contributionSearchFrm" id="contributionSearchFrm">
</form>
</body>
</html>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->