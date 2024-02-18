<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �ΰŽ� ��������-�Ϻ�
' History : 2016.09.20 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/report/maechul/statisticCls.asp" -->
<%
Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay
Dim vTot_CountPlus, vTot_CountMinus, vTot_MaechulPlus, vTot_MaechulMinus, vTot_Subtotalprice, vTot_Miletotalprice, vTot_subtotalprice_notexists_sumPaymentEtc
dim vTot_MaechulCountSum, vTot_MaechulPriceSum, vTot_sumPaymentEtc, page, pagesize, vSorting
dim sellchnl
	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= RequestCheckvar(request("sitename"),16)
	vDateGijun	= NullFillWith(RequestCheckvar(request("date_gijun"),16),"regdate")
	'vSYear		= NullFillWith(RequestCheckvar(request("syear"),4),Year(DateAdd("d",-13,now())))
	'vSMonth		= NullFillWith(RequestCheckvar(request("smonth"),2),Month(DateAdd("d",-13,now())))
	'vSDay		= NullFillWith(RequestCheckvar(request("sday"),2),Day(DateAdd("d",-13,now())))
	vSYear		= NullFillWith(RequestCheckvar(request("syear"),4),Year(now()))
	vSMonth		= NullFillWith(RequestCheckvar(request("smonth"),2),Month(now()))
	vSDay		= NullFillWith(RequestCheckvar(request("sday"),2),"01")
	vEYear		= NullFillWith(RequestCheckvar(request("eyear"),4),Year(now))
	vEMonth		= NullFillWith(RequestCheckvar(request("emonth"),2),Month(now))
	vEDay		= NullFillWith(RequestCheckvar(request("eday"),2),Day(now))
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	vSorting	= NullFillWith(RequestCheckvar(request("sorting"),32),"ddateD")

if (page = "") then
	page = 1
end if

if (pagesize = "") then
	pagesize = 3000
end if

Set cStatistic = New cacademyStatic_list
	cStatistic.FCurrPage = page
	cStatistic.FPageSize = pagesize
	cStatistic.FRectSort = vSorting
	cStatistic.FRectDateGijun = vDateGijun
	cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	cStatistic.FRectSellChannelDiv = sellchnl
	cStatistic.facademyStatistic_dailylist()


function getWeekdayStr2(yyyymmdd)
	dim wd
	if IsNULL(yyyymmdd) then Exit function
	wd = weekday(yyyymmdd)

	select case wd
		case 1
			getWeekdayStr2 = "��"
		case 2
			getWeekdayStr2 = "��"
		case 3
			getWeekdayStr2 = "ȭ"
		case 4
			getWeekdayStr2 = "��"
		case 5
			getWeekdayStr2 = "��"
		case 6
			getWeekdayStr2 = "��"
		case 7
			getWeekdayStr2 = "��"
		case else
			getWeekdayStr2 = yyyymmdd
	end select

end function
%>

<script type='text/javascript'>

function downloadexcel(){
    document.frm.target = "view"; 
    document.frm.action = "/academy/report/maechul/statistic_daily_excel.asp";  
	document.frm.submit();
    document.frm.target = ""; 
    document.frm.action = "";  
}

function searchSubmit(){
    frm.submit();
}

function jstrSort(vsorting){
	var tmpSorting = document.getElementById("img"+vsorting)

	if (-1 < tmpSorting.src.indexOf("_alpha")){
		frm.sorting.value= vsorting+"D";
	}else if (-1 < tmpSorting.src.indexOf("_bot")){
		frm.sorting.value= vsorting+"A";
	}else{
		frm.sorting.value= vsorting+"D";
	}
	searchSubmit();
}

</script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/fusioncharts.js"></script>
<script type="text/javascript" src="/lib/util/fusionchartsXT/js/themes/fusioncharts.theme.fint.js"></script>
<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="sorting" value="<%= vsorting %>">

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="30">
				* �Ⱓ :
				<select name="date_gijun" class="select">
					<option value="regdate" <%=CHKIIF(vDateGijun="regdate","selected","")%>>�ֹ���</option>
					<option value="ipkumdate" <%=CHKIIF(vDateGijun="ipkumdate","selected","")%>>������</option>
				</select>
				&nbsp;
				<% DrawDateBoxdynamic vSYear,"syear",vEYear,"eyear",vSMonth,"smonth",vEMonth,"emonth",vSDay,"sday",vEDay,"eday" %>
			</td>
		</tr>
		<tr>
		    <td>
		    	* ����Ʈ���� : <% drawradio_academy_sitename "sitename", vSiteName, "", "Y" %>
			    &nbsp;
            	* ä�α��� : <% drawSelectBox_SellChannel "sellchnl", sellchnl, "" %>
			</td>
		</tr>
	    </table>
	</td>
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="�˻�" onClick="javascript:searchSubmit();"></td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		* �˻� �Ⱓ�� ������� ����� �������ϴ�. �׷��� �˻� ��ư�� Ŭ���� �� �ƹ� ������ ����δٰ� ���� �˻���ư�� Ŭ������ ������.
	</td>
	<td align="right">	
		<input type="button" onclick="downloadexcel();" value="�����ٿ�ε�" class="button">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="25">
		�˻���� : <b><%=cStatistic.FresultCount%></b>&nbsp;&nbsp;�� �ִ� 1000�Ǳ����� ���� �˴ϴ�.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td rowspan="2" colspan="2" onClick="jstrSort('ddate'); return false;" style="cursor:hand;">
		�Ⱓ
		<img src="/images/list_lineup<%=CHKIIF(vSorting="ddateD","_bot","_top")%><%=CHKIIF(instr(vSorting,"ddate")>0,"_on","")%>.png" id="imgddate">
	</td>
    <td colspan="2">�����(+)</td>
    <td colspan="2">�����(-)</td>
    <td colspan="2">������հ�</td>
    <td width="150" rowspan="2" onClick="jstrSort('sumpaymentetc'); return false;" style="cursor:hand;">
    	��������<br>(���ϸ��� ����)
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="sumpaymentetcD","_bot","_top")%><%=CHKIIF(instr(vSorting,"sumpaymentetc")>0,"_on","")%>.png" id="imgsumpaymentetc">
    </td>
	<td width="150" rowspan="2" onClick="jstrSort('miletotalprice'); return false;" style="cursor:hand;">
		���ϸ���
		<img src="/images/list_lineup<%=CHKIIF(vSorting="miletotalpriceD","_bot","_top")%><%=CHKIIF(instr(vSorting,"miletotalprice")>0,"_on","")%>.png" id="imgmiletotalprice">
	</td>
    <td width="150" rowspan="2" onClick="jstrSort('subtotalprice_notexists_sumpaymentetc'); return false;" style="cursor:hand;">
    	�ǰ����Ѿ�
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="subtotalprice_notexists_sumpaymentetcD","_bot","_top")%><%=CHKIIF(instr(vSorting,"subtotalprice_notexists_sumpaymentetc")>0,"_on","")%>.png" id="imgsubtotalprice_notexists_sumpaymentetc">
    </td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td onClick="jstrSort('countplus'); return false;" style="cursor:hand;">
    	�ֹ��Ǽ�
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="countplusD","_bot","_top")%><%=CHKIIF(instr(vSorting,"countplus")>0,"_on","")%>.png" id="imgcountplus">
    </td>
    <td onClick="jstrSort('maechulplus'); return false;" style="cursor:hand;">
    	�ݾ�
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="maechulplusD","_bot","_top")%><%=CHKIIF(instr(vSorting,"maechulplus")>0,"_on","")%>.png" id="imgmaechulplus">
    </td>
    <td onClick="jstrSort('countminus'); return false;" style="cursor:hand;">
    	�ֹ��Ǽ�
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="countminusD","_bot","_top")%><%=CHKIIF(instr(vSorting,"countminus")>0,"_on","")%>.png" id="imgcountminus">
    </td>
    <td onClick="jstrSort('maechulminus'); return false;" style="cursor:hand;">
    	�ݾ�
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="maechulminusD","_bot","_top")%><%=CHKIIF(instr(vSorting,"maechulminus")>0,"_on","")%>.png" id="imgmaechulminus">
    </td>
    <td onClick="jstrSort('count_plus_minus'); return false;" style="cursor:hand;">
    	�ֹ��Ǽ�
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="count_plus_minusD","_bot","_top")%><%=CHKIIF(instr(vSorting,"count_plus_minus")>0,"_on","")%>.png" id="imgcount_plus_minus">
    </td>
    <td onClick="jstrSort('maechul_plus_minus'); return false;" style="cursor:hand;">
    	�ݾ�
    	<img src="/images/list_lineup<%=CHKIIF(vSorting="maechul_plus_minusD","_bot","_top")%><%=CHKIIF(instr(vSorting,"maechul_plus_minus")>0,"_on","")%>.png" id="imgmaechul_plus_minus">
    </td>
</tr>

<% if cStatistic.FTotalCount > 0 then %>
	<% For i = 0 To cStatistic.FTotalCount -1 %>
	<tr bgcolor="#FFFFFF">
		<td align="center">
			<% if right(FormatDateTime(cStatistic.FItemList(i).FRegdate,1),3) = "�����" then %>
				<font color="blue"><%= cStatistic.FItemList(i).FRegdate %></font>
			<% elseif right(FormatDateTime(cStatistic.FItemList(i).FRegdate,1),3) = "�Ͽ���" then %>
				<font color="red"><%= cStatistic.FItemList(i).FRegdate %></font>
			<% else %>
				<%= cStatistic.FItemList(i).FRegdate %>
			<% end if %>
		</td>
		<td align="center"><%= getWeekdayStr(DatePart("w",cStatistic.FItemList(i).FRegdate)) %></td>
		<td align="center"><%= FormatNumber(cStatistic.FItemList(i).FCountPlus,0) %></td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).FMaechulPlus,0) %></td>
		<td align="center"><%= FormatNumber(cStatistic.FItemList(i).FCountMinus,0) %></td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).FMaechulMinus,0) %></td>
		<td align="center"><%= FormatNumber(cStatistic.FItemList(i).fcount_plus_minus) %></td>
		<td align="right" style="padding-right:5px;" bgcolor="#E6B9B8"><b><%= FormatNumber(cStatistic.FItemList(i).fmaechul_plus_minus,0) %></b></td>
		<td align="right"><%= FormatNumber(cStatistic.FItemList(i).FsumPaymentEtc,0) %></td>
		<td align="right"><%= FormatNumber(cStatistic.FItemList(i).FMiletotalprice,0) %></td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fsubtotalprice_notexists_sumPaymentEtc,0) %></td>
	</tr>
	<%
	vTot_CountPlus			= vTot_CountPlus + CLng(FormatNumber(cStatistic.FItemList(i).FCountPlus,0))
	vTot_MaechulPlus		= vTot_MaechulPlus + CLng(FormatNumber(cStatistic.FItemList(i).FMaechulPlus,0))
	vTot_CountMinus			= vTot_CountMinus + CLng(FormatNumber(cStatistic.FItemList(i).FCountMinus,0))
	vTot_MaechulMinus		= vTot_MaechulMinus + CLng(FormatNumber(cStatistic.FItemList(i).FMaechulMinus,0))
	vTot_MaechulCountSum	= vTot_MaechulCountSum + CLng(FormatNumber(cStatistic.FItemList(i).fcount_plus_minus,0))
	vTot_MaechulPriceSum	= vTot_MaechulPriceSum + CLng(FormatNumber(cStatistic.FItemList(i).fmaechul_plus_minus,0))
	vTot_sumPaymentEtc		= vTot_sumPaymentEtc + CLng(FormatNumber(cStatistic.FItemList(i).FsumPaymentEtc,0))
	vTot_Miletotalprice		= vTot_Miletotalprice + CLng(FormatNumber(cStatistic.FItemList(i).FMiletotalprice,0))
	vTot_Subtotalprice		= vTot_Subtotalprice + CLng(FormatNumber(cStatistic.FItemList(i).FSubtotalprice,0))
	vTot_subtotalprice_notexists_sumPaymentEtc = vTot_subtotalprice_notexists_sumPaymentEtc + CLng(FormatNumber(cStatistic.FItemList(i).fsubtotalprice_notexists_sumPaymentEtc,0))
	Next
	%>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" colspan="2">�հ�</td>
		<td align="center"><%=FormatNumber(vTot_CountPlus)%></td>
		<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_MaechulPlus,0)%></td>
		<td align="center"><%=FormatNumber(vTot_CountMinus)%></td>
		<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_MaechulMinus,0)%></td>
		<td align="center"><%=FormatNumber(vTot_MaechulCountSum)%></td>
		<td align="right" style="padding-right:5px;"><b><%=FormatNumber(vTot_MaechulPriceSum,0)%></b></td>
		<td align="right"><%=FormatNumber(vTot_sumPaymentEtc,0)%></td>
		<td align="right"><%=FormatNumber(vTot_Miletotalprice,0)%></td>
		<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_subtotalprice_notexists_sumPaymentEtc,0)%></td>
	</tr>
<% ELSE %>
	<tr  align="center" bgcolor="#FFFFFF">
		<td colspan="25">��ϵ� ������ �����ϴ�.</td>
	</tr>
<% end if %>

</table>

<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<script>
<!--
FusionCharts.ready(function () {
    var revenueChart = new FusionCharts({
        type: 'msline',
        renderAt: 'chart-container',
        width: '<%=35*cStatistic.FTotalCount%>',
        height: '500',
        dataFormat: 'json',
        dataSource: {
            "chart": {
				"caption": "��������-�Ϻ�",
                "subCaption": "<%=vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay) & " ~ " & vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay) %>",
				"formatnumberscale": "0",
				"paletteColors": "#0075c2,#1aaf5d",
				"labeldisplay": "ROTATE",
                "bgcolor": "#ffffff",
                "showBorder": "0",
                "showShadow": "0",
                "showCanvasBorder": "0",
                "usePlotGradientColor": "0",
                "legendBorderAlpha": "0",
                "legendShadow": "0",
                "showAxisLines": "0",
                "showAlternateHGridColor": "0",
                "divlineThickness": "1",
                "divLineIsDashed": "1",
                "divLineDashLen": "1",
                "divLineGapLen": "1",
                "xAxisName": "Day",
                "showValues": "0"
            },            
            "categories": [
                {
                    "category": [
						<%
						if cStatistic.FTotalCount > 0 then
							For i = 0 To cStatistic.FTotalCount -1
								Response.Write "{" & vbCrLf
								Response.Write """label"": """ & cStatistic.FItemList(i).FRegdate & """," & vbCrLf
								Response.Write "}"
								If i <> cStatistic.FTotalCount-1 Then
									Response.Write ","
								End If
								Response.Write vbCrLf
							Next
						End If
						%>
                    ]
                }
            ],            
            "dataset": [
                {
                    "seriesname": "�Ϻ� ����",
                    "data": [
						<%
						if cStatistic.FTotalCount > 0 then
							For i = 0 To cStatistic.FTotalCount -1
								Response.Write "{" & vbCrLf
								Response.Write """value"": """ & cStatistic.FItemList(i).fmaechul_plus_minus & """" & vbCrLf
								Response.Write "}"
								If i <> cStatistic.FTotalCount-1 Then
									Response.Write ","
								End If
								Response.Write vbCrLf
							Next
						End If
						%>
                    ]
                }
            ]
        }
    }).render();    
});
//-->
</script>
<div id="chart-container">FusionCharts will render here</div>
<%
Set cStatistic = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->