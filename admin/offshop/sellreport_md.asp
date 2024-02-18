<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 MD별 매출통계
' History : 2012.05.10 강준구 생성(기존매뉴 이전생성)
'			2013.01.24 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellreportmd_cls.asp"-->
<%
dim page,shopid ,yyyymmdd1,yyymmdd2 ,offgubun ,oldlist ,fromDate,toDate ,yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim i, sum1, sum2, sum3 ,makerid ,datefg , parameter ,CurrencyUnit, CurrencyChar, ExchangeRate ,FmNum, vOffCateCode, vOffMDUserID
dim dategubun, inc3pl
	dategubun = requestCheckVar(request("dategubun"),1)
	shopid = requestCheckVar(request("shopid"),32)
	page = requestCheckVar(request("page"),10)
	if page="" then page=1
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	yyyy2 = requestCheckVar(request("yyyy2"),4)
	mm2 = requestCheckVar(request("mm2"),2)
	dd2 = requestCheckVar(request("dd2"),2)
	offgubun = requestCheckVar(request("offgubun"),10)
	oldlist = requestCheckVar(request("oldlist"),2)
	makerid = requestCheckVar(request("makerid"),32)
	datefg = requestCheckVar(request("datefg"),16)
	vOffCateCode = requestCheckVar(request("offcatecode"),32)
	vOffMDUserID = requestCheckVar(request("offmduserid"),32)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if datefg = "" then datefg = "maechul"
if dategubun = "" then dategubun = "G"	
	
sum1 =0
sum2 =0
sum3 =0

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-7)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

'/매장
if (C_IS_SHOP) then
	
	'/어드민권한 점장 미만
	if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID
	end if
else
	'/업체
	if (C_IS_Maker_Upche) then
		makerid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
		else
		end if
	end if
end if

if shopid<>"" then offgubun=""

dim ooffsell
set ooffsell = new COffShopSellReportMD
	ooffsell.FRectShopID = shopid
	ooffsell.FRectNormalOnly = "on"
	ooffsell.FRectStartDay = fromDate
	ooffsell.FRectEndDay = toDate
	ooffsell.FRectOffgubun = offgubun
	ooffsell.FRectOldData = oldlist
	ooffsell.frectmakerid = makerid
	ooffsell.frectdatefg = datefg
	ooffsell.frectdategubun = dategubun
	ooffsell.frectoffcatecode = vOffCateCode
	ooffsell.frectoffmduserid = vOffMDUserID
	ooffsell.FCurrPage = page
	ooffsell.FRectInc3pl = inc3pl	
	ooffsell.Fpagesize=5000
	ooffsell.GetMDSellSumList

'Call fnGetOffCurrencyUnit(shopid,CurrencyUnit, CurrencyChar, ExchangeRate)
'FmNum = CHKIIF(CurrencyUnit="WON" or CurrencyUnit="KRW",0,2)

parameter = "&datefg="& datefg &"&shopid="& shopid &"&offgubun="& offgubun &"&oldlist="& oldlist &"&yyyy1="&yyyy1&"&mm1="&mm1&"&dd1="&dd1&"&yyyy2="&yyyy2&"&mm2="&mm2&"&dd2="&dd2&"&offcatecode="&vOffCateCode&"&offmduserid="&vOffMDUserID&"&makerid="&makerid&""
%>

<script language="javascript">

function frmsubmit(){
	frm.submit();
}

function goExceldown()
{
	document.location.href = "sellreport_md_xls.asp?1=1<%=parameter%>";
}

</script>

<!-- 표 상단바 시작-->
<table width="100%" cellpadding="1" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">  
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* 기간 : <% drawmaechul_datefg "datefg" ,datefg ,""%>
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >3년이전
				&nbsp;&nbsp;
				<%
				'직영/가맹점
				if (C_IS_SHOP) then
				%>	
					<% if getoffshopdiv(shopid) <> "1" and shopid <> "" then %>
						* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
					<% end if %>
				<% else %>
					* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
				<% end if %>
				<p>
				* 브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
				&nbsp;&nbsp;
				* 매장구분 : <% Call DrawShopDivCombo("offgubun",offgubun) %>
				&nbsp;&nbsp;
				* 카테고리 : <% SelectBoxBrandCategory "offcatecode", vOffCateCode %>
				&nbsp;&nbsp;
				* 오프라인담당MD : <% drawSelectBoxCoWorker_OnOff "offmduserid", vOffMDUserID, "off" %>
				<p>
	            <b>* 매출처구분</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>				
			</td>
		</tr>
		</table> 
    </td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onclick="frmsubmit();">
	</td>
</tr>
</table>
<!-- 표 상단바 끝-->
<br>
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<table width="100%" class="a">
		<tr>
			<td>검색결과 : <b><%= ooffsell.FResultCount %></b> ※ 최대 5000건 까지 검색됩니다.&nbsp;&nbsp;※ 브랜드에 담당MD가 지정된 것만 검색이 됩니다.</td>
			<td align="right"><img src="http://webadmin.10x10.co.kr/images/btn_excel.gif" style="cursor:pointer;" onClick="goExceldown();"></td>
		</tr>
		</table>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>담당자</td>
	<td>브랜드</td>
	<td>매출액</td>
	<td>매입가합</td>
	<td>마진율</td>
	<td>총매출액</td>
</tr>
<%
	Dim vBody, vTotalTD, v, vTmpMDname, vTmpCnt, vTotalSum
	v = vbCrLf
	vTmpMDname = ""
	vTotalSum = 0
	For i=0 To ooffsell.FresultCount-1
		
		vBody = vBody & "<tr bgcolor=""#FFFFFF"" align=""center"">" & v

		If i = 0 Then
			vTotalSum = vTotalSum + ooffsell.FItemList(i).FSum
		End If

		If vTmpMDname <> ooffsell.FItemList(i).Fmdname Then
			vBody = Replace(vBody,"id='nametd'","rowspan="""&vTmpCnt&"""")
			vBody = Replace(vBody,"id='totaltd'>","rowspan="""&vTmpCnt&""">"&FormatNumber(vTotalSum,0)&"")
			
			vBody = vBody & "	<td id='nametd'>" & ooffsell.FItemList(i).Fmdname & "</td>" & v
			vTotalTD = "	<td id='totaltd'></td>" & v
			vTmpCnt = "1"
			If i <> 0 Then
				vTotalSum = ooffsell.FItemList(i).FSum
			End IF
		Else
			vTotalTD = ""
			vTmpCnt = vTmpCnt + 1
			vTotalSum = vTotalSum + ooffsell.FItemList(i).FSum
		End If
		
		If ooffsell.FItemList(i).FChargeDiv = "6" Then
			vBody = vBody & "	<td><b><font color=""#3333CC""><a href=""javascript:PopBrandInfoEdit('" & ooffsell.FItemList(i).FMakerid & "')"">" & ooffsell.FItemList(i).FMakerid & "</a></font></b></td>" & v
		Else
			vBody = vBody & "	<td><a href=""javascript:PopBrandInfoEdit('" & ooffsell.FItemList(i).FMakerid & "')"">" & ooffsell.FItemList(i).FMakerid & "</a></td>" & v
		End If
		
		vBody = vBody & "	<td style=""padding-right:5px;"" align=""right"" bgcolor=""#E6B9B8"">" & FormatNumber(ooffsell.FItemList(i).FSum,0) & "</td>" & v
		vBody = vBody & "	<td style=""padding-right:5px;"" align=""right"">" & FormatNumber(ooffsell.FItemList(i).fsuplyprice,0) & "</td>" & v
		vBody = vBody & "	<td style=""padding-right:5px;"" align=""right"">"
		
		If ooffsell.FItemList(i).fsuplyprice > 0 and ooffsell.FItemList(i).FSum > 0 Then
			vBody = vBody & "" & FormatNumber(100-ooffsell.FItemList(i).fsuplyprice/ooffsell.FItemList(i).FSum*100,0) & "%"
		Else
			vBody = vBody & "0%"
		End If
		
		vBody = vBody & "	</td>" & v
		vBody = vBody & vTotalTD
		vBody = vBody & "</tr>" & v
		
		vTmpMDname = ooffsell.FItemList(i).Fmdname
		
		If i = ooffsell.FresultCount-1 Then
			vBody = Replace(vBody,"id='nametd'","rowspan="""&vTmpCnt&"""")
			vBody = Replace(vBody,"id='totaltd'>","rowspan="""&vTmpCnt&""">"&FormatNumber(vTotalSum,0)&"")
		End IF
	Next
	
	Response.Write vBody
%>
</table>

<%
set ooffsell = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->