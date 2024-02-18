<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  핑거스 매출집계-일별
' History : 2016.09.20 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/report/order_saacls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nowdateStr, startdateStr, nextdateStr, page
dim rpttype,addstand,oldDataYn, vSiteName

addstand = RequestCheckvar(request("addstand"),10)
if addstand = "" then addstand = 1
yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1 = RequestCheckvar(request("mm1"),2)
dd1 = RequestCheckvar(request("dd1"),2)
yyyy2 = RequestCheckvar(request("yyyy2"),4)
mm2 = RequestCheckvar(request("mm2"),2)
dd2 = RequestCheckvar(request("dd2"),2)
rpttype = RequestCheckvar(request("rpttype"),16)
oldDataYn=request("oldDataYn")
page = RequestCheckvar(request("page"),10)
vSiteName 	= RequestCheckvar(request("sitename"),16)

if page="" then page=1

nowdateStr = CStr(now())


if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = "01"
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

if (rpttype="") then rpttype="day"

startdateStr = yyyy1 + "-" + mm1 + "-01"
nextdateStr = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

dim orderreport
set orderreport = new UserJoinClass
orderreport.FRectStart = startdateStr
orderreport.FRectEnd =  nextdateStr
orderreport.FRectGroup = rpttype
orderreport.FoldDataYn = oldDataYn
orderreport.FRectSiteName=vSiteName

dim i

const MAXBARSIZE = 500
dim totno, MsexPercent,WsexPercent

totno = orderreport.FManNo + orderreport.FWoManNo

if totno<>0 then
	MsexPercent = CInt(orderreport.FManNo/totno*100)
	WsexPercent = CInt(orderreport.FWoManNo/totno*100)
else
	MsexPercent = 0
	WsexPercent = 0
end if

orderreport.GetUserJoinByNai

Dim Area1, Area2, Area3, Area4, Area5, Area6, Area7, Area8
Area1=0
Area2=0
Area3=0
Area4=0
Area5=0
Area6=0
Area7=0
Area8=0
dim orderreport2
set orderreport2 = new UserJoinClass
orderreport2.FRectStart = startdateStr
orderreport2.FRectEnd =  nextdateStr
orderreport2.FRectGroup = rpttype
orderreport2.FoldDataYn = oldDataYn
orderreport2.FRectSiteName=vSiteName
orderreport2.FRectBeasongArea = addstand
orderreport2.GetUserJoinByArea2
for i=0 to orderreport2.FResultCount -1
	If left(orderreport2.FItemList(i).FArea,1)="0" Then
		Area1=Area1+orderreport2.FItemList(i).FCount
	ElseIf orderreport2.FItemList(i).FArea>="10" And orderreport2.FItemList(i).FArea<="23" Then
		Area4=Area4+orderreport2.FItemList(i).FCount
	ElseIf orderreport2.FItemList(i).FArea>="24" And orderreport2.FItemList(i).FArea<="26" Then
		Area2=Area2+orderreport2.FItemList(i).FCount
	ElseIf orderreport2.FItemList(i).FArea>="27" And orderreport2.FItemList(i).FArea<="35" Then
		Area3=Area3+orderreport2.FItemList(i).FCount
	ElseIf orderreport2.FItemList(i).FArea>="36" And orderreport2.FItemList(i).FArea<="43" Then
		Area7=Area7+orderreport2.FItemList(i).FCount
	ElseIf (orderreport2.FItemList(i).FArea>="44" And orderreport2.FItemList(i).FArea<="53") Or orderreport2.FItemList(i).FArea="63" Then
		Area6=Area6+orderreport2.FItemList(i).FCount
	ElseIf orderreport2.FItemList(i).FArea>="54" And orderreport2.FItemList(i).FArea<="62" Then
		Area5=Area5+orderreport2.FItemList(i).FCount
	Else
		Area8=Area8+orderreport2.FItemList(i).FCount
	End If
Next

dim tmppercent, totalcnt
orderreport.FRectBeasongArea = addstand
orderreport.GetUserJoinByArea
totalcnt = orderreport.FTotalUsercount+orderreport2.FTotalUsercount

'Response.Buffer=False
Response.Charset = "euc-kr"
Response.Buffer = True
Response.AddHeader "charset", "euc-kr"  '// 이 코드가 한글 깨짐을 해결
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>

<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>

<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#DDDDFF">
    	<td colspan="4">2. 연령별 구매비율 / 주문수 (현재 나이기준)</td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td width="100">전체</td>
    	<td width="100" align="right">
    		<%= FormatNumber(orderreport.FNaiMaster.FManTotal,0) %><br>
    		<%= FormatNumber(orderreport.FNaiMaster.FWoManTotal,0) %>
    	</td>
    	<td width="50" align="right">
    		<%= orderreport.FNaiMaster.GetManTotalPercent %> (%)<br>
    		<%= orderreport.FNaiMaster.GetWoManTotalPercent %> (%)
    	</td>
    	<td width="50" align="right">100 (%)</td>
    </tr>
    <% for i=0 to orderreport.FNaiMaster.FItemCount - 1  %>
    <tr bgcolor="#FFFFFF">
    	<td width="100"><%= orderreport.FNaiMaster.FItemList(i).FNaiStr %></td>
    	<td width="100" align="right">
    		<%= FormatNumber(orderreport.FNaiMaster.FItemList(i).FManCount,0) %><br>
    		<%= FormatNumber(orderreport.FNaiMaster.FItemList(i).FWoManCount,0) %>
    	</td>
    	<td width="50" align="right">
    		<%= orderreport.FNaiMaster.GetManPercent(i) %> (%)<br>
    		<%= orderreport.FNaiMaster.GetWoManPercent(i) %> (%)
    	</td>
    	<td width="50" align="right"><%= orderreport.FNaiMaster.GetTotPercent(i) %> (%)</td>
    </tr>
    <% next %>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#DDDDFF">
    	<td colspan="3">3. 지역별 구매비율 / 주문수 </td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td width="120">전체</td>
    	<td width="100" align="right"><%= FormatNumber(orderreport.FTotalUsercount,0) %></td>
    	<td width="100" align="right">100 (%)</td>
    </tr>
    <% for i=0 to orderreport.FResultCount -1 %>
    <%
    if orderreport.FTotalUsercount=0 then
    	tmppercent = 0
    Else
		If orderreport.FItemList(i).FArea="1" Then
			tmppercent = round(orderreport.FItemList(i).FCount+Area1/totalcnt*100,2)
		ElseIf orderreport.FItemList(i).FArea="2" Then
			tmppercent = round(orderreport.FItemList(i).FCount+Area2/totalcnt*100,2)
		ElseIf orderreport.FItemList(i).FArea="3" Then
			tmppercent = round(orderreport.FItemList(i).FCount+Area3/totalcnt*100,2)
		ElseIf orderreport.FItemList(i).FArea="4" Then
			tmppercent = round(orderreport.FItemList(i).FCount+Area4/totalcnt*100,2)
		ElseIf orderreport.FItemList(i).FArea="5" Then
			tmppercent = round(orderreport.FItemList(i).FCount+Area5/totalcnt*100,2)
		ElseIf orderreport.FItemList(i).FArea="6" Then
			tmppercent = round(orderreport.FItemList(i).FCount+Area6/totalcnt*100,2)
		ElseIf orderreport.FItemList(i).FArea="7" Then
			tmppercent = round(orderreport.FItemList(i).FCount+Area7/totalcnt*100,2)
		End If
    end if
    %>
    <tr bgcolor="#FFFFFF">
    	<td width="120"><%= orderreport.FItemList(i).GetArea %> </td>
    	<td width="100" align="right">
			<% If orderreport.FItemList(i).FArea="1" Then %>
				<%= FormatNumber(orderreport.FItemList(i).FCount+Area1,0) %>
			<% ElseIf orderreport.FItemList(i).FArea="2" Then %>
				<%= FormatNumber(orderreport.FItemList(i).FCount+Area2,0) %>
			<% ElseIf orderreport.FItemList(i).FArea="3" Then %>
				<%= FormatNumber(orderreport.FItemList(i).FCount+Area3,0) %>
			<% ElseIf orderreport.FItemList(i).FArea="4" Then %>
				<%= FormatNumber(orderreport.FItemList(i).FCount+Area4,0) %>
			<% ElseIf orderreport.FItemList(i).FArea="5" Then %>
				<%= FormatNumber(orderreport.FItemList(i).FCount+Area5,0) %>
			<% ElseIf orderreport.FItemList(i).FArea="6" Then %>
				<%= FormatNumber(orderreport.FItemList(i).FCount+Area6,0) %>
			<% ElseIf orderreport.FItemList(i).FArea="7" Then %>
				<%= FormatNumber(orderreport.FItemList(i).FCount+Area7,0) %>
			<%End If%>
		</td>
    	<td width="100" align="right"><%= tmppercent %> (%)</td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
    	<td width="120">기타(비회원)</td>
    	<td width="100" align="right"><%= FormatNumber(Area8,0) %></td>
    	<td width="100" align="right"><%=CInt(Area8/totalcnt*100)%> (%)</td>
    	<td></td>
    </tr>
</table>

<%
Set orderreport = Nothing
Set orderreport2 = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->