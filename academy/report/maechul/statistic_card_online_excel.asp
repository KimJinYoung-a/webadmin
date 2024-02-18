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
<!-- #include virtual="/academy/lib/classes/diysell_reportcls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2, vIsOldOrder
dim yyyymmdd1,yyymmdd2
dim fromDate,toDate
dim ordertype, vSiteName, vSorting

yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1 = RequestCheckvar(request("mm1"),2)
dd1 = RequestCheckvar(request("dd1"),2)
yyyy2 = RequestCheckvar(request("yyyy2"),4)
mm2 = RequestCheckvar(request("mm2"),2)
dd2 = RequestCheckvar(request("dd2"),2)
vSiteName 	= RequestCheckvar(request("sitename"),16)
vSorting	= NullFillWith(RequestCheckvar(request("sorting"),32),"ddateD")

If vIsOldOrder = "" Then
	vIsOldOrder = "n"
End If

ordertype = RequestCheckvar(request("ordertype"),1)
if ordertype = "" then ordertype = "D"

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = "1"

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

dim oreport
set oreport = new CDiyReportMaster
oreport.FRectFromDate = fromDate
oreport.FRectToDate = toDate
oreport.FRectSiteName = vSiteName
oreport.FRectSort = vSorting
if ordertype="D" then
oreport.SearchCardOnline
else
oreport.SearchCardOnlineMonth
end if

dim i,p1,p2
dim prename
dim buftext, bufname, bufimage
dim sumtotal
dim ch1,ch2,ch3,ch4,sellcnt1,sellcnt2,sellcnt3,sellcnt4

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
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
		<td>
			기간
		</td>
		<td>
			신용카드
		</td>
		<td>
			실시간계좌이체
		</td>
		<td>
			무통장
		</td>
		<td>
			수기입력
		</td>
		<td>
			총액
		</td>
	</tr>
	<% for i=0 to oreport.FResultCount-1 %>
	<% if (prename<>oreport.FMasterItemList(i).Fsitename) then %>
	<% if (prename<>"") then %>
	<tr bgcolor="#FFFFFF">
		<td align="center">
			<% =bufname %>
		</td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(ch2,0) & sellcnt1 %></td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(ch3,0) & sellcnt2 %></td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(ch1,0) & sellcnt3 %></td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(ch4,0) & sellcnt4 %></td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(ch1+ch2+ch3+ch4,0) %></td>
	</tr>
	<%
		buftext = ""
		bufimage = ""
		sumtotal = 0
		ch1 = 0
		ch2 = 0
		ch3 = 0
		ch4 = 0
		sellcnt1=""
		sellcnt2=""
		sellcnt3=""
		sellcnt4=""
	%>
	<% end if %>
	<% end if %>
	<%
	If ordertype="D" Then
			bufname = oreport.FMasterItemList(i).Fsitename + "(" + oreport.FMasterItemList(i).GetDpartName + ")"
	Else
		bufname = oreport.FMasterItemList(i).Fsitename
	End If
	prename = oreport.FMasterItemList(i).Fsitename
	if oreport.FMasterItemList(i).Faccountdiv=7 then'무통
		ch1 = oreport.FMasterItemList(i).Fselltotal
		sellcnt1 = " (" + FormatNumber(oreport.FMasterItemList(i).Fsellcnt,0) + "건)"
	elseif oreport.FMasterItemList(i).Faccountdiv=100 then'신용
		ch2 = oreport.FMasterItemList(i).Fselltotal
		sellcnt2 = " (" + FormatNumber(oreport.FMasterItemList(i).Fsellcnt,0) + "건)"
	elseif oreport.FMasterItemList(i).Faccountdiv=20 then'실시간
		ch3 = oreport.FMasterItemList(i).Fselltotal
		sellcnt3 = " (" + FormatNumber(oreport.FMasterItemList(i).Fsellcnt,0) + "건)"
	elseif oreport.FMasterItemList(i).Faccountdiv=900 then'수기
		ch4 = oreport.FMasterItemList(i).Fselltotal
		sellcnt4 = " (" + FormatNumber(oreport.FMasterItemList(i).Fsellcnt,0) + "건)"
	end if
	%>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td align="center">
			<% =bufname %>
		</td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(ch2,0) & sellcnt1 %></td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(ch3,0) & sellcnt2 %></td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(ch1,0) & sellcnt3 %></td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(ch4,0) & sellcnt4 %></td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(ch1+ch2+ch3+ch4,0) %></td>
	</tr>
</table>

<%
Set oreport = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->