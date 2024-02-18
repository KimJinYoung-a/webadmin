<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  핑거스 매출집계-카테고리별
' History : 2016.03.15 corpse2 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/academy_reportcls.asp"-->
<%
dim oreport
dim stdate
dim yyyy1,mm1
Dim sort

yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1	  = RequestCheckvar(request("mm1"),2)
sort	  = RequestCheckvar(request("sort"),10)

if yyyy1="" then
	stdate = CStr(Now)
	yyyy1 = Left(stdate,4)
	mm1 = Mid(stdate,6,2)
end if

set oreport = new CJumunMaster
oreport.FRectFromDate = yyyy1 + "-" + mm1
oreport.FRectSort = sort
oreport.GetLecturerMonthMeaChul

Dim i,p1,p2
Dim premonth_sellsum,premonth_sellcnt

dim selltotal, sellcnt

'Response.Buffer=False
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
	<td>강사</td>
    <td>금액</td>
	<td>주문건수</td>
	<td>객단가</td>
</tr>
<% for i=0 to oreport.FResultCount-1 %>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= oreport.FMasterItemList(i).Fsitename %> (<%= oreport.FMasterItemList(i).Flecturer %>)</td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(oreport.FMasterItemList(i).Fselltotal,0) %>원</td>
	<td align="center"><%= FormatNumber(oreport.FMasterItemList(i).Fsellcnt,0) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#7CCE76"><% if sellcnt<>0 then %><%= FormatNumber(Clng(oreport.FMasterItemList(i).Fselltotal/oreport.FMasterItemList(i).Fsellcnt),0) %>원<% end if %></td>
</tr>
<%
selltotal = selltotal + oreport.FMasterItemList(i).Fselltotal
sellcnt = sellcnt + oreport.FMasterItemList(i).Fsellcnt
Next
%>
<tr bgcolor="#FFFFFF">
	<td align="center"></td>
    <td align="right"><%= FormatNumber(selltotal,0) %>원</td>
	<td align="center"><%= FormatNumber(sellcnt,0) %>건</td>
	<td align="right"><% if sellcnt<>0 then %><%= FormatNumber(selltotal/sellcnt,0) %>원<% end if %></td>
</tr>
</table>
<%
Set oreport = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->