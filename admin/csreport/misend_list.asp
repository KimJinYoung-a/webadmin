<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/oldmisendcls.asp" -->
<%
response.write "사용중지 - 서팀 문의 요망"
response.end

dim i, ix
dim page
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nowdateStr, startdateStr, nextdateStr

page = request("page")
if (page = "") then
        page = "1"
end if


yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

nowdateStr = CStr(now())


if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

startdateStr = yyyy1 + "-" + mm1 + "-" + dd1
nextdateStr = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

dim oldmisend, delaydate
delaydate = request("delaydate")

if delaydate="" then delaydate=5

set oldmisend = New COldMiSend
oldmisend.FPageSize = 30
oldmisend.FRectDelayDate = delaydate
oldmisend.FRectStart = startdateStr
oldmisend.FRectEnd =  nextdateStr
oldmisend.FRectNotInCludeUpcheCheck = "on"

oldmisend.GetOldMisendListSearch
%>

<script language="JavaScript">
<!--

function NextPage(ipage){
	document.noticeform.page.value= ipage;
	document.noticeform.submit();
}

//-->
</script>



<table width="800"  class="a">
<tr>
	<td width=100></td>
	<td align="center">****** <%= delaydate %>일 이상 미배송 목록 (최대 <%= oldmisend.FPageSize %>건) ******</td>
	<td width=100 align="right"><a href="/admin/ordermaster/newmisendlist.asp"><font color=red>전체보기</font></a></td>
</tr>
</table>
<table width="800" border="1" bordercolordark="White" bordercolorlight="black" cellpadding="0" cellspacing="0" class="a">
  <tr bgcolor="#DDDDFF" height="25">
    <td width="70" align="center">주문번호</td>
    <td width="66" align="center">입금일</td>
    <td width="40" align="center">주문인</td>
    <td width="80" align="center">업체ID</td>
    <td width="50" align="center">상품번호</td>
    <td align="center">상품명</td>
    <td width="60" align="center">옵션</td>
    <td width="60" align="center">배송구분</td>
    <td width="60" align="center">상태</td>
    <td width="100" align="center">비고</td>
  </tr>
<% for i = 0 to oldmisend.FResultCount - 1 %>
  <tr height="20">
    <td><%= oldmisend.FItemList(i).ForderSerial %></td>
    <td><%= Left(oldmisend.FItemList(i).FIpkumDate,10) %></td>
    <td><%= oldmisend.FItemList(i).FBuyName %></td>
    <td><%= oldmisend.FItemList(i).FMakerID %></td>
    <td><%= oldmisend.FItemList(i).FItemID %></td>
    <td><%= oldmisend.FItemList(i).FItemName %></td>
    <td><%= oldmisend.FItemList(i).GetOptionName %></td>
    <td><font color="<%= oldmisend.FItemList(i).GetBeagonGubunColor %>"><%= oldmisend.FItemList(i).GetBeagonGubunName %></font></td>
    <td><font color="<%= oldmisend.FItemList(i).GetBeagonStateColor %>"><%= oldmisend.FItemList(i).GetBeagonStateName %></font></td>
    <td><%= oldmisend.FItemList(i).getMiSendCodeName %> <%= oldmisend.FItemList(i).getIpgoMayDay %></td>
  </tr>
<% next %>
</table>
<%
set oldmisend = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->