<%@ language=vbscript %>
<% option Explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"

Server.ScriptTimeOut = 60*10		' 10분
%>
<%
'###########################################################
' Description : 장바구니 위시 쇼핑액션 통계
' History : 2023.06.27 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/crm/BaguniWishShoppingActionCls.asp"-->
<%
dim page, research, yyyy1,mm1,dd1,yyyy2,mm2,dd2, fromDate,toDate, i
    page = RequestCheckVar(getNumeric(request("page")),10)
    research = RequestCheckVar(request("research"),2)
    yyyy1 = RequestCheckVar(request("yyyy1"),4)
    mm1   = RequestCheckVar(request("mm1"),2)
    dd1   = RequestCheckVar(request("dd1"),2)
    yyyy2 = RequestCheckVar(request("yyyy2"),4)
    mm2   = RequestCheckVar(request("mm2"),2)
    dd2   = RequestCheckVar(request("dd2"),2)

if (yyyy1="") then yyyy1 = Cstr(Year(dateadd("d",-1,date())))
if (mm1="") then mm1 = Cstr(Month(dateadd("d",-1,date())))
if (dd1="") then dd1 = Cstr(day(dateadd("d",-1,date())))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
if (page="") then page=1
fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

dim oaction
set oaction = new CActionList
    oaction.FCurrPage = page
    oaction.FPageSize = 100
    oaction.FRectStartDate = fromDate
    oaction.FRectEndDate   = toDate
    oaction.GetBaguniWishShoppingActionList
%>
<script type='text/javascript'>

function NextPage(page){
	document.frm.target = "";
	document.frm.action = "";
    document.frm.page.value=page;
    document.frm.submit();
}

function Actiondownloadexcel(){
	document.frm.target = "view";
	document.frm.action = "/admin/crm/BaguniWishShoppingAction_excel.asp";
	document.frm.submit();
	document.frm.target = "";
	document.frm.action = "";
}

</script>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
    <td align="left">
		* 날짜 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
    </td>
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
        <input type="button" class="button_s" value="검색" onClick="NextPage('1');">
    </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td align="left"></td>
</tr>
</table>
</form>
<br>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
    <td align="left">
        ※ <%= CStr(DateSerial(yyyy1, mm1, dd1)) %>~<%= CStr(DateSerial(yyyy2, mm2, dd2)) %>에 쇼핑액션(장바구니,위시)을 하였으나, <%= CStr(DateSerial(yyyy1, mm1, dd1)) %>~현재까지 미구매한 고객 리스트 입니다.<br>느린 매뉴 입니다. 클릭후 기다려 주세요.
    </td>
    <td align="right">
		<input type="button" onclick="Actiondownloadexcel();" value="엑셀다운로드" class="button">
    </td>
</tr>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan="8">
		검색결과 : <b><%= oaction.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oaction.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>고객아이디</td>
    <td>고객명</td>
    <td>회원등급</td>
    <td>푸시수신</td>
    <td>문자수신</td>
    <td>이메일수신</td>
    <td>마지막로그인</td>
    <td>비고</td>
</tr>
<% if oaction.FResultCount>0 then %>
<% for i=0 to oaction.FResultCount-1 %>
<tr bgcolor="#FFFFFF" align="center">
    <td>
        <% if C_CriticInfoUserLV1 then %>
            <%= oaction.FItemList(i).fuserid %>
        <% else %>
            <%= printUserId(oaction.FItemList(i).fuserid,2,"*") %>
        <% end if %>
    </td>
    <td>
        <% if C_CriticInfoUserLV1 then %>
            <%= oaction.FItemList(i).fusername %>
        <% else %>
            <%= printUserId(oaction.FItemList(i).fusername,2,"*") %>
        <% end if %>
    </td>
    <td><%= oaction.FItemList(i).fuserlevel %></td>
    <td><%= oaction.FItemList(i).fpushYn %></td>
    <td><%= oaction.FItemList(i).fsmsok %></td>
    <td><%= oaction.FItemList(i).femailok %></td>
    <td><%= oaction.FItemList(i).flastlogin %></td>
    <td></td>
</tr>
<% next %>

<tr bgcolor="FFFFFF">
	<td colspan="8" align="center">
		<% if oaction.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oaction.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oaction.StartScrollPage to oaction.FScrollCount + oaction.StartScrollPage - 1 %>
			<% if i>oaction.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oaction.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="8" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height=300 frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no"></iframe>
<% end if %>

<%
set oaction = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
