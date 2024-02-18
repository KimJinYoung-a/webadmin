<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 삽별구역설정
' Hieditor : 2011.11.25 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/zone2/zone_cls.asp"-->
<%
dim omanager ,i

set omanager = new czone_list
	omanager.FPageSize = 500
	omanager.FCurrPage = 1
    'omanager.frectpart_sn = "18"
	omanager.Getshopzonemanager
%>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">

	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= omanager.fresultcount %></b>
	</td>
</tr>
<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>사원번호</td>
	<td>이름</td>
	<td>비고</td>
</tr>
<% if omanager.FresultCount > 0 then %>
<%
for i=0 to omanager.FresultCount - 1
%>
<tr align="center" bgcolor="#FFFFFF">
	<td><%= omanager.FItemList(i).fempno %></td>
	<td><%= omanager.FItemList(i).fusername %></td>
	<td><input type="button" class="button" value="선택" onclick="opener.addSelectedmanager('<%= omanager.FItemList(i).fempno %>','<%= omanager.FItemList(i).fusername %>')"></td>
</tr>
<%
next
else
%>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan=15>검색 결과가 없습니다</td>
</tr>
<%
end if
%>
</table>

<%
set omanager = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->