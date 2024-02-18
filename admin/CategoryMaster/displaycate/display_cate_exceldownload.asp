<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 전시카테고리 엑셀다운로드
' Hieditor : 2020.01.06 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" --> 
<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
dim menupos, ocate, i
    menupos= requestCheckvar(request("menupos"),10)

set ocate = new cDispCate
    ocate.GetDispCateAllList

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN_ITEM" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
Response.Buffer = true    '버퍼사용여부
%>
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>

<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>1depth</td>
	<td>코드</td>
	<td>한글카테고리명</td>
	<td>영문카테고리명</td>
	<td>2depth</td>
	<td>코드</td>
	<td>한글카테고리명</td>
	<td>영문카테고리명</td>
	<td>3depth</td>
	<td>코드</td>
	<td>한글카테고리명</td>
	<td>영문카테고리명</td>
	<td>4depth</td>
	<td>코드</td>
	<td>한글카테고리명</td>
	<td>영문카테고리명</td>
</tr>
<% if ocate.FresultCount>0 then %>
	<% for i=0 to ocate.FresultCount -1 %>
	<tr bgcolor="#FFFFFF" align="center">
		<td><%= ocate.FItemList(i).fdepth1 %></td>
		<td class="txt"><%= ocate.FItemList(i).fcatecode1 %></td>
		<td align="left"><%= ocate.FItemList(i).fcatename1 %></td>
		<td align="left"><%= ocate.FItemList(i).fcatename_e1 %></td>
		<td><%= ocate.FItemList(i).fdepth2 %></td>
		<td class="txt"><%= ocate.FItemList(i).fcatecode2 %></td>
		<td align="left"><%= ocate.FItemList(i).fcatename2 %></td>
		<td align="left"><%= ocate.FItemList(i).fcatename_e2 %></td>
		<td><%= ocate.FItemList(i).fdepth3 %></td>
		<td class="txt"><%= ocate.FItemList(i).fcatecode3 %></td>
		<td align="left"><%= ocate.FItemList(i).fcatename3 %></td>
		<td align="left"><%= ocate.FItemList(i).fcatename_e3 %></td>
		<td><%= ocate.FItemList(i).fdepth4 %></td>
		<td class="txt"><%= ocate.FItemList(i).fcatecode4 %></td>
		<td align="left"><%= ocate.FItemList(i).fcatename4 %></td>
		<td align="left"><%= ocate.FItemList(i).fcatename_e4 %></td>
	</tr>
	<%
    if i mod 3000 = 0 then
        Response.Flush		' 버퍼리플래쉬
    end if
    next
    %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="16" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
<%
set ocate=nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->