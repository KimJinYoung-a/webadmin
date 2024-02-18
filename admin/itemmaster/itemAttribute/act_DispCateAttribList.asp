<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemAttribCls.asp"-->
<%
'###############################################
' Discription : 전시카테고리-상품속성 목록 Ajax
' History : 2013.08.06 허진원 : 신규 생성
'###############################################
Response.CharSet = "euc-kr"

'// 변수 선언
Dim dispCate
Dim oAttrib, lp
Dim page

'// 파라메터 접수
dispCate = request("dispcate")
page = request("page")
if page="" then page="1"

'// 페이지정보 목록
	set oAttrib = new CAttrib
	oAttrib.FPageSize = 40
	oAttrib.FCurrPage = page
	oAttrib.FRectDispCate = dispCate
    oAttrib.GetDispCateAttribList
%>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="2">
		검색결과 : <b><%=oAttrib.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oAttrib.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>카테고리</td>
    <td>속성구분</td>
</tr>
<%	for lp=0 to oAttrib.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
    <td><a href="javascript:viewDispCateAttrib('<%=oAttrib.FItemList(lp).Fcatecode%>')"><%="[" & oAttrib.FItemList(lp).Fcatecode & "] " & Replace(oAttrib.FItemList(lp).Fcatename,"^^"," > ") %></a></td>
	<td><%="[" & oAttrib.FItemList(lp).FattribDiv & "] " & oAttrib.FItemList(lp).FattribDivName%></td>
</tr>
<%	Next %>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center">
    <% if oAttrib.HasPreScroll then %>
		<a href="javascript:goPage('<%=dispCate%>','<%= oAttrib.StartScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for lp=0 + oAttrib.StartScrollPage to oAttrib.FScrollCount + oAttrib.StartScrollPage - 1 %>
		<% if lp>oAttrib.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(lp) then %>
		<font color="red">[<%= lp %>]</font>
		<% else %>
		<a href="javascript:goPage('<%=dispCate%>','<%= lp %>');">[<%= lp %>]</a>
		<% end if %>
	<% next %>

	<% if oAttrib.HasNextScroll then %>
		<a href="javascript:goPage('<%=dispCate%>','<%= lp %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>
<%
	set oAttrib = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->