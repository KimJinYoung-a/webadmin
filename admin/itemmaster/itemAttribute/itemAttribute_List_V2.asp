<%@ language=vbscript %>
<% option explicit %>
<!DOCTYPE html>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemAttribCls.asp"-->
<%

Dim page, i
dim masteridx

page = requestCheckVar(request("page"), 32)
masteridx = requestCheckVar(request("masteridx"), 32)

if page="" then page="1"

dim oAttrib
set oAttrib = new CAttrib
oAttrib.FPageSize = 20
oAttrib.FCurrPage = page
oAttrib.FRectMasterIDX = masteridx

oAttrib.GetAttribList_V2

%>
<!-- 상단 검색폼 시작 -->
<form name="frm" method="get" action="" style="margin:0;">
<input type="hidden" name="research" value="on" />
<input type="hidden" name="page" value="" />
<input type="hidden" name="menupos" value="<%= request("menupos") %>" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
	    속성구분:
	    <%= drawSelectAttributeMaster("masteridx", masteridx, "") %>
	</td>
	<td width="80" rowspan="2" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" value="검색" />
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10px 0 5px 0;">
<tr>
    <td align="left">
    	<input type="button" value="선택저장" class="button" onClick="saveList()" title="우선순위 및 사용여부를 일괄저장합니다.">
    </td>
    <td align="right">
    	<input type="button" value="신규속성 등록" class="button" onClick="popAttribute('');">
    </td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 목록 시작 -->
<form name="frmList" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="attrArr">
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" style="table-layout: fixed;">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		검색결과 : <b><%=oAttrib.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oAttrib.FtotalPage%></b>
	</td>
</tr>
<colgroup>
	<col width="40" />
    <col width="50" />
    <col width="80" />
    <col width="*" />
    <col width="80" />
    <col width="80" />
    <col width="*" />
    <col width="80" />
	<col width="160" />
</colgroup>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><span class="ui-icon ui-icon-arrowthick-2-n-s"></span></td>
    <td><input type="checkbox" name="allChk" onclick="chkAllItem()"></td>
    <td>IDX</td>
    <td>속성구분</td>
    <td>표시순서</td>
    <td>IDX</td>
    <td>속성상세</td>
    <td>표시순서</td>
	<td><span class="ui-icon ui-icon-wrench"></span></td>
</tr>
<tbody id="attrList">
<%	for i = 0 to oAttrib.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td><span class="rowHaddle ui-icon ui-icon-grip-solid-horizontal" style="cursor:grab;" title="정렬순서를 변경합니다."></span></td>
    <td><input type="checkbox" name="chkCd" value="<%= oAttrib.FItemList(i).Fidx %>" /></td>
    <td><%= oAttrib.FItemList(i).Fidx %></td>
    <td><%= oAttrib.FItemList(i).FattMasterName %></td>
    <td><%= oAttrib.FItemList(i).Fdispno %></td>
    <td><%= oAttrib.FItemList(i).Fdetailidx %></td>
    <td><%= oAttrib.FItemList(i).FattDetailName %></td>
    <td><%= oAttrib.FItemList(i).Fdetaildispno %></td>
	<td>
	</td>
</tr>
<%	Next %>
</tbody>
<tr bgcolor="#FFFFFF">
    <td colspan="9" align="center">
    <% if oAttrib.HasPreScroll then %>
		<a href="javascript:goPage('<%= oAttrib.StartScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i = 0 + oAttrib.StartScrollPage to oAttrib.FScrollCount + oAttrib.StartScrollPage - 1 %>
		<% if i>oAttrib.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:goPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oAttrib.HasNextScroll then %>
		<a href="javascript:goPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>
</form>
<!-- 목록 끝 -->

<%
	set oAttrib = Nothing
%>
<!-- 표 하단바 끝-->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
