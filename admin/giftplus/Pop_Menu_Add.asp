<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  기프트플러스
' History : 2010.04.02 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/giftplus/giftplus_cls.asp"-->

<%
dim Depth, cdL, cdM, cdS , objView , listtype
	Depth = request("depth")
	cdL= request("cdL")
	cdM= request("cdM")
	cdS= request("cdS")

set objView = new giftManagerView
	objView.getMenuView cdL,cdM,cdS
	
	if cdL <> "" then
	'listtype = objView.listtype
	listtype = getlisttype(cdl)
	end if
	
	if listtype = "" then listtype = "menu"

if listtype = "search" and DEPTH = "S" then		
	response.write "<script>alert('표시형식이 검색형일경우 소카테고리를 추가하실수 없습니다'); self.close();</script>"
	dbget.close() : response.end
end if
%>
<table width="400" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="UpdateFRM" action="Menu_Process.asp" target="">
	<input type="hidden" name="Depth" value="<%= Depth %>">
	<tr>
		<td width="130" bgcolor="#FFFFFF"></td>
		<td bgcolor="#FFFFFF"></td>
	</tr>
<% IF objView.LCode <>"" then %>
	<input type="hidden" name="LCode" size="4" value="<%= objView.LCode %>" />
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">대 카테고리</td>
		<td bgcolor="#FFFFFF"> [<font color="red"><%= objView.LCode %></font>] <%= objView.LCodeNm %>
	</tr>
<% END IF %>

<% IF objView.MCode <>"" then %>
	<input type="hidden" name="MCode" size="4" value="<%= objView.MCode %>" />
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">중 카테고리</td>
		<td bgcolor="#FFFFFF"> [<font color="red"><%= objView.MCode %></font>] <%= objView.MCodeNm %>
	</tr>
<% END IF %>

<% IF objView.SCode <>"" then %>
	<input type="hidden" name="SCode" size="4" value="<%= objView.SCode %>" />
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">소 카테고리</td>
		<td bgcolor="#FFFFFF"> [<font color="red"><%= objView.SCode %></font>]<%= objView.SCodeNm %>
	</tr>
<% END IF %>


<% SELECT CASE DEPTH %>
	<% CASE "L" %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">대 카테고리</td>
		<td bgcolor="#FFFFFF"><input type="text" size="4" name="LCode" value="">(1 ~ 99)</td>
	</tr>	
	<% CASE "M" %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">중 카테고리</td>
		<td bgcolor="#FFFFFF"><input type="text" size="4" name="MCode" value="">(1 ~ 99)</td>
	</tr>		
	<% CASE "S" %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">소 카테고리</td>
		<td bgcolor="#FFFFFF"><input type="text" size="4" name="SCode" value="">(1 ~ 99)</td>
	</tr>		
<% END SELECT %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">순서</td>
		<td bgcolor="#FFFFFF"><input type="text" size="4" name="OrderNo" value="0">(1 ~ 99)</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" align="center">카테고리명</td>
		<td bgcolor="#FFFFFF"><input type="text" size="16" name="CodeNm" value=""></td>
	</tr>
	<% IF DEPTH = "L" THEN %>
	<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">표시형식</td>
	<td bgcolor="#FFFFFF">
		<% drawListType "listtype" , listtype, "" %>
	</td>
	<% END IF %>	
	<tr>
		<td bgcolor="#FFFFFF" colspan="2" align="center"><input type="submit" class="button" value="적용"></td>
	</tr>
	</form>
</table>
<% 
set objView = nothing 
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->