<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/lotteon/lotteonCls.asp"-->
<%
Response.CharSet = "euc-kr"
Dim oLotteon, i, page, std_cat_id
std_cat_id		= requestCheckVar(request("std_cat_id"),30)

If page = ""	Then page = 1
'// ��� ����
Set oLotteon = new CLotteon
 	oLotteon.FPageSize = 1000
 	oLotteon.FCurrPage = page
 	oLotteon.FRectStdCateId = std_cat_id
 	oLotteon.getLotteonDispCateList
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="YELLOW">
	<td></td>
	<td width="20%">����ī�װ�</td>
	<td>����ī�װ���</td>
</tr>
<% If oLotteon.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	Else
		For i = 0 to oLotteon.FresultCount - 1
%>
<tr align="center" height="25" bgcolor="#FFFFFF">
	<td>
		<input type="radio" class="radio" name="disp_cat_id" value="<%= oLotteon.FItemList(i).FDisp_cat_id %>" />
	</td>
	<td><%= oLotteon.FItemList(i).FDisp_cat_id %></td>
	<td align="LEFT">
		<%= oLotteon.FItemList(i).FDisp_cat_nm %>
	</td>
</tr>
<%
		Next
	End If
%>
</table>
<% Set oLotteon = nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
