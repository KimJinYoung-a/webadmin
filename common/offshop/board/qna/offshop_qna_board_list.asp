<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������̿빮��
' Hieditor : 2009.04.07 ������ ����
'			 2011.05.03 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/offshop_function.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/classes/board/offshopqnacls.asp" -->
<%
dim i, j ,shopid, page ,SearchKey, SearchString, param, isNew ,boardqna
	page = Request("page")
	shopid = Request("shopid")
	isNew = Request("isNew")
	SearchKey = Request("SearchKey")
	SearchString = Request("SearchString")
	menupos = Request("menupos")

''���������� �������ΰ�� �ھ� �ִ´�
if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
end if
    

if page="" then page=1
if SearchKey="" then SearchKey="title"
if isNew="" then isNew="Y"

param = "&SearchKey=" & SearchKey & "&SearchString=" & Server.URLencode(SearchString) & "&shopid=" & shopid & "&isNew=" & isNew & "&menupos=" & menupos

'���� 1:1�����亯
set boardqna = New CMyQNA
	boardqna.FPageSize = 20
	boardqna.FCurrPage = page
	boardqna.fSearchNew = isNew
	boardqna.FRectDesigner = shopid
	boardqna.FRectSearchKey = SearchKey
	boardqna.FRectSearchString = SearchString
	boardqna.list()
%>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		ó������
		<select name="isNew">
			<option value="all">��ü</option>
			<option value="Y">��ó��</option>
			<option value="N">ó���Ϸ�</option>
		</select>
		<% if fnChkAuth(session("ssBctDiv"),session("ssBctID"),session("ssBctBigo"))="" then %>
			/ ����
			<% call printOffShopSelectBox(isNew, shopid)%>
		<% end if %>
		/ Ű����
		<select name="SearchKey">
			<option value="title">����</option>
			<option value="userid">�ۼ���ID</option>
			<option value="contents">����</option>
		</select>
		<input type="text" name="SearchString" size="12" value="<%=SearchString%>">

		<script language="javascript">
			document.frm.SearchKey.value="<%=SearchKey%>";
			document.frm.isNew.value="<%=isNew%>";
		</script>	
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">	
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if boardqna.FResultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= boardqna.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= boardqna.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>��(���̵�/�ֹ���ȣ)</td>
    <td>����</td>
    <td>����</td>
    <td>ó������</td>
    <td>�ۼ���</td>
    <td>���</td>
</tr>
<% for i=0 to boardqna.FresultCount-1 %>
<form action="" name="frmBuyPrc<%=i%>" method="get">			

<% if boardqna.FItemList(i).isusing = "Y" then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='#ffffff';>
<% else %>    
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="orange"; onmouseout=this.style.background='#FFFFaa';>
<% end if %>
	<td>&nbsp;<%= printUserId(boardqna.FItemList(i).userid, 2, "*") %><%= boardqna.FItemList(i).orderserial %></td>
	<td align="center"><%= boardqna.FItemList(i).Fshopname %></td>
	<td align="left">&nbsp;<%= db2html(boardqna.FItemList(i).title) %></td>
	<td align="center">
	<%
		if (boardqna.FItemList(i).replyuser=""  or isnull(boardqna.FItemList(i).replyuser)) then
			Response.Write "<font color='darkred'>��ó��</font>"
		else
			Response.Write "<font color='darkblue'>ó���Ϸ�</font>"
		end if
	%>
	</td>
	<td align="center"><%= FormatDate(boardqna.FItemList(i).regdate, "0000-00-00") %></td>
	<td align="center">
		<a href="offshop_qna_board_reply.asp?idx=<%= boardqna.FItemList(i).idx %>&page=<%=page & param%>">��</a>
	</td>
</tr>   
</form>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if boardqna.HasPreScroll then %>
			<span class="list_link"><a href="?page=<%= boardqna.StartScrollPage-1 & param %>">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + boardqna.StartScrollPage to boardqna.StartScrollPage + boardqna.FScrollCount - 1 %>
			<% if (i > boardqna.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(boardqna.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="?page=<%= i & param %>" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if boardqna.HasNextScroll then %>
			<span class="list_link"><a href="?page=<%= i & param%>">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="3" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<%
set boardqna = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
