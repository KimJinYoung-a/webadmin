<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/PositInfoCls.asp" -->
<%
	Dim page, SearchKey, SearchString

	page = Request("page")
	SearchKey = Request("SearchKey")
	SearchString = Request("SearchString")
	if page="" then page=1


	'// ���� ����
	dim oPosit, lp
	Set oPosit = new CPosit

	oPosit.FPagesize = 15
	oPosit.FCurrPage = page
	oPosit.FRectsearchKey = searchKey
	oPosit.FRectsearchString = searchString
	
	oPosit.GetPositList
%>
<!-- �˻� ���� -->
<script language="javascript">
<!--
	// �ű� ���� ���
	function AddItem()
	{
		window.open("pop_positInfo_add.asp","popAddIem","width=378,height=410,scrollbars=yes");
	}

	// ���� ����/����
	function ModiItem(psn)
	{
		window.open("pop_positInfo_add.asp?posit_sn="+psn,"popModiIem","width=360,height=200,scrollbars=no");
	}

	// ������ �̵�
	function goPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.submit();
	}
//-->
</script>
<table width="750" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="25" valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td>&nbsp;</td>
	<td align="right">
		<select name="SearchKey">
			<option value="">::����::</option>
			<option value="posit_sn">���޹�ȣ</option>
			<option value="posit_name">���޸�</option>
		</select>
		<script language="javascript">document.frm.SearchKey.value="<%=SearchKey%>";</script>
		<input type="text" name="SearchString" size="12" value="<%=SearchString%>">
		<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0" align="absmiddle">
	</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<!-- ��� �� ���� -->
<table width="750" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr><td height="1" colspan="15" bgcolor="#BABABA"></td></tr>
<tr height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="right">
		<table width="100%" border=0 cellspacing=0 cellpadding=0 class="a">
		<tr>
			<td>�� <%=oPosit.FtotalCount%> �� ����</td>
			<td align="right">page : <%= page %>/<%=oPosit.FtotalPage%></td>
		</tr>
		</table>
	</td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ��� �� �� -->
<!-- ���� ��� ���� -->
<table width="750" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#E6E6E6">
	<td width="60">���޹�ȣ</td>
	<td>���޸�</td>
	<td width="100">��������</td>
</tr>
<%
	if oPosit.FResultCount=0 then
%>
<tr>
	<td colspan="4" height="60" align="center" bgcolor="#FFFFFF">���(�˻�)�� ������ �����ϴ�.</td>
</tr>
<%
	else
		for lp=0 to oPosit.FResultCount - 1
%>
<tr align="center" bgcolor="<% if oPosit.FitemList(lp).Fposit_isDel="N" then Response.Write "#FFFFFF": else Response.Write "#F0F0F0": end if %>">
	<td><%=oPosit.FitemList(lp).Fposit_sn%></td>
	<td align="left"><a href="javascript:ModiItem(<%=oPosit.FitemList(lp).Fposit_sn%>)"><%=oPosit.FitemList(lp).Fposit_name%></a></td>
	<td><%=oPosit.FitemList(lp).Fposit_isDel%></td>
</tr>	
<%
		next
	end if
%>
</table>
<!-- ���� ��� �� -->
<!-- ������ ���� -->
<table width="750" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr valign="bottom" height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr valign="bottom">
			<td align="center">
			<!-- ������ ���� -->
			<%
				if oPosit.HasPreScroll then
					Response.Write "<a href='javascript:goPage(" & oPosit.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
				else
					Response.Write "[pre] &nbsp;"
				end if

				for lp=0 + oPosit.StartScrollPage to oPosit.FScrollCount + oPosit.StartScrollPage - 1

					if lp>oPosit.FTotalpage then Exit for
	
					if CStr(page)=CStr(lp) then
						Response.Write " <font color='red'>[" & lp & "]</font> "
					else
						Response.Write " <a href='javascript:goPage(" & lp & ")'>[" & lp & "]</a> "
					end if

				next

				if oPosit.HasNextScroll then
					Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
				else
					Response.Write "&nbsp; [next]"
				end if
			%>
			<!-- ������ �� -->
			</td>
			<td width="75" align="right"><a href="javascript:AddItem('');"><img src="/images/icon_new_registration.gif" width="75" border="0"></a></td>
		</tr>
		</table>
	</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top" height="10">
	<td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- ������ �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->