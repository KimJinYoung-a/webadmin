<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : RelateKeywordLink_List.asp
' Discription : ī�װ� ���� Ű���� ���
' History : 2008.03.28 ������ ����
'			2022.07.05 �ѿ�� ����(isms�������ġ, ǥ���ڵ����κ���)
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/CategoryCls.asp"-->
<%
	Dim page, SearchKey, SearchString

	page = Request("page")
	SearchKey = Request("SearchKey")
	SearchString = Request("SearchString")
	if page="" then	page=1


	'// ���� ����
	dim oRelate, lp
	Set oRelate = new CRelateList

	oRelate.FPagesize = 15
	oRelate.FCurrPage = page
	oRelate.FRectCDL = request("cdl")
	oRelate.FRectCDM = request("cdm")
	oRelate.FRectCDS = request("cds")
	oRelate.FRectsearchKey = searchKey
	oRelate.FRectsearchString = searchString
	
	oRelate.GetRelateLinkList
%>
<!-- �˻� ���� -->
<script type='text/javascript'>
<!--
	// ������ �̵�
	function goPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.action="RelateKeywordLink_list.asp";
		document.frm.submit();
	}

	// ������ ������(����) ������ �̵�
	function goEdit(rid)
	{
		document.frm.rid.value=rid;
		document.frm.page.value='<%= page %>';
		document.frm.action="RelateKeywordLink_Edit.asp";
		document.frm.submit();
	}

	// ������ ���� ����
	function goDel(rid)
	{
		if(confirm("[" + rid + "]�� ����Ű���带 �����Ͻðڽ��ϱ�?\n\n�� �Ϸ�� ������ �����Ǹ� ������ �� �����ϴ�.")) {
			document.frm.rid.value=rid;
			document.frm.mode.value="delete";
			document.frm.action="DoRelate_Process.asp";
			document.frm.submit();
		}
	}

	// �űԵ�� �������� �̵�
	function goAddItem()  {
		self.location="RelateKeywordLink_Edit.asp?menupos=<%=menupos%>";
	}
//-->
</script>
<form name="frm" method="get" action="" action="RelateKeywordLink_list.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="rid" value="">
<input type="hidden" name="mode" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td>
			<table width="100%" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td><!-- #include virtual="/common/module/categoryselectbox.asp"--></td>
				<td align="right">
					Ű���� :
					<select class="select" name="SearchKey">
						<option value="">::����::</option>
						<option value="linkCode">��ũ�ڵ�</option>
						<option value="linkKeyword">Ű����</option>
						<option value="linkURL">��ũ</option>
					</select>
					<input type="text" class="text" name="SearchString" size="20" value="<%=SearchString%>">
					<script language="javascript">
						document.frm.SearchKey.value="<%=SearchKey%>";
					</script>
				</td>
			</tr>
			</table>
		</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="submit" class="button_s" value="�˻�">
		</td>
	</tr>
</table>
</form>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td align="right">
		<input type="button" class="button" value="�űԵ��" onClick="goAddItem()">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="5">
		�˻���� : <b><%=oRelate.FtotalCount%></b>
		&nbsp;
		������ : <b><%= page %> / <%=oRelate.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�ڵ�</td>
	<td>ī�װ�</td>
	<td>Ű����</td>
	<td>��ũ</td>
	<td>����/����</td>
</tr>
<%
	if oRelate.FResultCount=0 then
%>
<tr>
	<td colspan="5" height="60" align="center" bgcolor="#FFFFFF">���(�˻�)�� �������� �����ϴ�.</td>
</tr>
<%
	else
		for lp=0 to oRelate.FResultCount - 1
%>
<tr align="center" bgcolor="#FFFFFF">
	<td><%=oRelate.FitemList(lp).Flinkcode%></td>
	<td align="left">
	<%
		Response.Write oRelate.FitemList(lp).FCDL_nm
		if Not(oRelate.FitemList(lp).FCDM_nm="" or isNull(oRelate.FitemList(lp).FCDM_nm)) then Response.Write " > " & oRelate.FitemList(lp).FCDM_nm
		if Not(oRelate.FitemList(lp).FCDS_nm="" or isNull(oRelate.FitemList(lp).FCDS_nm)) then Response.Write " > " & oRelate.FitemList(lp).FCDS_nm
	%>
	</td>
	<td align="left"><%= ReplaceBracket(oRelate.FitemList(lp).FlinkKeyword) %></td>
	<td align="left"><%= ReplaceBracket(oRelate.FitemList(lp).FlinkURL) %></td>
	<td>
		<input type="button" value="����" class="button" onClick="goEdit(<%=oRelate.FitemList(lp).Flinkcode%>)">
		<input type="button" value="����" class="button" onClick="goDel(<%=oRelate.FitemList(lp).Flinkcode%>)">
	</td>
</tr>	
<%
		next
	end if
%>
<!-- ���� ��� �� -->
<!-- ������ ���� -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
	<!-- ������ ���� -->
	<%
		if oRelate.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & oRelate.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + oRelate.StartScrollPage to oRelate.FScrollCount + oRelate.StartScrollPage - 1

			if lp>oRelate.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>" & lp & "</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>" & lp & "</a> "
			end if

		next

		if oRelate.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
	%>
	<!-- ������ �� -->
	</td>
</tr>
</table>
<!-- ������ �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->