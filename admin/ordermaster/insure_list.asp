<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/order/myorder_insurecls.asp"-->
<%
	'// ���� ���� //
	dim InsureIdx
	dim page, searchDiv, searchKey, searchString, param

	dim oInsure, i, lp, bgcolor, strIsue, strConfirm


	'// �Ķ���� ���� //
	InsureIdx = request("InsureIdx")
	page = request("page")
	searchDiv = request("searchDiv")
	searchKey = request("searchKey")
	searchString = request("searchString")

	if page="" then
		page=1
		searchDiv = "Y"
	end if
	if searchKey="" then searchKey="orderserial"

	param = "&menupos=" & menupos & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString

	'// Ŭ���� ����
	set oInsure = new CInsure
	oInsure.FCurrPage = page
	oInsure.FPageSize = 20
	oInsure.FRectsearchDiv = searchDiv
	oInsure.FRectsearchKey = searchKey
	oInsure.FRectsearchString = searchString

	oInsure.GetInsureList
%>
<script language='javascript'>
<!--
	function chk_form()
	{
		var frm = document.frm_search;

		if(!frm.searchKey.value)
		{
			alert("�˻� ������ �������ֽʽÿ�.");
			frm.searchKey.focus();
			return;
		}
		else if(!frm.searchString.value)
		{
			alert("�˻�� �Է����ֽʽÿ�.");
			frm.searchString.focus();
			return;
		}

		frm.submit();
	}

	function goPage(pg)
	{
		var frm = document.frm_search;

		frm.page.value= pg;
		frm.submit();
	}

	function chgDiv()
	{
		var frm = document.frm_search;
		frm.submit();
	}

//-->
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm_search" method="GET" action="Insure_list.asp" onSubmit="return false">
	<input type="hidden" name="page" value="<%=page%>">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			<a href="http://www.usafe.co.kr" target="_blank"><img src="/admin/images/link_usafe.gif" border="0" alt="U-Safe�� �̵�"></a>
			&nbsp;
			�߱޿���:
			<select class="select" name="searchDiv" onchange="chgDiv()">
				<option value="">����</option>
				<option value="Y">����</option>
				<option value="N">����</option>
			</select>
			&nbsp;
			�˻�����:
			<select class="select" name="searchKey">
				<option value="">����</option>
				<option value="InsureIdx">��ȣ</option>
				<option value="orderserial">�ֹ���ȣ</option>
				<option value="buyname">�������̸�</option>
			</select>
			<script language="javascript">
				document.frm_search.searchDiv.value="<%=searchDiv%>";
				document.frm_search.searchKey.value="<%=searchKey%>";
			</script>
			<input type="text" class="text" name="searchString" size="20" value="<%= searchString %>">
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="chk_form()">
		</td>
	</tr>
	</form>
</table>

<p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm_list" method="Post" action="">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= oInsure.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %> / <%= oInsure.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" width="80">�ֹ���ȣ</td>
		<td align="center">��ǰ��</td>
		<td align="center" width="70">���űݾ�</td>
		<td align="center" width="60">������</td>
		<td align="center" width="70">������</td>
		<td align="center" width="50">�ֹ�����</td>
		<td align="center" width="150" colspan="2">������</td>
	</tr>
	<%
		for lp=0 to oInsure.FResultCount - 1
			'������
			if oInsure.FInsureList(lp).FinsureCd="0" then
				strIsue = "<font color=darkblue>����</font>"
			else
				strIsue = "<font color=darkred>����</font>"
			end if
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><a href="Insure_view.asp?OrderIdx=<%= oInsure.FInsureList(lp).FOrderIdx %>&page=<%=page & param%>"><%= oInsure.FInsureList(lp).Forderserial %></a></td>
		<td align="left"><a href="Insure_view.asp?OrderIdx=<%= oInsure.FInsureList(lp).FOrderIdx %>&page=<%=page & param%>"><%= db2html(oInsure.FInsureList(lp).Fitemname) %></a></td>
		<td><%= CurrFormat(oInsure.FInsureList(lp).FsubtotalPrice) %></td>
		<td><%= oInsure.FInsureList(lp).Fbuyname %></td>
		<td><%= FormatDate(oInsure.FInsureList(lp).Fregdate,"0000.00.00") %></td>
		<td><%= NormalIpkumDivName(oInsure.FInsureList(lp).Fipkumdiv) %></td>
		<td><%= strIsue %></td>
		<td><%= oInsure.FInsureList(lp).FinsureMsg %></td>
	</tr>
	<%
		next
	%>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td align="center" class="a">
				<!-- ������ ���� -->
				<%
					if oInsure.HasPreScroll then
						Response.Write "<a href='javascript:goPage(" & oInsure.StarScrollPage-1 & ")'>[pre]</a> &nbsp;"
					else
						Response.Write "[pre] &nbsp;"
					end if
		
					for i=0 + oInsure.StarScrollPage to oInsure.FScrollCount + oInsure.StarScrollPage - 1
		
						if i>oInsure.FTotalpage then Exit for
		
						if CStr(page)=CStr(i) then
							Response.Write " <font color='red'>[" & i & "]</font> "
						else
							Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
						end if
		
					next
		
					if oInsure.HasNextScroll then
						Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
					else
						Response.Write "&nbsp; [next]"
					end if
				%>
				<!-- ������ �� -->
				</td>
			</tr>
			</table>
		</td>
	</tr>
</form>
</table>
<%
set oInsure = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->