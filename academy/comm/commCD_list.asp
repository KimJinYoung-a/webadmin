<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/commCd_cls.asp"-->
<%
	'// ���� ���� //
	dim CommCd
	dim page, searchDiv, searchKey, searchString, param, isusing

	dim oComm, i, lp, bgcolor, strUsing


	'// �Ķ���� ���� //
	CommCd = RequestCheckvar(request("CommCd"),10)
	page = RequestCheckvar(request("page"),10)
	searchDiv = RequestCheckvar(request("searchDiv"),16)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = RequestCheckvar(request("searchString"),128)
	isusing = RequestCheckvar(request("isusing"),2)

	if page="" then page=1
	if searchKey="" then searchKey="commNm"

	param = "&menupos=" & menupos & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey &_
			"&searchString=" & server.URLencode(searchString) & "&isusing=" & isusing

	'// Ŭ���� ����
	set oComm = new CComm
	oComm.FCurrPage = page
	oComm.FPageSize = 20
	oComm.FRectsearchDiv = searchDiv
	oComm.FRectsearchKey = searchKey
	oComm.FRectsearchString = searchString
	oComm.FRectisusing = isusing

	oComm.GetCommList
%>
<script language='javascript'>
<!--
	function chk_form(frm)
	{
		if(!frm.searchKey.value)
		{
			alert("�˻� ������ �������ֽʽÿ�.");
			frm.searchKey.focus();
			return false;
		}
		else if(!frm.searchString.value)
		{
			alert("�˻�� �Է����ֽʽÿ�.");
			frm.searchString.focus();
			return false;
		}

		frm.submit();
	}

	function goPage(pg)
	{
		var frm = document.frm_search;

		frm.page.value= pg;
		frm.submit();
	}
//-->
</script>
<!-- ��� �˻��� ���� -->
<table width="750" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<form name="frm_search" method="GET" action="CommCd_list.asp" onSubmit="return chk_form(this)">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="30">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td valign="top" align="right">
		����
		<select name="isusing" onChange="goPage(frm_search.page.value)">
			<option value="">��ü</option>
			<option value="Y">���</option>
			<option value="N">����</option>
		</select>
		/ �׷�
		<select name="searchDiv" onChange="goPage(frm_search.page.value)">
			<option value="">��ü</option>
			<%= oComm.optGroupCd(searchDiv)%>
		</select>
		/ �˻�
		<select name="searchKey">
			<option value="commCd">�����ڵ�</option>
			<option value="commNm">�ڵ��</option>
		</select>
		<script language="javascript">
			document.frm_search.isusing.value="<%=isusing%>";
			document.frm_search.searchKey.value="<%=searchKey%>";
		</script>
		<input type="text" name="searchString" size="20" value="<%= searchString %>">
       	<input type="image" src="/admin/images/search2.gif" style="width:74px;height:22px;border:0px;cursor:pointer" align="absmiddle">
	</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</form>
</table>
<!-- ��� �˻��� �� -->
<table width="750" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="#F0F0FD">
		<td colspan="5" align="left">�˻��Ǽ� : <%= oComm.FTotalCount %> �� Page : <%= page %>/<%= oComm.FTotalPage %></td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td align="center" width="40">��ȣ</td>
		<td align="center" width="140">�׷�</td>
		<td align="center" width="80">�����ڵ�</td>
		<td align="center">�ڵ��</td>
		<td align="center" width="50">����</td>
	</tr>
	<%
		for lp=0 to oComm.FResultCount - 1

			if oComm.FCommList(lp).Fisusing="<font color=darkblue>���</font>" then
				bgcolor = "#FFFFFF"
			else
				bgcolor = "#E0E0E0"
			end if
	%>
	<tr align="center" bgcolor="<%=bgcolor%>">
		<td><%= lp + (page * oComm.FPageSize)-oComm.FPageSize+1 %></td>
		<td><%= oComm.FCommList(lp).FgroupNm %></td>
		<td><%= oComm.FCommList(lp).FCommCd %></td>
		<td align="left"><a href="CommCd_modi.asp?CommCd=<%= oComm.FCommList(lp).FCommCd %>&page=<%=page & param%>"><%= db2html(oComm.FCommList(lp).FcommNm) %></a></td>
		<td><%= oComm.FCommList(lp).Fisusing %></td>
	</tr>
	<%
		next
	%>
	<tr bgcolor="#FFFFFF">
		<td colspan="5" height="30" align="center">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td align="center" class="a">
				<!-- ������ ���� -->
				<%
					if oComm.HasPreScroll then
						Response.Write "<a href='javascript:goPage(" & oComm.StarScrollPage-1 & ")'>[pre]</a> &nbsp;"
					else
						Response.Write "[pre] &nbsp;"
					end if
		
					for i=0 + oComm.StarScrollPage to oComm.FScrollCount + oComm.StarScrollPage - 1
		
						if i>oComm.FTotalpage then Exit for
		
						if CStr(page)=CStr(i) then
							Response.Write " <font color='red'>[" & i & "]</font> "
						else
							Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
						end if
		
					next
		
					if oComm.HasNextScroll then
						Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
					else
						Response.Write "&nbsp; [next]"
					end if
				%>
				<!-- ������ �� -->
				</td>
				<td width="77" align="right"><a href="commCd_write.asp?page=<%=param%>"><img src="/images/icon_new_registration.gif" width="75" height="20" border="0"></a></td>
			</tr>
			</table>
		</td>
	</tr>
</table>
<%
set oComm = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->