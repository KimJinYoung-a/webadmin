<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/partner_lecturecls.asp"-->
<%
	'// ���� ���� //
	dim idx
	dim page, searchKey, searchString, searchConfirm, param

	dim oLecture, i, lp, bgcolor, strUsing


	'// �Ķ���� ���� //
	idx = RequestCheckvar(request("idx"),10)
	page = RequestCheckvar(request("page"),10)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = request("searchString")
	searchConfirm = RequestCheckvar(request("searchConfirm"),1)
  	if searchString <> "" then
		if checkNotValidHTML(searchString) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end if

	if page="" then searchConfirm="N"
	if page="" then page=1
	if searchKey="" then searchKey="lecname"

	param = "&searchKey=" & searchKey & "&searchString=" & server.URLencode(searchString) &_
			"&searchConfirm=" & searchConfirm & "&menupos=" & menupos

	'// Ŭ���� ����
	set oLecture = new CPartnerFieldLecture
	oLecture.FCurrPage = page
	oLecture.FPageSize = 20
	oLecture.FRectsearchKey = searchKey
	oLecture.FRectsearchString = searchString
	oLecture.FRectsearchConfirm = searchConfirm

	oLecture.GetPartnerFieldList
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
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<form name="frm_search" method="GET" action="PartnerFieldLecture_list.asp" onSubmit="return chk_form(this)">
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
		�亯����
		<select name="searchConfirm" onchange="document.frm_search.submit()">
			<option value="">::����::</option>
			<option value="Y">�Ϸ�</option>
			<option value="N">���</option>
		</select>
		<script language="javascript">
			document.frm_search.searchConfirm.value="<%=searchConfirm%>";
		</script>
		/ �˻�
		<select name="searchKey">
			<option value="">::����::</option>
			<option value="idx">��ȣ</option>
			<option value="lecname">�����̸�</option>
			<option value="lectitle">���°���</option>
		</select>
		<script language="javascript">
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
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="#F0F0FD">
		<td colspan="9" align="left">�˻��Ǽ� : <%= oLecture.FTotalCount %> �� Page : <%= page %>/<%= oLecture.FTotalPage %></td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td align="center" width="40">��ȣ</td>
		<td align="center" width="60">�����о�</td>
		<td align="center">���°���</td>
		<td align="center" width="70">�����</td>
		<td align="center" width="60">�������</td>
		<td align="center" width="110">����ó</td>
		<td align="center" width="110">�޴���</td>
		<td align="center" width="75">�����</td>
		<td align="center" width="40">�亯</td>
	</tr>
	<%
		for lp=0 to oLecture.FResultCount - 1

			'������������� ���� �� ���¸� ����
			if oLecture.FItemList(lp).Fconfirmyn="N" then
				bgcolor="#FFFFFF"
				strUsing = "<font color=darkred>���</font>"
			else
				bgcolor="#F8F8F8"
				strUsing = "<font color=darkblue>�Ϸ�</font>"
			end if
	%>
	<tr align="center" bgcolor="<%=bgcolor%>">
		<td><%= oLecture.FItemList(lp).Fidx %></td>
		<td><%= oLecture.FItemList(lp).Flecarea %></td>
		<td align="left"><a href="/academy/Partnership/PartnerFieldLecture_view.asp?idx=<%= oLecture.FItemList(lp).Fidx %>&page=<%=page & param%>"><%= oLecture.FItemList(lp).Flectitle %></a></td>
		<td><a href="/academy/Partnership/PartnerFieldLecture_view.asp?idx=<%= oLecture.FItemList(lp).Fidx %>&page=<%=page & param%>"><%= oLecture.FItemList(lp).Flecname %></a></td>
		<td><%= oLecture.FItemList(lp).Flecbirthday %></td>
		<td><%= oLecture.FItemList(lp).Flectel %></td>
		<td><%= oLecture.FItemList(lp).Flechp %></td>
		<td><%= FormatDate(oLecture.FItemList(lp).Fregdate,"0000.00.00") %></td>
		<td align="center"><%=strUsing%></td>
	</tr>
	<%
		next
	%>
	<tr bgcolor="#FFFFFF">
		<td colspan="9" height="30" align="center">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td align="center" class="a">
				<!-- ������ ���� -->
				<%
					if oLecture.HasPreScroll then
						Response.Write "<a href='javascript:goPage(" & oLecture.StarScrollPage-1 & ")'>[pre]</a> &nbsp;"
					else
						Response.Write "[pre] &nbsp;"
					end if
		
					for i=0 + oLecture.StarScrollPage to oLecture.FScrollCount + oLecture.StarScrollPage - 1
		
						if i>oLecture.FTotalpage then Exit for
		
						if CStr(page)=CStr(i) then
							Response.Write " <font color='red'>[" & i & "]</font> "
						else
							Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
						end if
		
					next
		
					if oLecture.HasNextScroll then
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
</table>
<%
set oLecture = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->