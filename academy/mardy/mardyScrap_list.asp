<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/mardy_Scrapcls.asp"-->
<%
	'// ���� ���� //
	dim ScrapId
	dim page, searchKey, searchString, param

	dim oScrap, i, lp, bgcolor, strUsing


	'// �Ķ���� ���� //
	ScrapId = RequestCheckvar(request("ScrapId"),10)
	page = RequestCheckvar(request("page"),10)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = request("searchString")
  	if searchString <> "" then
		if checkNotValidHTML(searchString) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end if

	if page="" then page=1
	if searchKey="" then searchKey="title"

	param = "&searchKey=" & searchKey & "&searchString=" & searchString

	'// Ŭ���� ����
	set oScrap = new CMardyScrap
	oScrap.FCurrPage = page
	oScrap.FPageSize = 20
	oScrap.FRectsearchKey = searchKey
	oScrap.FRectsearchString = searchString

	oScrap.GetMardyScrapList
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
//-->
</script>
<!-- ��� �˻��� ���� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<form name="frm_search" method="POST" action="mardyScrap_list.asp" onSubmit="return false">
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
		<select name="searchKey">
			<option value="">����</option>
			<option value="ScrapId">�̾߱��ȣ</option>
			<option value="title">����</option>
			<option value="contents">����</option>
		</select>
		<script language="javascript">
			document.frm_search.searchKey.value="<%=searchKey%>";
		</script>
		<input type="text" name="searchString" size="20" value="<%= searchString %>">
       	<img src="/admin/images/search2.gif" onClick="chk_form()" style="width:74px;height:22px;border:0px;cursor:pointer" align="absmiddle">
	</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</form>
</table>
<!-- ��� �˻��� �� -->
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="#F0F0FD">
		<td colspan="7" align="left">�˻��Ǽ� : <%= oScrap.FTotalCount %> �� Page : <%= page %>/<%= oScrap.FTotalPage %></td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td align="center" width="40">��ȣ</td>
		<td align="center" width="52">�̹���</td>
		<td align="center">����</td>
		<td align="center" width="70">�����</td>
		<td align="center" width="50">��ȸ��</td>
		<td align="center" width="80">�����</td>
		<td align="center" width="40">����</td>
	</tr>
	<%
		for lp=0 to oScrap.FResultCount - 1

			'������������� ���� �� ���¸� ����
			if oScrap.FItemList(lp).Fisusing="N" then
				bgcolor="#E0E0E0"
				strUsing = "<font color=darkred>����</font>"
			else
				bgcolor="#FFFFFF"
				strUsing = "<font color=darkblue>���</font>"
			end if
	%>
	<tr align="center" bgcolor="<%=bgcolor%>">
		<td><%= oScrap.FItemList(lp).FScrapId %></td>
		<td><img src="<%= oScrap.FItemList(lp).FimgIcon_full %>" width="50"></td>
		<td align="left"><a href="/academy/mardy/mardyScrap_subList.asp?menupos=<%=menupos%>&ScrapId=<%= oScrap.FItemList(lp).FScrapId %>&page=<%=page & param%>"><%= oScrap.FItemList(lp).Ftitle %></a></td>
		<td><%= oScrap.FItemList(lp).Fusername %></td>
		<td align="center"><%= oScrap.FItemList(lp).FhitCount %></td>
		<td><%= FormatDate(oScrap.FItemList(lp).Fregdate,"0000.00.00") %></td>
		<td align="center"><%= strUsing%></td>
	</tr>
	<%
		next
	%>
	<tr bgcolor="#FFFFFF">
		<td colspan="7" height="30" align="center">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td align="center" class="a">
				<!-- ������ ���� -->
				<%
					if oScrap.HasPreScroll then
						Response.Write "<a href='javascript:goPage(" & oScrap.StarScrollPage-1 & ")'>[pre]</a> &nbsp;"
					else
						Response.Write "[pre] &nbsp;"
					end if
		
					for i=0 + oScrap.StarScrollPage to oScrap.FScrollCount + oScrap.StarScrollPage - 1
		
						if i>oScrap.FTotalpage then Exit for
		
						if CStr(page)=CStr(i) then
							Response.Write " <font color='red'>[" & i & "]</font> "
						else
							Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
						end if
		
					next
		
					if oScrap.HasNextScroll then
						Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
					else
						Response.Write "&nbsp; [next]"
					end if
				%>
				<!-- ������ �� -->
				</td>
				<td width="80" align="right">
					<a href="mardyScrap_write.asp?menupos=<%=menupos%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
				</td>
			</tr>
			</table>
		</td>
	</tr>
</table>
<%
set oScrap = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->