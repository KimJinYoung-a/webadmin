<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/Compliment_cls.asp"-->
<%
	'// ���� ���� //
	dim cplId
	dim page, searchDiv, searchString, param

	dim oCompliment, i, lp, bgcolor


	'// �Ķ���� ���� //
	cplId = RequestCheckvar(request("cplId"),10)
	page = RequestCheckvar(request("page"),10)
	searchDiv = RequestCheckvar(request("searchDiv"),16)
	searchString = RequestCheckvar(request("searchString"),128)

	if page="" then page=1

	param = "&searchDiv=" & searchDiv & "&searchString=" & server.URLencode(searchString)

	'// Ŭ���� ����
	set oCompliment = new Ccpl
	oCompliment.FCurrPage = page
	oCompliment.FPageSize = 15
	oCompliment.FRectsearchDiv = searchDiv
	oCompliment.FRectsearchString = searchString

	oCompliment.GetcplList
%>
<script language='javascript'>
<!--
	function chk_form()
	{
		var frm = document.frm_search;

		if(!frm.searchString.value)
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
<!-- ��� �˻��� ���� -->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<form name="frm_search" method="POST" action="Compliment_list.asp" onSubmit="return false">
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
		<select name="searchDiv" onchange="chgDiv()">
			<option value="">����</option>
			<%=oCompliment.optCommCd("'E000'", searchDiv)%>
		</select>
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
		<td colspan="5" align="left">�˻��Ǽ� : <%= oCompliment.FTotalCount %> �� Page : <%= page %>/<%= oCompliment.FTotalPage %></td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td align="center" width="40">��ȣ</td>
		<td align="center" width="60">����</td>
		<td align="center">���� ���</td>
		<td align="center" width="60">����</td>
		<td align="center" width="80">�����</td>
	</tr>
	<%
		for lp=0 to oCompliment.FResultCount - 1

			if oCompliment.FcplList(lp).Fisusing="<font color=darkblue>���</font>" then
				bgcolor = "#FFFFFF"
			else
				bgcolor = "#E0E0E0"
			end if
	%>
	<tr align="center" bgcolor="<%=bgcolor%>">
		<td><%= oCompliment.FcplList(lp).FcplId %></td>
		<td align="center"><a href="Compliment_modi.asp?cplId=<%= oCompliment.FcplList(lp).FcplId %>&page=<%=page & param%>&menupos=<%=menupos%>"><%= db2html(oCompliment.FcplList(lp).FcommNm) %></a></td>
		<td align="left"><a href="Compliment_modi.asp?cplId=<%= oCompliment.FcplList(lp).FcplId %>&page=<%=page & param%>&menupos=<%=menupos%>"><%= oCompliment.FcplList(lp).FcplCont %></a></td>
		<td><%= oCompliment.FcplList(lp).Fisusing %></td>
		<td><%= FormatDate(oCompliment.FcplList(lp).Fregdate,"0000.00.00") %></td>
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
					if oCompliment.HasPreScroll then
						Response.Write "<a href='javascript:goPage(" & oCompliment.StarScrollPage-1 & ")'>[pre]</a> &nbsp;"
					else
						Response.Write "[pre] &nbsp;"
					end if
		
					for i=0 + oCompliment.StarScrollPage to oCompliment.FScrollCount + oCompliment.StarScrollPage - 1
		
						if i>oCompliment.FTotalpage then Exit for
		
						if CStr(page)=CStr(i) then
							Response.Write " <font color='red'>[" & i & "]</font> "
						else
							Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
						end if
		
					next
		
					if oCompliment.HasNextScroll then
						Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
					else
						Response.Write "&nbsp; [next]"
					end if
				%>
				<!-- ������ �� -->
				</td>
				<td width="80" align="right">
					<a href="Compliment_write.asp?menupos=<%=menupos%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
				</td>
			</tr>
			</table>
		</td>
	</tr>
</table>
<%
set oCompliment = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->