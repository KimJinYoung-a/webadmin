<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/faq_cls.asp"-->
<% 
	'// ���� ���� //
	dim faqid
	dim page, searchDiv, searchKey, searchString, param

	dim ofaq, i, lp, bgcolor, strUsing


	'// �Ķ���� ���� //
	faqid = RequestCheckvar(request("faqid"),10)
	page = RequestCheckvar(request("page"),10)
	searchDiv = RequestCheckvar(request("searchDiv"),16)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = RequestCheckvar(request("searchString"),128)

	if page="" then page=1
	if searchKey="" then searchKey="title"

	param = "&searchKey=" & searchKey & "&searchString=" & searchString

	'// Ŭ���� ����
	set ofaq = new Cfaq
	ofaq.FCurrPage = page
	ofaq.FPageSize = 20
	ofaq.FRectsearchDiv = searchDiv
	ofaq.FRectsearchKey = searchKey
	ofaq.FRectsearchString = searchString

	ofaq.GetFAQList
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
<form name="frm_search" method="POST" action="faq_list.asp" onSubmit="return false">
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
		<select name="searchDiv" onChange="goPage(frm_search.page.value)">
			<option value="">����</option>
			<%= ofaq.optCommCd("B000", searchDiv)%>
		</select>
		/ �˻�
		<select name="searchKey">
			<option value="">����</option>
			<option value="faqid">������ȣ</option>
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
		<td colspan="6" align="left">�˻��Ǽ� : <%= ofaq.FTotalCount %> �� Page : <%= page %>/<%= ofaq.FTotalPage %></td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td align="center" width="40">��ȣ</td>
		<td align="center" width="120">����</td>
		<td align="center">����</td>
		<td align="center" width="70">�����</td>
		<td align="center" width="50">��ȸ��</td>
		<td align="center" width="80">�����</td>
	</tr>
	<%
		for lp=0 to ofaq.FResultCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= ofaq.FfaqList(lp).Ffaqid %></td>
		<td><%= ofaq.FfaqList(lp).FcommNm %></td>
		<td align="left"><a href="faq_view.asp?faqid=<%= ofaq.FfaqList(lp).Ffaqid %>&page=<%=page & param%>"><%= db2html(ofaq.FfaqList(lp).Ftitle) %></a></td>
		<td><%= ofaq.FfaqList(lp).Fusername %></td>
		<td><%= ofaq.FfaqList(lp).FhitCount %></td>
		<td><%= FormatDate(ofaq.FfaqList(lp).Fregdate,"0000.00.00") %></td>
	</tr>
	<%
		next
	%>
	<tr bgcolor="#FFFFFF">
		<td colspan="6" height="30" align="center">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td align="center" class="a">
				<!-- ������ ���� -->
				<%
					if ofaq.HasPreScroll then
						Response.Write "<a href='javascript:goPage(" & ofaq.StarScrollPage-1 & ")'>[pre]</a> &nbsp;"
					else
						Response.Write "[pre] &nbsp;"
					end if
		
					for i=0 + ofaq.StarScrollPage to ofaq.FScrollCount + ofaq.StarScrollPage - 1
		
						if i>ofaq.FTotalpage then Exit for
		
						if CStr(page)=CStr(i) then
							Response.Write " <font color='red'>[" & i & "]</font> "
						else
							Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
						end if
		
					next
		
					if ofaq.HasNextScroll then
						Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
					else
						Response.Write "&nbsp; [next]"
					end if
				%>
				<!-- ������ �� -->
				</td>
				<td width="80" align="right">
					<a href="faq_write.asp?menupos=<%=menupos%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
				</td>
			</tr>
			</table>
		</td>
	</tr>
</table>
<%
set ofaq = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
