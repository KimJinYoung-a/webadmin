<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/board_cls.asp"-->
<%
	'// ���� ���� //
	dim brdId
	dim page, searchDiv, searchKey, searchString, param

	dim oBoard, i, lp, bgcolor, strUsing


	'// �Ķ���� ���� //
	brdId = requestCheckVar(request("brdId"),10)
	page = requestCheckVar(request("page"),10)
	searchDiv = requestCheckVar(request("searchDiv"),10)
	searchKey = requestCheckVar(request("searchKey"),10)
	searchString = requestCheckVar(request("searchString"),128)
  	if searchString <> "" then
		if checkNotValidHTML(searchString) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
		response.write "</script>"
		response.End
		end if
	end if

	if page="" then page=1
	if searchKey="" then searchKey="qstTitle"

	param = "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString & "&menupos=" & menupos

	'// Ŭ���� ����
	set oBoard = new Cboard
	oBoard.FCurrPage = page
	oBoard.FPageSize = 20
	oBoard.FRectsearchDiv = searchDiv
	oBoard.FRectsearchKey = searchKey
	oBoard.FRectsearchString = searchString
	oBoard.FRectuserid = Session("ssBctId")

	oBoard.GetBoardList
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
<!-- ��� �˻��� ���� -->
<table width="750" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<form name="frm_search" method="POST" action="board_list.asp" onSubmit="return false">
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
		<select name="searchDiv" onchange="chgDiv()">
		<option value="">����</option>
		<%=oBoard.optCommCd("'G000'", searchDiv)%>
		</select>
		/ �˻�
		<select name="searchKey">
			<option value="">����</option>
			<option value="brdId">��ȣ</option>
			<option value="qstTitle">����</option>
			<option value="qstCont">����</option>
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
<table width="750" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr align="center" bgcolor="#F0F0FD">
		<td colspan="5" align="left">�˻��Ǽ� : <%= oBoard.FTotalCount %> �� Page : <%= page %>/<%= oBoard.FTotalPage %></td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td align="center" width="40">��ȣ</td>
		<td align="center" width="70">����</td>
		<td align="center" width="490">����</td>
		<td align="center" width="70">����</td>
		<td align="center" width="80">�����</td>
	</tr>
	<%
		for lp=0 to oBoard.FResultCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= oBoard.FBoardList(lp).FbrdId %></td>
		<td><%= oBoard.FBoardList(lp).FcommNm %></td>
		<td align="left"><a href="board_view.asp?brdId=<%= oBoard.FBoardList(lp).FbrdId %>&page=<%=page & param%>"><%= db2html(oBoard.FBoardList(lp).FqstTitle) %></a></td>
		<td><%= oBoard.FBoardList(lp).Fisanswer %></td>
		<td><%= FormatDate(oBoard.FBoardList(lp).Fregdate,"0000.00.00") %></td>
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
					if oBoard.HasPreScroll then
						Response.Write "<a href='javascript:goPage(" & oBoard.StarScrollPage-1 & ")'>[pre]</a> &nbsp;"
					else
						Response.Write "[pre] &nbsp;"
					end if
		
					for i=0 + oBoard.StarScrollPage to oBoard.FScrollCount + oBoard.StarScrollPage - 1
		
						if i>oBoard.FTotalpage then Exit for
		
						if CStr(page)=CStr(i) then
							Response.Write " <font color='red'>[" & i & "]</font> "
						else
							Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
						end if
		
					next
		
					if oBoard.HasNextScroll then
						Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
					else
						Response.Write "&nbsp; [next]"
					end if
				%>
				<!-- ������ �� -->
				</td>
				<td width="80" align="right">
					<a href="board_write.asp?menupos=<%=menupos%>"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
				</td>
			</tr>
			</table>
		</td>
	</tr>
</table>
<%
set oBoard = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->