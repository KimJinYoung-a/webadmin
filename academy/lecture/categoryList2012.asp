<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/categoryCls.asp"-->
<%
	'// ���� ���� //
	dim largeCate, CateDiv
	dim page, searchKey, searchString, param, isusing

	dim oCate, i, lp, bgcolor, strUsing


	'// �Ķ���� ���� //
	page = RequestCheckvar(request("page"),10)
	CateDiv = RequestCheckvar(request("CateDiv"),16)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = request("searchString")
	isusing = RequestCheckvar(request("isusing"),1)
  	if searchString <> "" then
		if checkNotValidHTML(searchString) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
	searchKey = ""

	if page="" then page=1
	if CateDiv="" then CateDiv="code_large"
	if searchKey="" then searchKey=CateDiv

	param = "&menupos=" & menupos & "&CateDiv=" & CateDiv & "&searchKey=" & searchKey &_
			"&searchString=" & server.URLencode(searchString) & "&isusing=" & isusing

	'// Ŭ���� ����
	set oCate = new CCate
	oCate.FCateDiv = CateDiv
	oCate.FCurrPage = page
	oCate.FPageSize = 20
	oCate.FRectsearchKey = searchKey
	oCate.FRectsearchString = searchString
	oCate.FRectisusing = isusing

	If CateDiv = "code_large" Then
		oCate.GetLargeCateList		
	else
		oCate.GetMidCateList
	End If 
	
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
<form name="frm_search" method="GET" action="categoryList2012.asp" onSubmit="return chk_form(this)">
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
		ī�װ�
		<select name="CateDiv" onChange="goPage(frm_search.page.value)">
			<option value="code_large" <%=chkiif(CateDiv="code_large","selected","")%>>��ī�װ�</option>
			<option value="code_mid" <%=chkiif(CateDiv="code_mid","selected","")%>>��ī�װ�</option>
		</select>
		/ ����
		<select name="isusing" onChange="goPage(frm_search.page.value)">
			<option value="">��ü</option>
			<option value="Y">���</option>
			<option value="N">����</option>
		</select>
		/ �˻� 
		<select name="searchKey">
			<option value="<%=CateDiv%>" <%=chkiif(searchKey=CateDiv Or IsNull(searchKey),"selected","")%>>ī�װ��ڵ�</option>
			<option value="code_nm" <%=chkiif(searchKey=CateDiv,"selected","")%>>�ڵ��</option>
		</select>
		<script language="javascript">
			document.frm_search.CateDiv.value="<%=CateDiv%>";
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
		<td colspan="8" align="left">�˻��Ǽ� : <%= oCate.FTotalCount %> �� Page : <%= page %>/<%= oCate.FTotalPage %></td>
	</tr>
	<tr align="center" bgcolor="#DDDDFF">
		<td align="center" width="40">��ȣ</td>
		<% if CateDiv="code_mid" then %>
		<td align="center" width="80">��ī���ڵ�</td>
		<td align="center" width="80">��ī���ڵ��</td>
		<% End If %>
		<td align="center" width="80">�ڵ�</td>
		<td align="center">�ڵ��</td>
		<% if CateDiv="code_mid" then %><td align="center">������</td><% end if %>
		<td align="center" width="70">���ļ���</td>
		<td align="center" width="50">����</td>
	</tr>
	<%
		for lp=0 to oCate.FResultCount - 1

			if CateDiv<>"CateCD1" then
				if oCate.FCateList(lp).Fisusing="<font color=darkblue>���</font>" then
					bgcolor = "#FFFFFF"
				else
					bgcolor = "#E0E0E0"
				end if
			else
				bgcolor = "#FFFFFF"
			end If
			
			Dim modifyUrl 
				If CateDiv = "code_large" Then 
					modifyUrl = "categoryModi2012.asp?CateCD=" & oCate.FCateList(lp).FCateCD &"&page="& page & param &""
				Else
					modifyUrl = "categoryModi2012.asp?code_large=" & oCate.FCateList(lp).FcateLargeCd &"&CateCD=" & oCate.FCateList(lp).FCateCD &"&page="& page & param &""
				End If 
	%>
	<tr align="center" bgcolor="<%=bgcolor%>">
		<td><%= lp + (page * oCate.FPageSize)-oCate.FPageSize+1 %></td>
		<% if CateDiv="code_mid" then %>
		<td><a href="<%=modifyUrl%>"><%= oCate.FCateList(lp).FcateLargeCd %></a></td>
		<td><%= db2html(oCate.FCateList(lp).FlargeCate_Name) %></td>
		<% End If %>
		<td><a href="<%=modifyUrl%>"><%= oCate.FCateList(lp).FCateCD %></a></td>
		<td align="left"><a href="<%=modifyUrl%>"><%= db2html(oCate.FCateList(lp).FCateCD_Name) %></a></td>
		<% if CateDiv="code_mid" then %><td align="left"><%= db2html(oCate.FCateList(lp).FCateCD_NameEng) %></td><% end if %>
		<td><%= oCate.FCateList(lp).FsortNo %></td>
		<td><%= oCate.FCateList(lp).Fisusing %></td>
	</tr>
	<%
		next
	%>
	<tr bgcolor="#FFFFFF">
		<td colspan="8" height="30" align="center">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td align="center" class="a">
				<!-- ������ ���� -->
				<%
					if oCate.HasPreScroll then
						Response.Write "<a href='javascript:goPage(" & oCate.StarScrollPage-1 & ")'>[pre]</a> &nbsp;"
					else
						Response.Write "[pre] &nbsp;"
					end if
		
					for i=0 + oCate.StarScrollPage to oCate.FScrollCount + oCate.StarScrollPage - 1
		
						if i>oCate.FTotalpage then Exit for
		
						if CStr(page)=CStr(i) then
							Response.Write " <font color='red'>[" & i & "]</font> "
						else
							Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
						end if
		
					next
		
					if oCate.HasNextScroll then
						Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
					else
						Response.Write "&nbsp; [next]"
					end if
				%>
				<!-- ������ �� -->
				</td>
				<td width="77" align="right"><a href="categoryWrite2012.asp?page=<%=param%>"><img src="/images/icon_new_registration.gif" width="75" height="20" border="0"></a></td>
			</tr>
			</table>
		</td>
	</tr>
</table>
<%
set oCate = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->