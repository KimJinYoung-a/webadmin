<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2009.09.10 �ѿ�� ����/�߰�
'	Description : ��Ʈ�ʽ�
'#######################################################
%>
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
	set oLecture = new CPartnerLecture
	oLecture.FCurrPage = page
	oLecture.FPageSize = 20
	oLecture.FRectsearchKey = searchKey
	oLecture.FRectsearchString = searchString
	oLecture.FRectsearchConfirm = searchConfirm

	oLecture.GetPartnerLectureList
%>

<script language='javascript'>

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

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm_search" method="GET" action="PartnerLecture_list.asp" onSubmit="return chk_form(this)">
	<input type="hidden" name="page" value="<%=page%>">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
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
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm_search.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">

		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
		</td>
		<td align="right">	
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= oLecture.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= oLecture.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
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
	if oLecture.FTotalCount > 0 then 
		
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
		<td align="left"><a href="/academy/Partnership/PartnerLecture_view.asp?idx=<%= oLecture.FItemList(lp).Fidx %>&page=<%=page & param%>"><%= oLecture.FItemList(lp).Flectitle %></a></td>
		<td><a href="/academy/Partnership/PartnerLecture_view.asp?idx=<%= oLecture.FItemList(lp).Fidx %>&page=<%=page & param%>"><%= oLecture.FItemList(lp).Flecname %></a></td>
		<td><%= oLecture.FItemList(lp).Flecbirthday %></td>
		<td><%= oLecture.FItemList(lp).Flectel %></td>
		<td><%= oLecture.FItemList(lp).Flechp %></td>
		<td><%= FormatDate(oLecture.FItemList(lp).Fregdate,"0000.00.00") %></td>
		<td align="center"><%=strUsing%></td>
	</tr>
	<%
	next
	%>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
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
	<% end if %>
</table>

<%
	set oLecture = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->