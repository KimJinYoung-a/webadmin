<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� qna ����Ʈ
' Hieditor : 2009.11.27 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
'// ���� ���� //
dim qnaId , page, searchDiv, searchKey, searchString, param, isanswer
dim oQnA, i, lp, bgcolor, strUsing
	'// �Ķ���� ���� //
	qnaId = request("qnaId")
	page = request("page")
	searchDiv = 1
	searchKey = request("searchKey")
	searchString = request("searchString")
	isanswer = request("isanswer")

	if page="" then page=1
	if searchKey="" then searchKey="qstTitle"
	if isanswer="" then isanswer="N"

'// Ŭ���� ����
set oQnA = new CQnA
	oQnA.FCurrPage = page
	oQnA.FPageSize = 20
	oQnA.FRectsearchDiv = searchDiv
	oQnA.FRectsearchKey = searchKey
	oQnA.FRectsearchString = searchString
	oQnA.FRectisanswer = isanswer
	oQnA.GetQnAList
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
<form name="frm_search" method="GET" action="QnA_list.asp" onSubmit="return chk_form(this)">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>		
		<td align="left">
			����
			<select name="isanswer" onChange="goPage(frm_search.page.value)">
				<option value="Y">�Ϸ�</option>
				<option value="N">���</option>
			</select>
			/ �˻�
			<select name="searchKey">
				<option value="">����</option>
				<option value="qnaId">������ȣ</option>
				<option value="qstTitle">����</option>
				<option value="qstContents">����</option>
			</select>
			<script language="javascript">
				document.frm_search.isanswer.value="<%=isanswer%>";
				document.frm_search.searchKey.value="<%=searchKey%>";
			</script>
			<input type="text" name="searchString" size="20" value="<%= searchString %>">	
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="frm_search.submit();">
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
	<% if oQnA.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= oQnA.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= oQnA.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" width="40">��ȣ</td>
		<td align="center">����</td>
		<td align="center" width="70">�����</td>
		<td align="center" width="50">����</td>
		<td align="center" width="80">�����</td>
		<td align="center" width="120">��뿩��</td>
		<td align="center" width="80">���</td>
    </tr>
	<%
		for lp=0 to oQnA.FResultCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
		<td><%= oQnA.FQnAList(lp).FqnaId %></td>
		<td align="left"><%= db2html(oQnA.FQnAList(lp).FqstTitle) %></td>
		<td><%= oQnA.FQnAList(lp).FqstUserId %></td>
		<td><%= oQnA.FQnAList(lp).Fisanswer %></td>
		<td><%= FormatDate(oQnA.FQnAList(lp).Fregdate,"0000.00.00") %></td>
		<td><%= oQnA.FQnAList(lp).fisusing %></td>
		<td><input type="button" onclick="location.href='QnA_view.asp?qnaId=<%= oQnA.FQnAList(lp).FqnaId %>'" value="����" class="button"></td>	
	</tr>
	<%
		next
	%>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="10" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<!-- ������ ���� -->
			<%
				if oQnA.HasPreScroll then
					Response.Write "<a href='javascript:goPage(" & oQnA.StarScrollPage-1 & ")'>[pre]</a> &nbsp;"
				else
					Response.Write "[pre] &nbsp;"
				end if
	
				for i=0 + oQnA.StarScrollPage to oQnA.FScrollCount + oQnA.StarScrollPage - 1
	
					if i>oQnA.FTotalpage then Exit for
	
					if CStr(page)=CStr(i) then
						Response.Write " <font color='red'>[" & i & "]</font> "
					else
						Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
					end if
	
				next
	
				if oQnA.HasNextScroll then
					Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
				else
					Response.Write "&nbsp; [next]"
				end if
			%>
			<!-- ������ �� -->
		</td>
	</tr>
</table>

<%
	set oQnA = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->