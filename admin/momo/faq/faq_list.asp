<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� faq ����Ʈ
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
dim ntcId
dim page, searchDiv, searchKey, searchString
dim ofaq, i, lp, bgcolor, strUsing
	ntcId = request("ntcId")
	page = request("page")
	searchDiv = 2
	searchKey = request("searchKey")
	searchString = request("searchString")

	if page="" then page=1
	if searchKey="" then searchKey="title"

	'// Ŭ���� ����
	set ofaq = new CNotice
	ofaq.FCurrPage = page
	ofaq.FPageSize = 20
	ofaq.FRectsearchDiv = searchDiv
	ofaq.FRectsearchKey = searchKey
	ofaq.FRectsearchString = searchString
	ofaq.GetNoitceList
%>
<script language='javascript'>

	function chk_form()
	{
		var frm = document.frm_search;
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

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm_search" method="POST">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>		
		<td align="left">
			<select name="searchKey">
				<option value="">����</option>
				<option value="ntcId">������ȣ</option>
				<option value="title">����</option>
				<option value="contents">����</option>
			</select>
			<script language="javascript">
				document.frm_search.searchKey.value="<%=searchKey%>";
			</script>
			<input type="text" name="searchString" size="20" value="<%= searchString %>">		
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="chk_form();">
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
			<input type="button" onclick="javascript:location.href='faq_modi.asp'" value="�űԵ��" class="button">										
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if ofaq.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= ofaq.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= ofaq.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" >��ȣ</td>		
		<td align="center">����</td>
		<td align="center" >�����</td>
		<td align="center">�����</td>
		<td align="center" >��뿩��</td>
		<td align="center">���</td>
    </tr>
	<%
		for lp=0 to ofaq.FResultCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
		<td><%= ofaq.FNoticeList(lp).FntcId %></td>		
		<td align="left"><%= db2html(ofaq.FNoticeList(lp).Ftitle) %></td>
		<td><%= ofaq.FNoticeList(lp).Fuserid %></td>
		<td><%= FormatDate(ofaq.FNoticeList(lp).Fregdate,"0000.00.00") %></td>
		<td><%= ofaq.FNoticeList(lp).fisusing %></td>		
		<td><input type="button" onclick="location.href='faq_view.asp?ntcId=<%= ofaq.FNoticeList(lp).FntcId %>'" value="�󼼺���" class="button"></td>
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
	</tr>
</table>

<%
	set ofaq = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->