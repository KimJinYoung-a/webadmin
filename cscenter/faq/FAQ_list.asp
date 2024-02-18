<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [CS]��������>>[FAQ]���� 
' Hieditor : 2009.03.02 �̿��� ����
'			 2021.07.30 �ѿ�� ����(��뿩�� �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/cscenter/faq_cls.asp"-->
<%
	'// ���� ���� //
	dim faqid
	dim page, searchDiv, searchKey, searchString, param

	dim ofaq, i, lp, bgcolor, strUsing,isusing


	'// �Ķ���� ���� //
	faqid = request("faqid")
	page = request("page")
	searchDiv = request("searchDiv")
	searchKey = request("searchKey")
	searchString = request("searchString")
	isusing = requestcheckvar(request("isusing"),1)

	if page="" then page=1
	if searchKey="" then searchKey="title"

	param = "&searchKey=" & searchKey & "&searchString=" & searchString & "&searchDiv=" & searchDiv

	'// Ŭ���� ����
	set ofaq = new Cfaq
	ofaq.FCurrPage = page
	ofaq.FPageSize = 20
	ofaq.FRectsearchDiv = searchDiv
	ofaq.FRectsearchKey = searchKey
	ofaq.FRectsearchString = searchString
	ofaq.FRectisusing = isusing
	ofaq.GetFAQList
%>
<script language='javascript'>
<!--
	function chk_form(){
		var frm = document.frm_search;

//		if(!frm.searchKey.value){
//			alert("�˻� ������ �������ֽʽÿ�.");
//			frm.searchKey.focus();
//			return;
//		}
//		else if(!frm.searchString.value)
//		{
//			alert("�˻�� �Է����ֽʽÿ�.");
//			frm.searchString.focus();
//			return;
//		}

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
<!-- �˻� ���� -->
<form name="frm_search" method="POST" action="faq_list.asp" onSubmit="return false">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			����
    		<select name="searchDiv" class="select" onChange="goPage(frm_search.page.value)">
    			<option value="">����</option>
    			<%= db2html((ofaq.optCommCd("Z200", searchDiv))) %>
    		</select>
    		/ �˻�
    		<select name="searchKey" class="select">
    			<option value="">����</option>
    			<option value="title">����+����</option>
    		</select>
    		<script language="javascript">
    			document.frm_search.searchKey.value="<%=searchKey%>";
    		</script>
    		<input type="text" class="text" name="searchString" size="20" value="<%= searchString %>">
			/ ��뿩�� : <% drawSelectBoxUsingYN "isusing", isusing %>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="chk_form()">
		</td>
	</tr>
</table>
</form>
<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
		<input type="button" class="button" value="�űԵ��" onClick="location.href='faq_write.asp?menupos=<%=menupos%>'">			
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="11">
			�˻���� : <b><%= ofaq.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %> / <%= ofaq.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="40">IDX</td>
		<td width="40">����</td>
		<td width="60">�����ڵ�</td>
		<td width="120">����</td>
		<td>����</td>
		<td width="50">LinkURL</td>
		<td>Link��</td>
		<td width="70">�����</td>
		<td width="50">��ȸ��</td>
		<td width="80">�����</td>
		<td width="50">�������</td>
	</tr>
	<%
		for lp=0 to ofaq.FResultCount - 1
	%>
	<tr align="center" <% if ofaq.FfaqList(lp).Fisusing = "N" then %> bgcolor="#EEEEEE" <% else %> bgcolor="#FFFFFF" <% end if %> >
	    <td><%= ofaq.FfaqList(lp).Ffaqid %></td>
		<td><%= ofaq.FfaqList(lp).Fdisporder %></td>
		<td><%= ofaq.FfaqList(lp).FcommCd %></td>
		<td align="left"><%= db2html(ofaq.FfaqList(lp).Fcomm_name) %></td>
		<td align="left">
			<a href="faq_view.asp?faqid=<%= ofaq.FfaqList(lp).Ffaqid %>&page=<%=page & param%>&menupos=<%=menupos%>">
			<%= ReplaceBracket(db2html(ofaq.FfaqList(lp).Ftitle)) %></a>
		</td>
	    <td>
	        <% if ofaq.FfaqList(lp).Flinkurl<>"" then %>
	        <acronym title="<%= ReplaceBracket(db2html(ofaq.FfaqList(lp).Flinkurl)) %>">YES</acronym>
	        <% end if %>
	    </td>
	    <td><a href="<%= ReplaceBracket(db2html(ofaq.FfaqList(lp).Flinkurl)) %>" target="_blank"><%= ReplaceBracket(ofaq.FfaqList(lp).Flinkname) %></a></td>
		<td><%= ofaq.FfaqList(lp).Fuserid %></td>
		<td><%= ofaq.FfaqList(lp).FhitCount %></td>
		<td><%= FormatDate(ofaq.FfaqList(lp).Fregdate,"0000.00.00") %></td>
	    <td><%= ofaq.FfaqList(lp).Fisusing %></td>
	</tr>
	<%
		next
	%>
	<tr bgcolor="#FFFFFF">
		<td colspan="11" height="30" align="center">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td align="center" class="a">
					<% sbDisplayPaging "page="&page, ofaq.FTotalCount, ofaq.FPageSize, 10%>
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
