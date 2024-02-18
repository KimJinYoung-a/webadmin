<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [��ü]��������
' Hieditor : ������ ����
'			 2023.10.23 �ѿ�� ����(�̸��Ϲ߼� cdo->���Ϸ��� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/boardcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<%
dim i, brdid, mduserid, catecode, targt, mode, page, nboardmail
	brdid		= requestcheckvar(getNumeric(request("id")),10)
	mduserid    = requestcheckvar(request("mduserid"),32)
	catecode    = request("catecode")
	targt		= request("targt")
	mode		= requestcheckvar(request("mode"),32)
	page		= requestcheckvar(getNumeric(request("page")),10)

if page="" then page=1
if targt="" then targt="basic"

%>
<script type='text/javascript'>

function SendEmail(frm){
	if (confirm('��ü ��ü������ �߼� �Ͻðڽ��ϱ�?')){
		frm.action="dodesignernoticemail.asp";
		frm.method="POST";
		frm.submit();
	}
}

function previewTarget(frm) {
	frm.action="";
	frm.method="GET";
	frm.submit();
}

function goPage(pg) {
	frm.page.value=pg;
	frm.action="";
	frm.method="GET";
	frm.submit();
}

</script>
<span style="font-size:13px; font-weight:bold;">�� ��ü �������� �߼�</span>

<form name="frm" method="post" action="" style="margin:0px;">
<input type="hidden" name="id" value="<%=brdid%>">
<input type="hidden" name="mode" value="upcheall">
<input type="hidden" name="page" value="1">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td>
		* ����� : <% drawSelectBoxCoWorker "mduserid", mduserid %>
		&nbsp;
		* ī�װ� : <% SelectBoxBrandCategory "catecode", catecode %>
		&nbsp;
		<br>
		* �߼۴�� :
			<label><input type="radio" name="targt" value="basic" <% if targt="basic" then Response.Write "checked" %> onfocus="this.blur()">�⺻�����</label> &nbsp;
			<label><input type="radio" name="targt" value="deliver" <% if targt="deliver" then Response.Write "checked" %> onfocus="this.blur()">��۴����</label> &nbsp;
			<label><input type="radio" name="targt" value="account" <% if targt="account" then Response.Write "checked" %> onfocus="this.blur()">��������</label>
	</td>
	<td align="center" width=60>
		<input type="button" value="���Ϲ߼�" onClick="SendEmail(frm);" class="button">
	</td>
</tr>

</table>
</form>

<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<input type="button" value="����� ����" onClick="previewTarget(frm)" class="button">
	</td>
	<td align="right"></td>
</tr>
</table>
<!-- �׼� �� -->

<%
'// ����� �����ϰ�� ��� ���
if mode="upcheall" then
	set nboardmail = new CBoard
		nboardmail.FCurrPage		= page
		nboardmail.FPageSize		= 15
		nboardmail.FRectMDid		= mduserid
		nboardmail.FRectCatCD		= catecode
		nboardmail.FRectTarget		= targt
		nboardmail.design_notice_mail_preview
%>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
	<tr bgcolor="FFFFFF">
		<td colspan="3">
			�˻���� : <b><%= nboardmail.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %>/ <%= nboardmail.FTotalPage %></b>
			&nbsp;
			������(�ߺ�����) : <b><%= nboardmail.Fint_total %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td><b>�귣��ID</b></td>
		<td><b>�귣���</b></td>
		<td><b>�̸���</b></td>
	</tr>

	<% if nboardmail.FResultCount>0 then %>
		<% for i=0 to nboardmail.FResultCount-1 %>	
		<tr align='center' bgcolor='#FFFFFF'>
			<td><%= nboardmail.BoardItem(i).FRectDesignerID %></td>
			<td><%= nboardmail.BoardItem(i).FRectName %></td>
			<td><%= nboardmail.BoardItem(i).FRectEmail %></td>
		</tr>	
		<% next %>

		<tr height="25" bgcolor="FFFFFF">
			<td colspan="3" align="center">
				<% if nboardmail.HasPreScroll then %>
					<span class="list_link"><a href="javascript:goPage(<%= nboardmail.StartScrollPage-1 %>)">[pre]</a></span>
				<% else %>
				[pre]
				<% end if %>
				<% for i = 0 + nboardmail.StartScrollPage to nboardmail.StartScrollPage + nboardmail.FScrollCount - 1 %>
					<% if (i > nboardmail.FTotalpage) then Exit for %>
					<% if CStr(i) = CStr(nboardmail.FCurrPage) then %>
					<span class="page_link"><font color="red"><b><%= i %></b></font></span>
					<% else %>
					<a href="javascript:goPage(<%= i %>)" class="list_link"><font color="#000000"><%= i %></font></a>
					<% end if %>
				<% next %>
				<% if nboardmail.HasNextScroll then %>
					<span class="list_link"><a href="javascript:goPage(<%= i %>)">[next]</a></span>
				<% else %>
				[next]
				<% end if %>
			</td>
		</tr>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
		</tr>
	<% end if %>
	</table>
<%
	set nboardmail = Nothing
end if
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->