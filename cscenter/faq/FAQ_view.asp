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
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cscenter/faq_cls.asp"-->
<%
	'// ���� ���� //
	dim faqid
	dim page, searchDiv, searchKey, searchString, param

	dim ofaq, i, lp

	'// �Ķ���� ���� //
	faqid = request("faqid")
	page = request("page")
	searchDiv = request("searchDiv")
	searchKey = request("searchKey")
	searchString = request("searchString")

	if page="" then page=1
	if searchKey="" then searchKey="titleLong"

	param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'������ ����

	'// ���� ����
	set ofaq = new Cfaq
	ofaq.FRectfaqid = faqid

	ofaq.GetFAQRead

%>
<script language="javascript">

	// �ۻ���
	function GotofaqDel(){
		if (confirm('���� �Ͻðڽ��ϱ�?')){
            document.frm_trans.mode.value = "DEL";
			document.frm_trans.submit();
		}
	}
	
    // �����ȯ
	function GotofaqUsing(){
		if (confirm('�����ȯ �Ͻðڽ��ϱ�?')){
            document.frm_trans.mode.value = "USE";
			document.frm_trans.submit();
		}
	}

</script>
<!-- ���� ȭ�� ���� -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#F0F0FD">
	<td colspan="2">
		<table width="100%" border="0" class="a" cellpadding="0" cellspacing="0">
		<tr height="25">
			<td align="left"><b>FAQ �� ����</b></td>
			<td align="right"><%=ofaq.FfaqList(0).Fregdate%>&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">�ۼ���</td>
	<td bgcolor="#FFFFFF"><%= ofaq.FfaqList(0).Fusername & "(" & ofaq.FfaqList(0).Fuserid & ")" %></td>
</tr>
<%	if Not(ofaq.FfaqList(0).FlastWorker="" or isNull(ofaq.FfaqList(0).FlastWorker)) then %>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">��������</td>
	<td bgcolor="#FFFFFF"><%= ofaq.FfaqList(0).FlastWorkerName & "(" & ofaq.FfaqList(0).FlastWorker & ") / " & ofaq.FfaqList(0).FlastUpdate %></td>
</tr>
<%	end if %>
<tr>
	<td align="center" bgcolor="#DDDDFF">����</td>
	<td bgcolor="#FFFFFF"><%= db2html(ofaq.FfaqList(0).Fcomm_name) %></td>
</tr>
<tr>
	<td align="center" bgcolor="#DDDDFF">ǥ�ü���</td>
	<td bgcolor="#FFFFFF"><%= ofaq.FfaqList(0).Fdisporder %></td>
</tr>
<tr>
	<td align="center" bgcolor="#DDDDFF">����</td>
	<td bgcolor="#F8F8FF"><%= ReplaceBracket(db2html(ofaq.FfaqList(0).Ftitle)) %></td>
</tr>
<tr>
	<td align="center" bgcolor="#DDDDFF">����</td>
	<td bgcolor="#FFFFFF"><%= nl2br(ReplaceBracket(db2html(ofaq.FfaqList(0).Fcontents))) %></td>
</tr>
<tr>
	<td align="center" bgcolor="#DDDDFF">Link��</td>
	<td bgcolor="#FFFFFF"><%= ReplaceBracket(db2html(ofaq.FfaqList(0).Flinkname)) %></td>
</tr>
<tr>
	<td align="center" bgcolor="#DDDDFF">LinkURL</td>
	<td bgcolor="#FFFFFF">
        <% if ofaq.FfaqList(0).Flinkurl<>"" then %>
    	<%= ReplaceBracket(db2html(ofaq.FfaqList(0).Flinkurl)) %>
    	&nbsp;&nbsp;
    	<a href="<%= db2html(ofaq.FfaqList(0).Flinkurl) %>" target="_blank"><font color="blue">>><%= db2html(ofaq.FfaqList(0).Flinkname) %> �ٷΰ���</font></a>
    	<% end if %>
    </td>
</tr>
<tr>
	<td align="center" bgcolor="#DDDDFF">��뿩��</td>
	<td bgcolor="#FFFFFF">
        <%= ofaq.FfaqList(0).fisusing %>
    </td>
</tr>
<tr>
	<td colspan="2" height="30" bgcolor="#FAFAFA" align="center">
		<input type="button" class="button" value="����" onClick="self.location='faq_modi.asp?menupos=<%=menupos%>&faqid=<%=faqid & param%>'"> &nbsp;
		<% if ofaq.FfaqList(0).Fisusing = "Y" then %>
		<input type="button" class="button" name="mode" value="����" onClick="GotofaqDel()"> &nbsp;
		<% else %>
		<input type="button" class="button" name="mode" value="�����ȯ" onClick="GotofaqUsing()"> &nbsp;
	    <% end if %>
		<input type="button" class="button" value="����Ʈ" onClick="self.location='faq_list.asp?menupos=<%=menupos & param %>'">
	</td>
</tr>
<form name="frm_trans" method="POST" action="faq_process.asp">
<input type="hidden" name="faqid" value="<%=faqid%>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
</form>


</table>
<!-- ���� ȭ�� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->