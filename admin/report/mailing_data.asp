<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ���� ���
' History : 2007.08.27 �ѿ�� ����
' History : 2016.12.07 ���¿� ������ �� ������ ����Ʈ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
response.write "������� �Ŵ� �Դϴ�. �������v2>>�������������� ����� �ּ���."
response.end

dim page, i, omd
	page = requestcheckvar(getNumeric(request("page")),10)

if page="" then page=1

set omd = New CMailzine
	omd.FCurrPage = page
	omd.FPageSize=100
	omd.GetMailingList
%>
<script type="text/javascript">

function TnMailDataReg(frm){
	if(frm.title.value == ""){
		alert("�߼��̸��� �����ּ���");
		frm.title.focus();
	}
	else if(frm.gubun.value == ""){
		alert("�߼۱����� �����ּ���");
		frm.gubun.focus();
	}
	else if(frm.startdate.value == ""){
		alert("�߼۽��۽ð��� �����ּ���");
		frm.startdate.focus();
	}
	else if(frm.enddate.value == ""){
		alert("�߼�����ð��� �����ּ���");
		frm.enddate.focus();
	}
	else if(frm.reenddate.value == ""){
		alert("��߼�����ð��� �����ּ���");
		frm.reenddate.focus();
	}
	else if(frm.totalcnt.value == ""){
		alert("�Ѵ���ڼ��� �����ּ���");
		frm.totalcnt.focus();
	}
	else if(frm.realcnt.value == ""){
		alert("�ǹ߼������ �����ּ���");
		frm.realcnt.focus();
	}
	else if(frm.realpct.value == ""){
		alert("�ǹ߼ۺ����� �����ּ���");
		frm.realpct.focus();
	}
	else if(frm.filteringcnt.value == ""){
		alert("���͸������ �����ּ���");
		frm.filteringcnt.focus();
	}
	else if(frm.filteringpct.value == ""){
		alert("���͸������� �����ּ���");
		frm.filteringpct.focus();
	}
	else if(frm.successcnt.value == ""){
		alert("�����߼������ �����ּ���");
		frm.successcnt.focus();
	}
	else if(frm.successpct.value == ""){
		alert("�������� �����ּ���");
		frm.successpct.focus();
	}
	else if(frm.failcnt.value == ""){
		alert("���й߼������ �����ּ���");
		frm.failcnt.focus();
	}
	else if(frm.failpct.value == ""){
		alert("�������� �����ּ���");
		frm.failpct.focus();
	}
	else if(frm.opencnt.value == ""){
		alert("��������� �����ּ���");
		frm.opencnt.focus();
	}
	else if(frm.openpct.value == ""){
		alert("�������� �����ּ���");
		frm.openpct.focus();
	}
	else if(frm.noopencnt.value == ""){
		alert("�̿�������� �����ּ���");
		frm.noopencnt.focus();
	}
	else if(frm.noopenpct.value == ""){
		alert("�̿������� �����ּ���");
		frm.noopenpct.focus();
	}
	else{
		frm.submit();
	}
}

</script>

<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%=adminColor("tablebg")%>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= omd.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%=adminColor("tabletop")%>" height="30">
	<td>�߼� �̸�</td>
	<td>���� ����</td>
	<td>�� ����ڼ�</td>
	<td>���� �߼ۼ�</td>
	<td>���� ���</td>
	<td>Ŭ�� ���</td>
	<td>�߼� �ð�</td>
	<td>�Ϸ� �ð�</td>
	<td>���Ϸ�</td>
	<td>ETC</td>
</tr>
<% if omd.FResultCount>0 then %>
	<% for i=0 to omd.FResultCount-1 %>
	<tr bgcolor="FFFFFF">
		<td width="200"><% = omd.FItemList(i).Ftitle %></td>
		<td width="200"><% = omd.FItemList(i).fsubject %></td>
		<td align="right"><% = FormatNumber(omd.FItemList(i).Ftotalcnt,0) %></td>
		<td align="right"><% = FormatNumber(omd.FItemList(i).fsuccesscnt,0) %></td>
		<td align="right"><% = FormatNumber(omd.FItemList(i).fopencnt,0) %></td>
		<td align="right"><% = FormatNumber(omd.FItemList(i).fclickcnt,0) %></td>
		<td align="center"><% = omd.FItemList(i).Fstartdate %></td>
		<td align="center"><% = omd.FItemList(i).Fenddate %></td>
		<td align="center"><% = omd.FItemList(i).fmailergubun %></td>
		<td align="center"><a href="mailing_data_reg.asp?idx=<% = omd.FItemList(i).Fidx %>&mode=edit">�󼼳��뺸��</a></td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% End If %>
</table>

<% set omd = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->