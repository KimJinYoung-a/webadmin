<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ������
' History : ������ ����
'			2017.04.12 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/etcmeachulcls.asp"-->
<!-- #include virtual="/lib/classes/stock/cartoonboxcls.asp"-->
<%
dim idx
idx = requestCheckVar(request("idx"),10)

dim oetcmeachulmaster
set oetcmeachulmaster = new CEtcMeachul
oetcmeachulmaster.FRectidx = idx
oetcmeachulmaster.getOneEtcMeachul


'response.end

dim oetcmeachulsubmaster

set oetcmeachulsubmaster = new CEtcMeachul
oetcmeachulsubmaster.FPageSize=200
oetcmeachulsubmaster.FRectIDx = idx
oetcmeachulsubmaster.getEtcMeachulSubmasterList

dim oCCartoonBoxMasterItem

set oCCartoonBoxMasterItem = new CCartoonBoxMasterItem

oCCartoonBoxMasterItem.Fdelivermethod = oetcmeachulmaster.FOneItem.Fdelivermethod

dim i

dim totalsellcash,totalbuycash,totalsuplycash,totalorgsellcash, totalcount

dim mode

if (oetcmeachulmaster.FOneItem.Fworkidx <> "") then
	mode = "updateworkidx"
else
	mode = "insertworkidx"
end if


dim IsEMSAddNeed : IsEMSAddNeed = False

%>
<script language='javascript'>
function popSubdetailEdit(iid,itopid){
	var popwin = window.open('popetcmeachul_subdetail.asp?idx=' + iid + '&topidx=' + itopid,'franmeaippopsubdetail','width=1100, height=700, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function popEtcAdd(topidx,shopid){
	if ("<%=oetcmeachulmaster.FOneItem.FstateCd%>" >= "4")
	{
		alert("��꼭 ���� ���Ŀ��� ��Ÿ�����߰� �� �� �����ϴ�.")
		return;
	}

	var popwin = window.open('popetcmeachul_etcjungsanadd.asp?topidx=' + topidx + '&shopid=' + shopid,'franetcjungsanadd','width=600, height=200, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function PopExportSheet(v){
	var popwin;
	popwin = window.open('/admin/fran/cartoonbox_modify.asp?menupos=1357&idx=' + v ,'PopExportSheet','width=740,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function UpdateWorkidx(frm) {
	if (CheckBox(frm) == true) {
		if (confirm('�����Ͻðڽ��ϱ�?') == true) {
			frm.submit();
		}
	}
}

function popAddDeliverPay(frm) {
	if (frm == undefined) {
		alert("�ؿ������ EMS����� �߰��� �� �ֽ��ϴ�.");
		return;
	}

	if (frm.workidx.value == "") {
		alert("���� �����۾��� �Է��ϼ���.");
		return;
	}

	if (CheckBox(frm) == true) {
		if (confirm('EMS��ۺ���� �߰��Ͻðڽ��ϱ�?') == true) {
			frm.mode.value = "addemsprice";
			frm.submit();
		}
	}
}

function CheckBox(frm) {
	if (frm.workidx.value*0 != 0) {
		alert("���ڸ� �����մϴ�.");
		frm.workidx.focus();
		return false;
	}

	return true;
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>Index</td>
		<td bgcolor="#FFFFFF" ><%= oetcmeachulmaster.FOneItem.Fidx %></td>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>����ID</td>
		<td bgcolor="#FFFFFF" ><%= oetcmeachulmaster.FOneItem.Fshopid %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>����</td>
		<td bgcolor="#FFFFFF" ><font color="<%= oetcmeachulmaster.FOneItem.GetDivCodeColor %>"><%= oetcmeachulmaster.FOneItem.GetDivCodeName %></font></td>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>Title</td>
		<td bgcolor="#FFFFFF" ><%= oetcmeachulmaster.FOneItem.Ftitle %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>���ǸŰ�</td>
		<td bgcolor="#FFFFFF" ><%= FormatNumber(oetcmeachulmaster.FOneItem.Ftotalsellcash,0) %></td>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>�Ѹ��԰�</td>
		<td bgcolor="#FFFFFF" ><%= FormatNumber(oetcmeachulmaster.FOneItem.Ftotalbuycash,0) %>
		<font color="#AAAAAA">(��ü�κ��� ���޹��� ��ǰ����)</font></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>�Ѱ��ް�</td>
		<td bgcolor="#FFFFFF" ><%= FormatNumber(oetcmeachulmaster.FOneItem.Ftotalsuplycash,0) %>
		<font color="#AAAAAA">(������ ������ ��ǰ����)</font></td>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>�ѹ���ݾ�</td>
		<td bgcolor="#FFFFFF" ><%= FormatNumber(oetcmeachulmaster.FOneItem.Ftotalsum,0) %></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>��꼭������</td>
		<td bgcolor="#FFFFFF" ><%= oetcmeachulmaster.FOneItem.Ftaxdate %></td>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>�Ա���</td>
		<td bgcolor="#FFFFFF" ><%= oetcmeachulmaster.FOneItem.Fipkumdate %></td>
	</tr>
	<% if (CStr(oetcmeachulmaster.FOneItem.Fshopdiv) = "7") or (CStr(oetcmeachulmaster.FOneItem.Fshopdiv) = "8") then %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>����</td>
		<td bgcolor="#FFFFFF">
			<font color="<%= oetcmeachulmaster.FOneItem.GetStateColor %>"><%= oetcmeachulmaster.FOneItem.GetStateName %></font>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>�����۾�(�ؿ�)</td>
		<form name="frmMaster" method="post" action="franmeaippopsubmaster_process.asp">
		<input type="hidden" name="mode" value="<%= mode %>">
		<input type="hidden" name="masteridx" value="<%= idx %>">
		<input type="hidden" name="orgworkidx" value="<%= oetcmeachulmaster.FOneItem.Fworkidx %>">
		<td bgcolor="#FFFFFF" >
			<input type="text" class="text" name="workidx" value="<%= oetcmeachulmaster.FOneItem.Fworkidx %>" size=6 maxlength=6>
			<input type="button" class="button" value="�Է�" onClick="UpdateWorkidx(frmMaster)">
			<% if (oetcmeachulmaster.FOneItem.Fworkidx <> "") then %>
				<input type="button" class="button" value="��ȸ" onClick="PopExportSheet(<%= oetcmeachulmaster.FOneItem.Fworkidx %>)">
				<%
				if (oCCartoonBoxMasterItem.GetDeliverMethodName = "EMS") then
					IsEMSAddNeed = True
				%>
					<font color=red>[��� : <%= oCCartoonBoxMasterItem.GetDeliverMethodName %>]</font>
				<% end if %>
			<% end if %>
		</td>
		</form>
	</tr>
	<% else %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>����</td>
		<td bgcolor="#FFFFFF" colspan=3>
			<font color="<%= oetcmeachulmaster.FOneItem.GetStateColor %>"><%= oetcmeachulmaster.FOneItem.GetStateName %></font>
		</td>
	</tr>
	<% end if %>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>��Ÿ����</td>
		<td bgcolor="#FFFFFF" colspan=3>
		<%= nl2Br(oetcmeachulmaster.FOneItem.Fetcstr) %>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>���ʵ����</td>
		<td bgcolor="#FFFFFF" ><%= oetcmeachulmaster.FOneItem.Fregusername %>(<%= oetcmeachulmaster.FOneItem.Freguserid %>)</td>
		<td bgcolor="<%= adminColor("tabletop") %>" width=100>����ó����</td>
		<td bgcolor="#FFFFFF" ><%= oetcmeachulmaster.FOneItem.Ffinishusername %>(<%= oetcmeachulmaster.FOneItem.Ffinishuserid %>)</td>
	</tr>
</table>

<p>

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="��Ÿ�����߰�" onClick="popEtcAdd('<%= idx %>','<%= oetcmeachulmaster.FOneItem.Fshopid %>')">
			<input type="button" class="button" value="EMS��ۺ���߰�" onClick="popAddDeliverPay(document.frmMaster)" <% if IsNull(oCCartoonBoxMasterItem.GetDeliverMethodName) or (oCCartoonBoxMasterItem.GetDeliverMethodName <> "EMS") then %>disabled<% end if %>>
		</td>
		<td align="right">
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<p>

<% if oetcmeachulmaster.FOneItem.FDivcode="MC" then %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="<%= adminColor("tabletop") %>" align=center height="25">
		<td >���ó</td>
		<td width=70>����ڵ�</td>
		<td width=70>�ֹ��ڵ�</td>
		<td width=70>������</td>
		<td width=70>�����</td>
		<td width=80>���ǸŰ�</td>
		<td width=80><b>�����</b></td>
		<td width=80>�Ѹ��԰�</td>
		<td>���</td>
	</tr>
	<% for i=0 to  oetcmeachulsubmaster.FResultCount -1 %>
	<%
	totalsellcash	=	totalsellcash + oetcmeachulsubmaster.FItemList(i).Ftotalsellcash
	totalbuycash	=	totalbuycash + oetcmeachulsubmaster.FItemList(i).Ftotalbuycash
	totalsuplycash	=	totalsuplycash + oetcmeachulsubmaster.FItemList(i).Ftotalsuplycash

	if (IsEMSAddNeed = True) and (oetcmeachulsubmaster.FItemList(i).Fcode02 = "temp") then
		IsEMSAddNeed = False
	end if
	%>
	<tr align="center" bgcolor="#FFFFFF" height="25">
		<td ><a href="javascript:popSubdetailEdit('<%= oetcmeachulsubmaster.FItemList(i).Fidx %>','<%= idx %>');"><%= oetcmeachulsubmaster.FItemList(i).Fshopid %></a></td>
		<td ><%= oetcmeachulsubmaster.FItemList(i).Fcode01 %></td>
		<td ><%= oetcmeachulsubmaster.FItemList(i).Fcode02 %></td>
		<td ><%= Left(oetcmeachulsubmaster.FItemList(i).Fbaljudate, 10) %></td>
		<td ><%= oetcmeachulsubmaster.FItemList(i).Fexecdate %></td>
		<td align=right><%= FormatNumber(oetcmeachulsubmaster.FItemList(i).Ftotalsellcash,0) %></td>
		<td align=right><b><%= FormatNumber(oetcmeachulsubmaster.FItemList(i).Ftotalsuplycash,0) %></b></td>
		<td align=right><%= FormatNumber(oetcmeachulsubmaster.FItemList(i).Ftotalbuycash,0) %></td>
		<td align=center><input type="button" class="button" value="����" onClick="popSubdetailEdit('<%= oetcmeachulsubmaster.FItemList(i).Fidx %>','<%= idx %>')"></td>
	</tr>
	<% next %>
	<tr align="center" bgcolor="#FFFFFF">
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td align=right><%= FormatNumber(totalsellcash,0) %></td>
		<td align=right><b><%= FormatNumber(totalsuplycash,0) %></b></td>
		<td align=right><%= FormatNumber(totalbuycash,0) %></td>
		<td></td>
	</tr>
</table>

<% elseif oetcmeachulmaster.FOneItem.FDivcode="WS" then %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="<%= adminColor("tabletop") %>" align=center height="25">
		<td>������ID</td>
		<td width=90>�Ǹſ�</td>
		<td>�귣��ID</td>
		<td width=90>�ѼҺ�</td>
		<td width=90>���ǸŰ�</td>
		<td width=90><b>�����</b></td>
		<td width=90>�Ѹ��԰�</td>
		<td>���</td>
	</tr>
	<% for i=0 to  oetcmeachulsubmaster.FResultCount -1 %>
	<%
	totalsellcash	=	totalsellcash + oetcmeachulsubmaster.FItemList(i).Ftotalsellcash
	totalbuycash	=	totalbuycash + oetcmeachulsubmaster.FItemList(i).Ftotalbuycash
	totalsuplycash	=	totalsuplycash + oetcmeachulsubmaster.FItemList(i).Ftotalsuplycash
	totalorgsellcash =  totalorgsellcash + oetcmeachulsubmaster.FItemList(i).Ftotalorgsellcash
	%>
	<tr bgcolor="#FFFFFF" height="25">
		<td align=center ><%= oetcmeachulsubmaster.FItemList(i).Fshopid %></td>
		<td align=center ><a href="javascript:popSubdetailEdit('<%= oetcmeachulsubmaster.FItemList(i).Fidx %>','<%= idx %>');"><%= oetcmeachulsubmaster.FItemList(i).Fcode01 %></a></td>
		<td ><%= oetcmeachulsubmaster.FItemList(i).Fcode02 %></td>
		<td align=right><%= Formatnumber(oetcmeachulsubmaster.FItemList(i).Ftotalorgsellcash,0) %></td>
		<td align=right><%= Formatnumber(oetcmeachulsubmaster.FItemList(i).Ftotalsellcash,0) %></td>
		<td align=right><b><%= Formatnumber(oetcmeachulsubmaster.FItemList(i).Ftotalsuplycash,0) %></b></td>
		<td align=right><%= Formatnumber(oetcmeachulsubmaster.FItemList(i).Ftotalbuycash,0) %></td>
		<td align=center><input type="button" class="button" value="����" onClick="popSubdetailEdit('<%= oetcmeachulsubmaster.FItemList(i).Fidx %>','<%= idx %>')"></td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td></td>
		<td></td>
		<td></td>
		<td align=right><%= FormatNumber(totalorgsellcash,0) %></td>
		<td align=right><%= FormatNumber(totalsellcash,0) %></td>
		<td align=right><b><%= FormatNumber(totalsuplycash,0) %></b></td>
		<td align=right><%= FormatNumber(totalbuycash,0) %></td>
		<td></td>
	</tr>
</table>

<% elseif oetcmeachulmaster.FOneItem.FDivcode="AA" or oetcmeachulmaster.FOneItem.FDivcode="BB" or oetcmeachulmaster.FOneItem.FDivcode="CC" then %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="<%= adminColor("tabletop") %>" align=center height="25">
		<td>������ID</td>
		<td width=90>�Ǹ���</td>
		<td>�ѰǼ�</td>
		<td width=90>�ѼҺ�</td>
		<td width=90>���ǸŰ�</td>
		<td width=90><b>�����</b></td>
		<td width=90>�Ѹ��԰�</td>
		<td>���</td>
	</tr>
	<% for i=0 to  oetcmeachulsubmaster.FResultCount -1 %>
	<%
	totalcount		= 	totalcount + oetcmeachulsubmaster.FItemList(i).Ftotalcount
	totalsellcash	=	totalsellcash + oetcmeachulsubmaster.FItemList(i).Ftotalsellcash
	totalbuycash	=	totalbuycash + oetcmeachulsubmaster.FItemList(i).Ftotalbuycash
	totalsuplycash	=	totalsuplycash + oetcmeachulsubmaster.FItemList(i).Ftotalsuplycash
	totalorgsellcash =  totalorgsellcash + oetcmeachulsubmaster.FItemList(i).Ftotalorgsellcash
	%>
	<tr bgcolor="#FFFFFF" height="25">
		<td align=center ><%= oetcmeachulsubmaster.FItemList(i).Fshopid %></td>
		<td align=center ><a href="javascript:popSubdetailEdit('<%= oetcmeachulsubmaster.FItemList(i).Fidx %>','<%= idx %>');"><%= oetcmeachulsubmaster.FItemList(i).Fcode01 %></a></td>
		<td align=right><%= Formatnumber(oetcmeachulsubmaster.FItemList(i).Ftotalcount,0) %></td>
		<td align=right><%= Formatnumber(oetcmeachulsubmaster.FItemList(i).Ftotalorgsellcash,0) %></td>
		<td align=right><%= Formatnumber(oetcmeachulsubmaster.FItemList(i).Ftotalsellcash,0) %></td>
		<td align=right><b><%= Formatnumber(oetcmeachulsubmaster.FItemList(i).Ftotalsuplycash,0) %></b></td>
		<td align=right><%= Formatnumber(oetcmeachulsubmaster.FItemList(i).Ftotalbuycash,0) %></td>
		<td align=center><input type="button" class="button" value="����" onClick="popSubdetailEdit('<%= oetcmeachulsubmaster.FItemList(i).Fidx %>','<%= idx %>')"></td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td></td>
		<td></td>
		<td align=right><%= FormatNumber(totalcount,0) %></td>
		<td align=right><%= FormatNumber(totalorgsellcash,0) %></td>
		<td align=right><%= FormatNumber(totalsellcash,0) %></td>
		<td align=right><b><%= FormatNumber(totalsuplycash,0) %></b></td>
		<td align=right><%= FormatNumber(totalbuycash,0) %></td>
		<td></td>
	</tr>
</table>

<% else %>

<!-- �󼼳��� ���� -->

<% end if %>
<%
set oetcmeachulmaster = Nothing
set oetcmeachulsubmaster = Nothing

if (IsEMSAddNeed = True) then
	response.write "<script>alert('EMS��ۺ���� �߰��ϼ���.')</script>"
end if

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
