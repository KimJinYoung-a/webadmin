<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ٹ����� ������
' History : 2018.04.27 �̻� ����(���Ϸ� ���� ���� ���Ϸ��� �߼� ���� ����. ���� �������� ����.)
'			2019.06.24 ������ ����(���ø� ��� �ű� �߰�)
'			2020.05.28 �ѿ�� ����(TMS ���Ϸ� �߰�)
'###########################################################
%>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbTMSopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/mailzinenewcls.asp"-->

<%
Dim omail,ix,page,sDt, eDt, area, isusing, SearchKey, mailergubun
	page = requestcheckvar(getNumeric(request("page")),10)
	sDt = requestcheckvar(request("sDt"),10)
	eDt = requestcheckvar(request("eDt"),10)
	area = requestcheckvar(request("area"),32)
	isusing = requestcheckvar(request("isusing"),1)
	SearchKey = requestcheckvar(request("SearchKey"),256)
	mailergubun = requestcheckvar(request("mailergubun"),16)

if page = "" then page = 1
if sDt = "" then sDt = dateadd("d",-30,date)
if eDt = "" then eDt = dateadd("d",7,date)
'if mailergubun = "" then mailergubun = "EMS"

if mailergubun="" or isnull(mailergubun) then
	response.write "���Ϸ� ������ �����ϴ�."
	dbget.close() : response.end
end if

set omail = new CMailzineList
	omail.FPageSize = 20
	omail.FCurrPage = page
	omail.FrectSDate = sDt
	omail.FrectEDate = eDt
	omail.FrectSearchKey = SearchKey
	omail.FrectUsing = isusing
	omail.FrectArea = area
	omail.frectmailergubun = mailergubun

	if mailergubun<>"" then
		omail.MailzineList()
	end if
%>

<script type="text/javascript">

// ���(�������)
function editreg(idx){
	var editreg = window.open('/admin/mailzine/mailzine_detail.asp?idx='+idx+'&mailergubun=<%= mailergubun %>&menupos=<%= menupos %>','editreg','width=1400,height=800,scrollbars=yes,resizable=yes');
	editreg.focus();
}

// ���(�ڵ�)
function jsModifyMailzine(idx) {
	var popwin = window.open('/admin/mailzine/mailzine_detail_new.asp?idx='+idx+'&mailergubun=<%= mailergubun %>&menupos=<%= menupos %>','jsModifyMailzine','width=1400,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ���(�ڵ�,���ø�)
function jsModifyNewMailzine(idx) {
	var popwin = window.open('/admin/mailzine/template/mailzine_detail_setting.asp?idx='+idx+'&mailergubun=<%= mailergubun %>&menupos=<%= menupos %>','jsModifyMailzine','width=1400,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function siteyn(idx, isusing, gubun){

	var chisusing;

	if(isusing == 'Y'){
		chisusing = 'N';
	}else{
		chisusing = 'Y';
		if (gubun != '5') {
			alert('\n\n�������� �ϼ����°� �ƴմϴ�.\n\n');
			return false;
		}
	}

	if (confirm('������´�'+isusing+'�Դϴ�.\n'+chisusing+'�� �����Ͻðڽ��ϱ�?') == true) {
		FrameCKP.location.href='/admin/mailzine/mailzine_siteyn.asp?idx='+idx+'&isusing='+isusing+'&menupos=<%= menupos %>';
	}else{
		return false;
	}
}

function blackListReg(){
	var popBlackList = window.open('/admin/mailzine/mailzine_blacklist_pop.asp?menupos=<%= menupos %>','popBlackListReg','width=600,height=200,scrollbars=yes,resizable=yes');
	popBlackList.focus();
}

function displayManual(idx,member){
	var popDisplayManual = "";

	if(member=='member'){
		popDisplayManual = window.open('/admin/mailzine/mailzine_display.asp?idx='+idx+'&menupos=<%= menupos %>','displayManual','width=1400,height=800,scrollbars=yes,resizable=yes');
	}else{
		popDisplayManual = window.open('/admin/mailzine/mailzine_display_not.asp?idx='+idx+'&menupos=<%= menupos %>','displayManual','width=1400,height=800,scrollbars=yes,resizable=yes');
	}

	popDisplayManual.focus();
}

function displayNew(idx, member, type) {
	var display = window.open('/admin/mailzine/mailzine_display_new.asp?idx='+idx + '&member=' + member + '&type=' + type+'&menupos=<%= menupos %>','displayNew','width=1400,height=800,scrollbars=yes,resizable=yes');
	display.focus();
}

// �������ø� ���� �߼�
function displayTemplates(idx, member, type) {
	var popDisplayTemplates = window.open('/admin/mailzine/template/mailzine_display.asp?idx='+idx + '&member=' + member + '&type=' + type+'&mailergubun=<%= mailergubun %>&menupos=<%= menupos %>','displayTemplates','width=1400,height=800,scrollbars=yes,resizable=yes');
	popDisplayTemplates.focus();
}

function mailCodeView(idx,member){
	var popmailCodeView = "";

	if(member=='member'){
		popmailCodeView = window.open('/admin/mailzine/mailzine_code_view.asp?idx='+idx+'&menupos=<%= menupos %>','code','width=1400,height=800,scrollbars=yes,resizable=yes');
	}else if(member=='basicMailFormCopy'){
		popmailCodeView = window.open('/admin/mailzine/mailzine_target_templet.asp?menupos=<%= menupos %>','code','width=1400,height=800,scrollbars=yes,resizable=yes');
	}else{
		popmailCodeView = window.open('/admin/mailzine/mailzine_code_view_not.asp?idx='+idx+'&menupos=<%= menupos %>','code','width=1400,height=800,scrollbars=yes,resizable=yes');
	}

	popmailCodeView.focus();
}

function goPage(pg) {
	document.frm.page.value=pg;
	document.frm.submit();
}
function reservationOK(idx, saveHTML){
	//alert(saveHTML);
	if(confirm("���� Ȯ���� ���� �ϼ̽��ϱ�??")){
		FrameCKP.location.href='/admin/mailzine/mailzine_siteyn.asp?idx='+idx+'&reservationOK=OK&saveHtml=' + saveHTML+'&menupos=<%= menupos %>';
	}
}

function jsMailzineCode() {
	var winCodeView = window.open('/admin/mailzine/code/PopManageCode.asp','codeview','width=1400,height=800,scrollbars=yes,resizable=yes');
	winCodeView.focus();
}

function jsMailzineTemplate() {
	var winTemplateView = window.open('/admin/mailzine/code/PopManageTemplate.asp','templateview','width=1400,height=800,scrollbars=yes,resizable=yes');
	winTemplateView.focus();
}

// ��ü ����,���
function chgSel_on_off(){
	var frm = document.monthly;
	if (frm.lineSel.length){
		for(var i=0;i<frm.lineSel.length;i++)
		{
			frm.lineSel[i].checked=frm.tt_sel.checked;
		}
	}else{
		frm.lineSel.checked=frm.tt_sel.checked;
	}
}

// ���õ� �׸� ����/����
function fnDeleteMail(){
	var i, chk=0;
	var frm = document.monthly;


	if (frm.lineSel.length){
		for(i=0;i<frm.lineSel.length;i++)
		{
			if(frm.lineSel[i].checked)
			{
				chk++;
			}
		}
	}else{
			if(frm.lineSel.checked)
			{
				chk++;
			}
	}

	if(chk==0){
		alert("�� �� �̻� �������ֽʽÿ�.");
		return;
	}else{
		if(confirm("�����Ͻ� " + chk + "����  �׸��� ���� �Ͻðڽ��ϱ�?")){
			frm.mode.value="delete";
			frm.target="FrameCKP";
			frm.action="mailzine_siteyn.asp";
			frm.submit();
		}else{
			return;
		}
	}
}

</script>

<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mailergubun" value="<%= mailergubun %>">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="#EEEEEE">�˻�<br>����</td>
	<td align="left">
		* ���Ϸ����� : <%= mailergubun %>
		&nbsp;&nbsp;
		* �߼۱Ⱓ :
        <input id="sDt" name="sDt" value="<%=sDt%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
        <input id="eDt" name="eDt" value="<%=eDt%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="eDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		&nbsp;&nbsp;
		* ���� :
		<input type="text" class="text" name="SearchKey" value="<%=SearchKey%>" size="20">
		&nbsp;&nbsp;
		* ���⿩�� :
		<% Drawisusing "isusing" , isusing , "class='select'" %>
		&nbsp;&nbsp;
		* �߼����� :
		<% Drawareagubun "area" , area , "class='select'" %>
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "sDt", trigger    : "sDt_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "eDt", trigger    : "eDt_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
	<td width="50" bgcolor="#EEEEEE">
		<input type="button" class="button_s" value="�˻�" onClick="goPage('');">
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<br>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<% mailzine_member_count %><Br><br><% mailzine_notmember_count %>
	</td>
	<td align="right"></td>
</tr>
<tr>
	<td align="left">
		<input type="button" value="���û���" onclick="fnDeleteMail();" class="button">
	</td>
	<td align="right">
		<input type="button" value="�űԵ��(���ø�)" onclick="jsModifyNewMailzine(-1);" class="button">
		<input type="button" value="�űԵ��(����)" onclick="editreg('');" class="button">
		<!--<input type="button" value="�űԵ��(�ڵ�)" onclick="jsModifyMailzine(-1);" class="button">-->
		&nbsp;
		<input type="button" value="�ڵ����" onclick="jsMailzineCode();" class="button">
		<input type="button" value="���ø�����" onclick="jsMailzineTemplate();" class="button">

		<% if C_ADMIN_AUTH then %>
			<Br>�����ڱ���:
			(<input type="button" class="button" value="�̸��ϱ⺻������" onclick="mailCodeView('','basicMailFormCopy');">
			<input type="button" value="������Ʈ���ۼ�(������)" onclick="blackListReg();" class="button">)
		<% end if %>
	</td>
</tr>
</table>
<!-- �׼� �� -->

<form method="post" name="monthly" style="margin:0px;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type="hidden" name="mode" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= omail.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= omail.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=20><input type="checkbox" name="tt_sel" onclick="chgSel_on_off()"></td>
	<td width=60>No</td>
	<td width=60>�߼���</td>
	<td>Title</td>
	<td width=90>�ۼ�����</td>
	<td width=50>������<Br>�ϼ�����</td>
	<td width=80>����<Br>������</td>
	<td width=80>����Ϸ�ð�</td>
	<td width=40>����Ʈ<Br>����</td>
	<td width=100>�߼�����</td>
	<td width=80>�߼�ȸ�����</td>
	<td width=40>���Ϸ�</td>
	<td width=40>���</td>
	<td width=220>�ڵ�����</td>
</tr>
<% if omail.FresultCount>0 then %>
<% for ix=0 to omail.FresultCount-1 %>
<tr align="center" <% if omail.FItemList(ix).farea="ten_china" then %>bgcolor="<%= adminColor("dgray") %>"<% elseif (isnull(omail.FItemList(ix).FreservationDATE) <> False ) AND (omail.FItemList(ix).farea="finger_all") Then %>bgcolor="<%= adminColor("pink") %>"<% elseif (isnull(omail.FItemList(ix).FreservationDATE) <> False ) AND (omail.FItemList(ix).farea="ten_all") Then %>bgcolor="<%= adminColor("green") %>"<% else %>bgcolor="#FFFFFF"  onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background="#FFFFFF";<% end if %> >
	<td><input type="checkbox" name="lineSel" value="<% = omail.FItemList(ix).Fidx %>"></td>
	<td><% = omail.FItemList(ix).Fidx %></td>
	<td><% = omail.FItemList(ix).Fregdate %></td>
	<td align="left">(����) <% = omail.FItemList(ix).Ftitle %></td>
	<td>
		<% if omail.FItemList(ix).Fregtype2="0" then %>
			<%= omail.FItemList(ix).GetRegTypeName() %>
		<% else %>
			<%= GetRegNewTypeName(omail.FItemList(ix).Fregtype2) %>
		<% end if %>
	</td>
	<td>
		<% if omail.FItemList(ix).Fgubun = "5" then %>
			�ϼ�
		<% else %>
			�̿ϼ�
		<% end if %>
	</td>
	<td>
		<%= left(omail.FItemList(ix).Flastupdate,10) %>
		<br><%= mid(omail.FItemList(ix).Flastupdate,11,12) %>
	</td>
	<td>
		<% If (C_ADMIN_AUTH or C_SYSTEM_Part or C_MD or C_MKT_Part) AND (omail.FItemList(ix).Fgubun = "5") AND (isnull(omail.FItemList(ix).FreservationDATE) <> False ) Then %>
			<input type="button" value="�߼�" onclick="javascript:reservationOK('<% = omail.FItemList(ix).Fidx %>', '<%= CHKIIF(omail.FItemList(ix).Fregtype <> "1", "Y", "N")%>');" class="button"><br>
		<% End If %>

		<%= Chkiif(isnull(omail.FItemList(ix).FreservationDATE), "���� ��", left(omail.FItemList(ix).FreservationDATE,10)&"<br>"&mid(omail.FItemList(ix).FreservationDATE,11,12)) %>
	</td>
	<td>
		<% = omail.FItemList(ix).Fisusing %>
		<Br>
		<input type="button" value="����" class="button" onclick="siteyn('<% = omail.FItemList(ix).Fidx %>','<% = omail.FItemList(ix).Fisusing %>', '<%= omail.FItemList(ix).Fgubun %>');">
	</td>
	<td>
		<%= getareagubun(omail.FItemList(ix).farea) %>
	</td>
	<td>
		<% = omail.FItemList(ix).fmemgubun %>
	</td>
	<td><% = omail.FItemList(ix).fmailergubun %></td>
	<td>
		<% if omail.FItemList(ix).Fregtype2<>"0" then %>
			<input type="button" value="����" class="button" onclick="jsModifyNewMailzine(<% = omail.FItemList(ix).Fidx %>);">
		<% elseif omail.FItemList(ix).Fregtype2="0" and omail.FItemList(ix).Fregtype <> "1" then %>
			<input type="button" value="����" class="button" onclick="jsModifyMailzine(<% = omail.FItemList(ix).Fidx %>);">
		<% else %>
			<input type="button" value="����" class="button" onclick="editreg(<% = omail.FItemList(ix).Fidx %>);">
		<% end if %>
	</td>
	<td>
		<%
		' ���� ���ø� ����
		if omail.FItemList(ix).Fregtype2<>"0" then
		%>
			<input type="button" value="�̸�����(ȸ��)" class="button" onclick="displayTemplates(<% = omail.FItemList(ix).Fidx %>,'member', 'view');">
			<input type="button" value="�̸�����(��ȸ��)" class="button" onclick="displayTemplates(<% = omail.FItemList(ix).Fidx %>,'notmember', 'view');">
			<Br>
			<input type="button" value="�ڵ�(ȸ��/��ȸ��)" class="button" onclick="displayTemplates(<% = omail.FItemList(ix).Fidx %>,'member', 'code');">
			<input type="button" value="�ڵ�(�׽�Ʈ�߼�)" class="button" onclick="displayTemplates(<% = omail.FItemList(ix).Fidx %>,'test', 'code');" <% if mailergubun<>"TMS" then response.write " disabled" %>>
		<%
		' �ڵ�����
		elseif ((omail.FItemList(ix).Fregtype2="0") or omail.FItemList(ix).Fregtype2="") and omail.FItemList(ix).Fregtype <> "1" then
		%>
			<input type="button" value="�̸�����(ȸ��)" class="button" onclick="displayNew(<% = omail.FItemList(ix).Fidx %>,'member', 'view');">
			<input type="button" value="�̸�����(��ȸ��)" class="button" onclick="displayNew(<% = omail.FItemList(ix).Fidx %>,'notmember', 'view');">
			<Br>
			<input type="button" value="�ڵ�(ȸ��)" class="button" onclick="displayNew(<% = omail.FItemList(ix).Fidx %>,'member', 'code');">
			<input type="button" value="�ڵ�(��ȸ��)" class="button" onclick="displayNew(<% = omail.FItemList(ix).Fidx %>,'notmember', 'code');">
		<%
		' �����ϸ���
		else
		%>
			<input type="button" value="�̸�����(ȸ��)" class="button" onclick="displayManual(<% = omail.FItemList(ix).Fidx %>,'member');">
			<input type="button" value="�̸�����(��ȸ��)" class="button" onclick="displayManual(<% = omail.FItemList(ix).Fidx %>,'notmember');">
			<Br>
			<input type="button" value="�ڵ�(ȸ��)" class="button" onclick="mailCodeView(<% = omail.FItemList(ix).Fidx %>,'member');">
			<input type="button" value="�ڵ�(��ȸ��)" class="button" onclick="mailCodeView(<% = omail.FItemList(ix).Fidx %>,'notmember');">
		<% end if %>
	</td>
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
       	<% if omail.HasPreScroll then %>
			<span class="list_link"><a href="javascript:goPage(<%= omail.StarScrollPage-1 %>)">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for ix = 0 + omail.StarScrollPage to omail.StarScrollPage + omail.FScrollCount - 1 %>
			<% if (ix > omail.FTotalpage) then Exit for %>
			<% if CStr(ix) = CStr(omail.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= ix %></b></font></span>
			<% else %>
			<a href="javascript:goPage(<%=ix%>)" class="list_link"><font color="#000000"><%= ix %></font></a>
			<% end if %>
		<% next %>
		<% if omail.HasNextScroll then %>
			<span class="list_link"><a href="javascript:goPage(<%=ix%>)">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>
</form>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe name="FrameCKP" src="" frameborder="0" width="100%" height="400" ></iframe>
<% else %>
	<iframe name="FrameCKP" src="" frameborder="0" width="0" height="0" ></iframe>
<% end if %>

<%
set omail = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbTMSclose.asp" -->