<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/hitchhikerCls.asp"-->
<%
	Dim idx, hitchmodi
	Dim mHvol, Evt_code, mEvt_code, mStartdate, mEnddate, mIsusing, mdelidate
	idx = Request("idx")
	
	If idx <> "" Then
		Set hitchmodi = new viphitchhker
			hitchmodi.FIdx = idx
			hitchmodi.fnhitchmodify

			mHvol		= hitchmodi.FOneItem.FHvol
			Evt_code	= hitchmodi.FOneItem.Fevt_code
			mEvt_code	= hitchmodi.FOneItem.Fmevt_code
			mStartdate	= hitchmodi.FOneItem.Fstartdate
			mEnddate	= hitchmodi.FOneItem.Fenddate
			mdelidate	= hitchmodi.FOneItem.Fdelidate
			mIsusing	= hitchmodi.FOneItem.Fisusing
		Set hitchmodi = nothing
	End If
%>

<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<script language='javascript'>
function form_check(){
	var frm = document.frm;
	if(frm.Hvol.value == ''){
		alert('ȸ���� �Է��ϼ���');
		frm.Hvol.focus();
		return false;
	}
	if(frm.evt_code.value == ''){
		alert('�̺�Ʈ�ڵ带 �Է��ϼ���');
		frm.evt_code.focus();
		return false;
	}
	if(frm.startdate.value == ''){
		alert('�������� �Է��ϼ���');
		frm.startdate.focus();
		return false;
	}
	if(frm.enddate.value == ''){
		alert('�������� �Է��ϼ���');
		frm.enddate.focus();
		return false;
	}
	if(frm.delidate.value == ''){
		alert('����� �Է��ϼ���');
		frm.delidate.focus();
		return false;
	}
	<%If idx <> "" Then%>
	var chk = 0;
	for(var j=0; j<frm.isusing.length; j++) {
		if(frm.isusing[j].checked) chk++;
	}
	if (chk < 1){
		alert("��뿩�ο� üũ�ϼ���");
		return false;
	}
	<% End If %>
	frm.submit();
}
</script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<form name="frm" action="hitchhiker_proc.asp" method="post">
<%If idx <> "" Then%>
<input type="hidden" name="idx" value="<%=idx%>">
<% End If %>
<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#E6E6E6" height="30">
	<td width="100">ȸ��</td>
	<td width="280" bgcolor="#FFFFFF" align="left">Vol.<input type="text" name="Hvol" value="<%=mHvol%>" size="3" maxlength="3"></td>
</tr>
<tr align="center" bgcolor="#E6E6E6" height="30">
	<td width="100">web�̺�Ʈ�ڵ�</td>
	<td width="280" bgcolor="#FFFFFF" align="left"><input type="text" name="evt_code" value="<%=Evt_code%>" size="10" maxlength="10"></td>
</tr>
<tr align="center" bgcolor="#E6E6E6" height="30">
	<td width="100">mobile�̺�Ʈ�ڵ�</td>
	<td width="280" bgcolor="#FFFFFF" align="left"><input type="text" name="m_evt_code" value="<%=mEvt_code%>" size="10" maxlength="10"></td>
</tr>
<tr align="center" bgcolor="#E6E6E6" height="30">
	<td>������ ~ ������</td>
	<td bgcolor="#FFFFFF" align="left">
        <input id="startdate" name="startdate" value="<%=mStartdate%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
        <input id="enddate" name="enddate" value="<%=left(mEnddate,10)%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var ENT_Start = new Calendar({
				inputField : "startdate", trigger    : "startdate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					ENT_End.args.min = date;
					ENT_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var ENT_End = new Calendar({
				inputField : "enddate", trigger    : "enddate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					ENT_Start.args.max = date;
					ENT_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	<br><font color="red">�� �������� 23:59:59�� �ڵ� �����˴ϴ�.</font>
	</td>
</tr>
<tr align="center" bgcolor="#E6E6E6">
	<td>�����</td>
	<td bgcolor="#FFFFFF" align="left">
        <input id="delidate" name="delidate" value="<%=mdelidate%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="delidate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var ENT_deli = new Calendar({
				inputField : "delidate", trigger    : "delidate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					ENT_End.args.min = date;
					ENT_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
</tr>
<%If idx <> "" Then%>
<tr align="center" bgcolor="#E6E6E6" height="30">
	<td width="100">��뿩��</td>
	<td width="280" bgcolor="#FFFFFF" align="left">
		<input type="radio" name="isusing" value="Y" <%=chkiif(mIsusing = "Y","checked","")%>>���&nbsp;&nbsp;&nbsp;
		<input type="radio" name="isusing" value="N" <%=chkiif(mIsusing = "N","checked","")%>>������
	</td>
</tr>
<% End If %>
</table>
<table width="380" cellpadding="0" cellspacing="0">
<tr height="30">
	<td align="right"><img src="http://testwebadmin.10x10.co.kr/images/icon_save.gif" border="0" style="cursor:pointer" onClick="form_check();"></td>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->