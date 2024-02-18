<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/event/eventCls.asp"-->
<%
Dim idx, oEvent
Dim gubun, actid, company_name, evt_startdate, evt_enddate, contentsCode, evt_name, isusing
Dim makerid
idx 	= requestCheckvar(request("idx"),10)
gubun	= requestCheckvar(request("gubun"),1)
makerid	= session("ssBctId")

If idx <> "" Then
	Set oEvent = new CEvent
		oEvent.FRectIdx = idx
		oEvent.getEventOneItem

		actid			= oEvent.FOneItem.FActid
		company_name	= oEvent.FOneItem.FCompany_name
		evt_startdate	= oEvent.FOneItem.FEvt_startdate
		evt_enddate		= oEvent.FOneItem.FEvt_enddate
		contentsCode	= oEvent.FOneItem.FContentsCode
		evt_name		= oEvent.FOneItem.FEvt_name
		isusing			=  oEvent.FOneItem.FIsusing
	Set oEvent = nothing
End If

%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script>
//���� ��ü���� �˾�
function pop_lecture(v){
    var popwin = window.open("/academy/events/pop_lecturerList.asp?gubun="+v+"","pop_lecture","width=500,height=700,scrollbars=yes,resizable=yes");
	popwin.focus();
}
//�������� ���º��� �˾�
function pop_lec(){
    var popwin2 = window.open("/lectureadmin/events/pop_lecList.asp","pop_lec","width=900,height=700,scrollbars=yes,resizable=yes");
	popwin2.focus();
}
//�Ǹ����� ��ǰ���� �˾�
function pop_art(){
    var popwin3 = window.open("/lectureadmin/events/pop_artList.asp","pop_art","width=900,height=700,scrollbars=yes,resizable=yes");
	popwin3.focus();
}
function frm_check(f){
	if($("#evt_startdate").val() == '' ){
		alert('�̺�Ʈ �������� �Է��ϼ���');
		return false;
	}

	if($("#evt_enddate").val() == '' ){
		alert('�̺�Ʈ �������� �Է��ϼ���');
		return false;
	}

	if( $("#evt_startdate").val() > $("#evt_enddate").val() ) {
		alert("�������� �����Ϻ��� �����ϴ�. �ٽ� �Է����ּ���");
		return false;
	}

	if("<%= gubun %>" == 'D'){
		if($("#diycode").val() == ''){
			alert('��ǰ�ڵ带 �Է��ϼ���');
			return false;
		}
	}else if("<%= gubun %>" == 'L'){
		if($("#lecidx").val() == ''){
			alert('�����ڵ带 �Է��ϼ���');
			return false;
		}
	}

	if($("#evt_startdate").val() == $("#evt_enddate").val()){
		alert('���� �Ⱓ�� ���� �Ұ��� �մϴ�');
		$("#evt_enddate").val('');
		$("#evt_enddate").focus();
		return false;
	}

	if($("#evt_name").val() == '' ){
		alert('�̺�Ʈ���� �Է��ϼ���');
		return false;
	}

	if(confirm('���� �Ͻðڽ��ϱ�?')){
		frm.action = '/lectureadmin/events/event_process.asp';
	<% If idx <> "" Then %>
		frm.mode.value = 'U';
	<% Else %>
		frm.mode.value = 'I';
	<% End If %>
		frm.submit();
	}
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="50" bgcolor="FFFFFF">
	<td>
	    <font size="2"><strong>*<%= Chkiif(idx <> "", "����", "�ű� ���") %></strong></font>
	    <br>- ����� �۰��������� ���� ���·� ǥ�õ˴ϴ�. (���� ��� ����˴ϴ�)  <a href="http://m.thefingers.co.kr/corner/lectureDetail.asp?lecturer_id=<%=session("ssBctID")%>" target="_blank"><font color='#0000FF'>[����� �۰������� ����]</font></a>
        <br>- ���� �Ⱓ�� �ϳ��� ��ũ�� �����մϴ�.
	</td>
</tr>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="gubun" value="<%= gubun %>">
<input type="hidden" name="idx" value="<%= idx %>">
<col width="20%" />
<col width="" />
<tr height="30" bgcolor="FFFFFF">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�Ⱓ</td>
	<td>
		<input id="evt_startdate" readonly name="evt_startdate" class="text" value="<%= evt_startdate %>" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="evt_startdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> 00:00:00 ~
		<input id="evt_enddate" readonly name="evt_enddate" class="text" value="<%= evt_enddate %>" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="evt_enddate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> 00:00:00
		<script language="javascript">
		var CAL_Start = new Calendar({
			inputField : "evt_startdate", trigger    : "evt_startdate_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_End.args.min = date;
				CAL_End.redraw();
				this.hide();
			}, bottomBar: true, dateFormat: "%Y-%m-%d"
		});
		var CAL_End = new Calendar({
			inputField : "evt_enddate", trigger    : "evt_enddate_trigger",
			onSelect: function() {
				var date = Calendar.intToDate(this.selection.get());
				CAL_Start.args.max = date;
				CAL_Start.redraw();
				this.hide();
			}, bottomBar: true, dateFormat: "%Y-%m-%d"
		});
		</script>
	</td>
</tr>
<tr height="30" bgcolor="FFFFFF">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�����ġ</td>
	<td><%= chkiif(gubun="D", "�۰� ������(��ǰ)", "���� ������(����)") %></td>
</tr>
<% If gubun = "D" Then %>
<tr height="30" bgcolor="FFFFFF">
	<td align="center" bgcolor="#E6E6E6">�۰�</td>
	<td><%= makerid %></td>
</tr>
<tr height="30" bgcolor="FFFFFF">
	<td align="center" bgcolor="#E6E6E6">��ǰ�ڵ�</td>
	<td>
		<input type="text" name="diycode" id="diycode" class="text" readonly value="<%= contentsCode %>">
		<input type="button" value="�Ǹ����� ��ǰ����" class="button" id="btnDiyView" onclick="pop_art();">
	</td>
</tr>
<% ElseIf gubun = "L" Then %>
<tr height="30" bgcolor="FFFFFF">
	<td align="center" bgcolor="#E6E6E6">����</td>
	<td><%= makerid %></td>
</tr>
<tr height="30" bgcolor="FFFFFF">
	<td align="center" bgcolor="#E6E6E6">�����ڵ�</td>
	<td>
		<input type="text" name="lecidx" class="text" id="lecidx" readonly value="<%= contentsCode %>">
		<input type="button" value="�������� ���º���" class="button" id="btnView" onclick="pop_lec();">
	</td>
</tr>
<% End If %>
<tr height="30" bgcolor="FFFFFF">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ��</td>
	<td><input type="text" class="text" name="evt_name" id="evt_name" value="<%= evt_name %>" size="50"></td>
</tr>
<tr height="30" bgcolor="FFFFFF">
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">��뿩��</td>
	<td>
		<label><input type="radio" name="isusing" value="Y" class="radio" <% If isusing="" OR isusing="Y" Then response.write "checked" %>>Y</label>
		<label><input type="radio" name="isusing" value="N" class="radio" <%= chkiif(isusing="N", "checked", "") %> >N</label>
	</td>
</tr>
<tr height="30" bgcolor="FFFFFF" align="center">
	<td colspan="2">
		<input type="button" class="button" value="����" onclick="frm_check(this.frm);" style="color:red;font-weight:bold">
		<input type="button" class="button" value="���" onclick="history.back(-1);">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->