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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/mailzinecls.asp"-->
<%
CONST MAXHeightPX = 1400    '''�� ��ġ�� ���ؼ��� Ȯ������ ����.. (2,000px ���� ���� ������ 2���� �־����� �� ������찡 ����)

dim idx, mode, mailergubun
	idx = requestCheckVar(request("idx"),32)
	mailergubun = requestcheckvar(request("mailergubun"),16)

if (idx = "") then
	idx = -1
end if

if (idx > 0) then
	mode = "modi"
else
	mode = "ins"
end if

if mailergubun="" or isnull(mailergubun) then
	response.write "���Ϸ� ������ �����ϴ�."
	dbget.close() : response.end
end if

dim omail
set omail = new CMailzineList
	omail.frectidx = idx
	''omail.FRectRegType = "2"
	omail.frectmailergubun = "EMS"
	omail.MailzineDetail()

if (omail.FOneItem.Fregtype = "") then
	omail.FOneItem.Fregtype = "2"
end if

%>
<style>
#mask {
	position:absolute;
	z-index:9000;
	background-color:#000;
	display:none;
	left:0;
	top:0;
}
.window{
	display: none;
	position:absolute;
	left:100px;
	bottom:10px;
	z-index:10000;
}
</style>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/JavaScript">

function jsSubmit(frm) {
	var regtype = jsGetRegType();

	if (frm.title.value == "") {
		alert('������ �Է��ϼ���.');
		return;
	}

	if (frm.regdate.value == "") {
		alert('���Ϲ߼� �������� �Է��ϼ���.');
		return;
	}

	if (frm.area.value == "") {
		alert('�߼������� �Է��ϼ���.');
		return;
	}

	if (frm.memgubun.value == "") {
		alert('ȸ������� �Է��ϼ���.');
		return;
	}

	if (frm.isusing.value == "") {
		alert('����Ʈ ���⿩�θ� �Է��ϼ���.');
		return;
	}

	if (frm.secretGubun.value == "") {
		alert('��ũ�������� �Է��ϼ���.');
		return;
	}

	if (regtype == '2') {
		if (frm.evt_code.value == "") {
			alert('�ָ�Ư�� �̺�Ʈ�ڵ带 �Է��ϼ���.');
			return;
		}

		if (frm.evt_code.value*0 != 0) {
			alert('�߸��� �ָ�Ư�� �̺�Ʈ�ڵ��Դϴ�.');
			return;
		}

		if (frm.img1editname.value == "") {
			alert('�������� ��ư�� ��������.');
			return;
		}
	}

	if (frm.mode.value == "ins") {
		// �ű� ��Ͻ�
		if (frm.isusing.value == "Y") {
			alert('�ű� ��Ͻÿ��� ����Ʈ ������ ������ �� �����ϴ�.');
			return;
		}
		if (frm.gubun.value == "5") {
			alert('�ű� ��Ͻÿ��� ������ �ۼ����¸� �Ϸ�� �� �� �����ϴ�.');
			return;
		}
	}

	if (confirm('�����Ͻðڽ��ϱ�?') == true) {
		frm.submit();
	}

}

function jsGetRegType() {
	return $('input:radio[name="regtype"]:checked').val();
}

function jsSetDisabledObj(obj, disabled) {
	obj.disabled = disabled;
	if (obj.type != 'textarea') {
		obj.style.background = disabled ? '#DDDDDD' : '#FFFFFF';
	}
}

function jsSetItemState() {
	var frm = document.frm;
	var regtype = jsGetRegType();
	if (regtype == undefined) { return; }

	jsSetDisabledObj(frm.evt_code, false);

	if (regtype == '2') {
		jsSetDisabledObj(frm.img2editname, false);
	} else if (regtype == '3') {
		jsSetDisabledObj(frm.img2editname, true);
	} else if (regtype == '4') {
		jsSetDisabledObj(frm.img2editname, false);

	// ���̾���丮
	} else if (regtype == '5') {
		jsSetDisabledObj(frm.evt_code, true);
		jsSetDisabledObj(frm.img1editname, false);
		jsSetDisabledObj(frm.img2editname, false);
	}

	jsSetDisabledObj(frm.img3editname, false);
	jsSetDisabledObj(frm.img4editname, false);
}

function jsGetList() {
	var frm = document.frm;
	var regtype = jsGetRegType();

	if (frm.regdate.value == "") {
		alert('���Ϲ߼� �������� �Է��ϼ���.');
		return;
	}

	if (regtype!='5'){
		if (frm.evt_code.value == "") {
			alert('���� �̺�Ʈ�ڵ带 �Է��ϼ���.');
			return;
		}
		if (frm.evt_code.value*0 != 0) {
			alert('�߸��� ���� �̺�Ʈ�ڵ��Դϴ�.');
			return;
		}
	}

	document.iframe_proc.location.href = '/admin/mailzine/mailzine_process.asp?mode=getlist&regdate=' + frm.regdate.value + '&regtype=' + regtype + '&evt_code=' + frm.evt_code.value;
}

$(document).ready(function(){
	jsSetItemState();
});

</script>

<form name="frm" method="post" action="/admin/mailzine/mailzine_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="mailergubun" value="<%= mailergubun %>">
<table width="95%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150">������ ����</td>
	<td><input type="text" name="title" class="input" size="55" value="<%= omail.FOneItem.Ftitle %>" /> * ���� : ${EMS_M_NAME}</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150"><b>������ �߼ۿ�����</b></td>
	<td>
		<input id="regdate" name="regdate" value="<%= omail.FOneItem.Fregdate %>" class="text_ro" size="10" maxlength="10" readonly /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="regdate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
		var regdate = new Calendar({
			inputField : "regdate", trigger    : "regdate_trigger",
			onSelect: function() {
				this.hide();
			}, bottomBar: true, dateFormat: "%Y.%m.%d", fdow: 0
		});
		</script>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150"><b>������ ����</b></td>
	<td>
		<input type="radio" name="regtype" value="2" <% if omail.FOneItem.Fregtype = "2" then response.write "checked"%> onClick="jsSetItemState()" <%= CHKIIF(mode="modi" and omail.FOneItem.Fregtype <> "2", "disabled", "") %>> �ָ�Ư��
		&nbsp;
		<input type="radio" name="regtype" value="3" <% if omail.FOneItem.Fregtype = "3" then response.write "checked"%> onClick="jsSetItemState()" <%= CHKIIF((mode="modi" and omail.FOneItem.Fregtype <> "3"), "disabled", "") %>> ��ȹ��
		&nbsp;
		<input type="radio" name="regtype" value="4" <% if omail.FOneItem.Fregtype = "4" then response.write "checked"%> onClick="jsSetItemState()" <%= CHKIIF((mode="modi" and omail.FOneItem.Fregtype <> "4"), "disabled", "") %>> ��ȹ��+MD's Pick
		&nbsp;
		<input type="radio" name="regtype" value="5" <% if omail.FOneItem.Fregtype = "5" then response.write "checked"%> onClick="jsSetItemState()" <%= CHKIIF((mode="modi" and omail.FOneItem.Fregtype <> "5"), "disabled", "") %>> ���̾���丮
		&nbsp;
		<input type="button" class="button" value="��������" onClick="jsGetList()">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">�߼�����</td>
	<td>
		<% Drawareagubun "area" , omail.FOneItem.Farea , "class='select'" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">�߼�ȸ�����</td>
	<td>
		<% DrawMemberGubun "memgubun" , omail.FOneItem.Fmemgubun , "class='select'" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">����Ʈ����</td>
	<td>
		<% Drawisusing "isusing" , omail.FOneItem.Fisusing , "class='select'" %> * ��� ������ ����˴ϴ�.
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">��ũ�� ����</td>
	<td>
		<% DrawsecretGubun "secretGubun" , omail.FOneItem.FsecretGubun , "class='select'" %> * ����Ʈ�����, ��ũ�� ������ Y�� �θ� Ÿ��Ʋ�� ����ǰ� Ŭ���� ���� �ʽ��ϴ�.
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150"><b>������ �ۼ�����</b></td>
	<td>
		<select name="gubun" class="select">
			<option value="1" <% if omail.FOneItem.Fgubun = "1" then response.write "selected"%>>�̿ϼ�</option>
			<option value="5" <% if omail.FOneItem.Fgubun = "5" then response.write "selected"%>>�ϼ�</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="1">
	<td align="center" width="150"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150">���� �̺�Ʈ�ڵ�</td>
	<td>
		<input type="text" name="evt_code" class="input" size="12" value="<%= omail.FOneItem.Fevt_code %>"> * �ָ�Ư�� �Ǵ� ���� �̺�Ʈ�ڵ�
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150">��ȹ�� �̺�Ʈ�ڵ���</td>
	<td>
		<textarea class="textarea" cols="20" rows="6" name="img1editname"><%= omail.FOneItem.Fimgmap1 %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150">MD Pick ��ǰ���</td>
	<td>
		<textarea class="textarea" cols="20" rows="6" name="img2editname"><%= omail.FOneItem.Fimgmap2 %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150">JUST 1DAY</td>
	<td>
		<input type="text" class="input" name="img3editname" size="20" value="<%= omail.FOneItem.Fimgmap3 %>"  readonly />
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150">�ٹ����� Ŭ����</td>
	<td>
		<textarea class="textarea" cols="20" rows="3" name="img4editname" readonly><%= omail.FOneItem.Fimgmap4 %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="1">
	<td align="center" width="150"></td>
	<td></td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150">�߼ۿ��� �Ϸ���</td>
	<td>
		<%= omail.FOneItem.FreservationDATE %> * ����Ϸ� ���Ŀ��� [��ǰ���]�̽��񿡰� �����ؾ߸� ���������� �ݿ��˴ϴ�.
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td align="center" width="150">�ۼ���(��������)</td>
	<td>
		<%= CHKIIF(IsNull(omail.FOneItem.Fmodiuserid), omail.FOneItem.Freguserid, omail.FOneItem.Fmodiuserid) %>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="50">
	<td align="center" colspan="2">
		<input type="button" class="button" value="����" onClick="jsSubmit(document.frm);">
	</td>
</tr>
</table>
</form>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe name="iframe_proc" width="100%" height="400" frameborder="0"></iframe>
<% else %>
	<iframe name="iframe_proc" width="0" height="0" frameborder="0"></iframe>
<% end if %>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
