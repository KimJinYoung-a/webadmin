<% option Explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �̺�Ʈ ��÷��
' History : 2009.04.17 ���ʻ����� ��
'			2016.06.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function CopyZip(frmname, post1, post2, addr, dong) {
    eval(frmname + ".zipcode").value = post1 + "-" + post2;

    eval(frmname + ".addr1").value = addr;
    eval(frmname + ".addr2").value = dong;
}

function PopSearchZipcode(frmname) {
	var popwin = window.open("/lib/searchzip3.asp?target=" + frmname,"PopSearchZipcode","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}

function jsPopCal(sName){
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function delThis(){
    var frm = document.infoform;

    if (confirm('���� �Ͻðڽ��ϱ�?')){
        if (confirm('������ ���� �Ͻðڽ��ϱ�?')){
            frm.mode.value="del";
    		frm.submit();
		}
	}
}

function gotowrite(){
    var frm = document.infoform;
	if(frm.gubuncd.value == ""){
		alert("������ �������ּ���.");
	    frm.gubuncd.focus();
	    return;
	}

	if(frm.gubunname.value == ""){
		alert("�̺�Ʈ��(���и�)�� �Է����ּ���.");
	    frm.gubunname.focus();
	    return;
	}

	if(frm.prizetitle.value == ""){
		alert("��ǰ���� �Է����ּ���.");
	    frm.prizetitle.focus();
	    return;
	}

	if ((frm.useDefaultAddr[1].checked == true) || (frm.useDefaultAddr[2].checked == true)) {
		if (frm.userid.value == '') {
			alert('���̵� �Էµ� ��쿡�� ���ð����մϴ�.');
			return;
		}
	} else {
		if(frm.username.value == ""){
			alert("��÷�ڼ����� �Է����ּ���.");
			frm.username.focus();
			return;
		}

		if(frm.reqname.value == ""){
			alert("�����ô� ���� �̸��� �Է����ּ���.");
			frm.reqname.focus();
			return;
		}

		if(frm.reqphone1.value == "" || frm.reqphone2.value == "" || frm.reqphone3.value == ""){
			alert("�����ô� ���� ��ȭ��ȣ�� �Է����ּ���.");
			frm.reqphone1.focus();
			return;
		}

		if(frm.reqhp1.value == "" || frm.reqhp2.value == "" || frm.reqhp3.value == ""){
			alert("�����ô� ���� �ڵ��� ��ȣ�� �Է����ּ���.");
			frm.reqphone1.focus();
			return;
		}

		if(frm.zipcode.value == ""){
			alert("�����ô� ���� �ּҸ� �Է����ּ���.");
			frm.zipcode.focus();
			return;
		}

		if(frm.addr2.value == ""){
			alert("�����ô� ���� �������ּҸ� �Է����ּ���.");
			frm.addr2.focus();
			return;
		}
	}

	if (frm.reqdeliverdate.value.length!=10){
	    alert('��� ��û���� �Է��ϼ���.');
	    frm.reqdeliverdate.focus();
	    return;
	}

	if ((!frm.isupchebeasong[0].checked)&&(!frm.isupchebeasong[1].checked)){
	    alert('��� ������ ���� �ϼ���.');
	    frm.isupchebeasong[0].focus();
	    return;
	}
	if(frm.isupchebeasong[1].checked&&(frm.jungsan.checked)&&(frm.jungsanValue.value=="")){
	    alert('�����(���԰�)�� �Է��ϼ���');
	    frm.jungsanValue.focus();
	    return;
	}
	if ((frm.isupchebeasong[1].checked)&&(frm.makerid.value.length<1)){
	    alert('��ü ����� ��� �귣�� ���̵�  ���� �ϼ���.');
	    frm.makerid.focus();
	    return;
	}
	if (confirm('�Է� ������ ��Ȯ�մϱ�?')){
		frm.submit();
	}
}

function disabledBox(comp){
    var frm = comp.form;
    if (comp.value=="Y"){
        frm.makerid.disabled = false;
        frm.jungsan.disabled = false;

		frm.jungsanValue.disabled = false;
        frm.jungsan.checked = true;
    }else{
        frm.makerid.selectedIndex = 0;
        frm.makerid.value = '';
		frm.makerid.disabled = true;
		frm.jungsan.disabled = true;

        frm.jungsanValue.value = '';
        frm.jungsanValue.disabled = true;
        frm.jungsan.checked = false;
    }
}

function disableAddressBox(comp) {
    var frm = comp.form;
    if (comp.value != "C"){
        frm.username.disabled = true;
        frm.reqname.disabled = true;
		frm.reqphone1.disabled = true;
		frm.reqphone2.disabled = true;
		frm.reqphone3.disabled = true;
		frm.reqhp1.disabled = true;
		frm.reqhp2.disabled = true;
		frm.reqhp3.disabled = true;
		frm.zipcode.disabled = true;
		frm.addr2.disabled = true;
    }else{
        frm.username.disabled = false;
        frm.reqname.disabled = false;
		frm.reqphone1.disabled = false;
		frm.reqphone2.disabled = false;
		frm.reqphone3.disabled = false;
		frm.reqhp1.disabled = false;
		frm.reqhp2.disabled = false;
		frm.reqhp3.disabled = false;
		frm.zipcode.disabled = false;
		frm.addr2.disabled = false;
    }

	var evtprize_enddate = $('.evtprize_enddate');
	if (comp.value == "N") {
		evtprize_enddate.show();
	} else {
		evtprize_enddate.hide();
	}
}

function jungsanYN(){
	var frm = document.infoform;
	if(frm.jungsan.checked==true){
		frm.jungsanValue.disabled = false;
	}else{
		frm.jungsanValue.value = '';
		frm.jungsanValue.disabled = true;
	}
}
function checkover1(obj) {
	var val = obj.value;
	if (val) {
		if (val.match(/^\d+$/gi) == null) {
			alert("���ڸ� ��������!");
			document.infoform.jungsanValue.value = '';
			obj.select();
			return;
		}
	}
}
</script>
<table width="100%" border="0" cellpadding="0" cellspacing=0 class="a">
<form name="infoform" method="post" action="/admin/etcsongjang/lib/doeventbeasonginfo.asp">
<input type="hidden" name="mode" value="I">
<tr>
	<td align="center">
		<table width="90%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr height="30">
			<td height="2" colspan="2" >* ��Ÿ��� ���� �Է�</td>
		</tr>
		<tr height="2">
			<td height="2" colspan="2" bgcolor="#AAAAAA"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">����</td>
			<td style="padding-left:7">
				<select name="gubuncd" class="select">
					<option value="">��ü
<!--
					<option value="96">��
					<option value="97">29cm��
-->
					<option value="98">����
					<option value="99">��Ÿ
					<option value="80">CS���
				</select>
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">�̺�Ʈ��(���и�) </td>
			<td style="padding-left:7">
				<input type="text" class="text" name="gubunname" size="40" maxlength="64" value="" >
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">��÷��ǰ</td>
			<td style="padding-left:7">
				<input type="text" class="text" name="prizetitle" size="40" maxlength="64" value="" > * <font color="red">����Ʈ����</font>
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">���̵�</td>
			<td style="padding-left:7">
				<input type="text" class="text" name="userid" size="20" maxlength="32" value="" >
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">����� ��ϱ���</td>
			<td style="padding-left:7">
				<input type=radio name="useDefaultAddr" value="C" checked onClick="disableAddressBox(this)">�����Է�
				<input type=radio name="useDefaultAddr" value="N" onClick="disableAddressBox(this)">User �� ����� �Է�
				<input type=radio name="useDefaultAddr" value="Y" onClick="disableAddressBox(this)">User �⺻ �ּ� ���
			</td>
		</tr>
		<tr height="1" class="evtprize_enddate" style="display: none;">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr class="evtprize_enddate" style="display: none;">
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">������Է� ������</td>
			<td style="padding-left:7">
				<input type="text" class="text_ro" name="evtprize_enddate" size="10" maxlength="10"  value="<%= Left(DateAdd("m", 3, Now()), 10) %>">
				<a href="javascript:jsPopCal('evtprize_enddate');"><img src="/images/calicon.gif" border="0" align="absmiddle"></a>
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">��÷�ڼ���</td>
			<td style="padding-left:7">
				<input type="text" class="text" name="username" size="20" maxlength="20" value="" >
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">�����μ���</td>
			<td style="padding-left:7">
				<input type="text" class="text" name="reqname" size="20" maxlength="20" value="" >
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">����ó</td>
			<td class="verdana_s" style="padding-left:7">
				<input type="text" class="text" name="reqphone1" size="3" class="verdana_s" maxlength="3" value="">
				-
				<input type="text" class="text" name="reqphone2" size="4" class="verdana_s" maxlength="4" value="">
				-
				<input type="text" class="text" name="reqphone3" size="4" class="verdana_s" maxlength="4" value="">
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">�ڵ���</td>
			<td class="verdana_s" style="padding-left:7">
				<input type="text" class="text" name="reqhp1" size="3" class="verdana_s"  maxlength="3" value="">
				-
				<input type="text" class="text" name="reqhp2" size="4" class="verdana_s"  maxlength="4" value="">
				-
				<input type="text" class="text" name="reqhp3" size="4" class="verdana_s"  maxlength="4" value="">
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">������ �ּ�</td>
			<td class="verdana_s" style="padding:5 0 5 7">
				<input type="text" class="text_ro" name="zipcode" size="7" class="verdana_s" readOnly value="">
				<input type="button" class="button" value="�˻�" onClick="FnFindZipNew('infoform','E')">
				<input type="button" class="button" value="�˻�(��)" onClick="TnFindZipNew('infoform','E')">
				<% '<input type="button" value="�˻�(��)" class="button" onclick="PopSearchZipcode('infoform');" onFocus="this.blur();"> %>
				<br>
				<input type="text" class="text_ro" name="addr1" size="16" maxlength="64"  readOnly value="" ><br>
				<input type="text" class="text" name="addr2" size="40" maxlength="64" value="" >
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">��Ÿ��û����</td>
			<td class="verdana_s" style="padding:5 0 5 7"><textarea class="text" name="reqetc" class="textarea" style="width:350px;height:40px;"></textarea></td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">����û��</td>
			<td class="verdana_s" style="padding:5 0 5 7">
			<input type="text" class="text_ro" name="reqdeliverdate" size="10" maxlength="10"  value="" >
			<a href="javascript:jsPopCal('reqdeliverdate');"><img src="/images/calicon.gif" border="0" align="absmiddle"></a>
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">��۱���</td>
			<td style="padding-left:7">
				<input type=radio name=isupchebeasong value="N" checked onClick="disabledBox(this);">�ٹ����ٹ��
				<input type=radio name=isupchebeasong value="Y" onClick="disabledBox(this);">��ü�������
			<br>
			<% drawSelectBoxDesignerwithName "makerid", "" %>
			</td>
		</tr>
		<tr height="1">
			<td height="1" colspan="2" bgcolor="#DDDDDD"></td>
		</tr>
		<script>
			document.infoform.makerid.disabled = true;
		</script>
		<tr>
			<td width="100" height="30" bgcolor="#f7f7f7" style="padding-left:10" class="bbstext">���꿩��</td>
			<td style="padding-left:7">
				<input type="checkbox" class="checkbox" name="jungsan" id="jungsan" onclick="javascript:jungsanYN();" disabled >������&nbsp;&nbsp;
				�����(���԰�) : <input type="text" class="text" id="jungsanValue" name="jungsanValue" value="" onkeyup="checkover1(this)">��
			</td>
		</tr>
		<tr height="2">
			<td height="2" colspan="2" bgcolor="#AAAAAA"></td>
		</tr>
		<tr height="30">
			<td colspan="2" align="center">
			<input type="button" class="button" value=" �� �� " onClick="gotowrite();" onfocus="this.blur();">
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
