<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ����Ȯ���� ����� �޴�����ȣ ���� �˾�
' History : 2011.05.30 ������ ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%
dim cMember
dim userid
dim empno, susername, sjuminno, susercell, hp1, hp2, hp3

empno = session("ssBctSn")

'// ���� �⺻���� ����
Set cMember = new CTenByTenMember
	cMember.Fempno = empno
	cMember.fnGetMemberData

	empno   		= cMember.Fempno
	susername      	= cMember.Fusername
	susercell      	= cMember.Fusercell

Set cMember = Nothing

if empno="" or isNull(empno) then
	Call Alert_close("���� ������ �����ϴ�.\n�����ڿ��� ���ǿ��")
	response.End
end if

'//�޴��� ��ȣ �и�
if Not(trim(susercell)="" or isNull(susercell)) then
	susercell = split(susercell,"-")
	if ubound(susercell)>1 then
		hp1 = susercell(0)
		hp2 = susercell(1)
		hp3 = susercell(2)
	end if
end if
%>
<script language='javascript'>
var chkSendAuth = false;

// ������ȣ ����
function chkHPIdentify(){
    var frm = document.frmChkId;

	if(frm.hpNum2.value.length<3){
		alert('�ڵ�����ȣ�� �Է����ּ���');
		frm.hpNum2.focus();
		return ;
	}

	if(frm.hpNum3.value.length<4){
		alert('�ڵ�����ȣ�� �Է����ּ���');
		frm.hpNum3.focus();
		return ;
	}

	if(!chkSendAuth) {
		alert('[������ȣ �ޱ�]�� ������ȣ�� �޾��ּ���.');
		return ;
	}

	frm.target ="hidFrm";
	frm.submit();
}

// SMS ������ȣ �߼�
function popSMSAuthNo(frm) {
	if(frm.hpNum2.value.length<3){
		alert('�ڵ�����ȣ�� �Է����ּ���');
		frm.hpNum2.focus();
		return ;
	}

	if(frm.hpNum3.value.length<4){
		alert('�ڵ�����ȣ�� �Է����ּ���');
		frm.hpNum3.focus();
		return ;
	}
	frm.chgHp.value = frm.hpNum1.value+"-"+frm.hpNum2.value+"-"+frm.hpNum3.value;

	hidFrm.location.href="/tenmember/member/iframe_adminChgHP_SendSMS.asp?eno="+frm.empNo.value+"&chp="+frm.chgHp.value;
	chkSendAuth = true;
}

// SMS�Է� ī���� �۵�(3�а�:180��)
var iSecond=180;
var timerchecker = null;

function startLimitCounter(cflg) {
	if(cflg=="new") {
		if(timerchecker != null) {
			alert("�̹� ������ȣ�� �߼��Ͽ����ϴ�.\n�޴����� SMS�� Ȯ�����ּ���.");
			return;
		} else if(timerchecker == null) {
			document.getElementById("lySMSTime").style.display="";
		}
		iSecond=180;
	}
    rMinute = parseInt(iSecond / 60);
    rSecond = iSecond % 60;
    if(rSecond<10) {rSecond="0"+rSecond};

    if(iSecond > 0) {
        document.forms[0].sLimitTime.value = rMinute+":"+rSecond;
        iSecond--;
        timerchecker = setTimeout("startLimitCounter()", 1000); // 1�� �������� üũ
    } else {
        clearTimeout(timerchecker);
        document.forms[0].sLimitTime.value = "0:00";
        timerchecker = null;
        chkSendAuth = false;
        alert("������ȣ �Է� �ð��� ����Ǿ����ϴ�.\n\nSMS�� ���� ���ߴٸ� �ٽ� ��ȣ�� �޾��ּ���.");
        document.getElementById("lySMSTime").style.display="none";
    }
}
</script>
<form name="frmChkId" method="post" action="doChangeHPIdentify.asp" onsubmit="return false;">
<input type="hidden" name="empNo" value="<%=empno%>">
<input type="hidden" name="chgHp" value="">
<table width="100%" cellpadding="2" cellspacing="1" border="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td colspan="2" bgcolor="#E8F0FF"><b>�޴�����ȣ ���� / ���� Ȯ��</b></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�����ȣ</td>
	<td bgcolor="white"><%=empno%></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�̸�</td>
	<td bgcolor="white"><input type="text" name="username" value="<%=susername%>" readonly class="text_ro" size="10"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�޴�����ȣ</td>
	<td bgcolor="white">
		<select name="hpNum1" class="select">
			<option value="010" <%=chkIIF(hp1="010","checked","")%>>010</option>
			<option value="011" <%=chkIIF(hp1="011","checked","")%>>011</option>
			<option value="016" <%=chkIIF(hp1="016","checked","")%>>016</option>
			<option value="017" <%=chkIIF(hp1="017","checked","")%>>017</option>
			<option value="018" <%=chkIIF(hp1="018","checked","")%>>018</option>
			<option value="019" <%=chkIIF(hp1="019","checked","")%>>019</option>
		</select>-
		<input name="hpNum2" type="text" class="text" size="4" maxlength="4" value="<%=hp2%>">-
		<input name="hpNum3" type="text" class="text" size="4" maxlength="4" value="<%=hp3%>">
		<input type="button" value='������ȣ �ޱ�' onclick="popSMSAuthNo(this.form)" style="padding-top:1px; width:80px; height:20px; border:1px solid #E0E0E0; background-color:#E8F0FF;font-size:11px;">
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">������ȣ</td>
	<td bgcolor="white">
		<input name="authNo" type="text" class="text" size="6" maxlength="6">
	</td>
</tr>
<!-- // SMS������ȣ �Է� ��ȿ�ð� // -->
<tr id="lySMSTime" style="display:none;">
	<td bgcolor="<%= adminColor("tabletop") %>">&nbsp;</td>
    <td align="left" bgcolor="white">
      	�Է� ��ȿ�ð� : <input type=text name="sLimitTime" value="-:--" readolny style="width:40px; border:1px dotted #E0E0E0; text-align:center;background-color:#F8F8F8;">
    </td>
</tr>
<tr>
	<td colspan="2" bgcolor="#FEF8F8">
		�� �Է��� �޴������� �������ڰ� �߼۵Ǹ�, ����Ȯ���� �Ϸ�Ǹ� [�� ����]�� [�޴�����ȣ]�� �����˴ϴ�.<br>
		&nbsp;&nbsp;(����Ȯ���� �޴�����ȣ ���� �� ���� 1ȸ�� Ȯ��)
	</td>
</tr>
<tr>
	<td colspan="2" bgcolor="white" align="center">
		<input type="button" class="button" value="����Ȯ��" onclick="chkHPIdentify()">
		&nbsp;&nbsp;<input type="button" class="button" value=" â�ݱ� " onclick="self.close();">
	</td>
</tr>
</table>
<iframe id="hidFrm" name="hidFrm" src="about:blank" frameborder="0" width="0" height="0"></iframe>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->