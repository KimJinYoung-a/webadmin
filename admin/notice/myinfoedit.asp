<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��Ʈ��� �������� ����
' History : 2007.07.30 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member/10x10staffcls.asp"-->
<%
dim txBirthday1, txBirthday2, txBirthday3, txPhone1, txPhone2 , birth_isSolar
dim txPhone3, txPhone4, txCell1, txCell2, txCell3, txZip1, txZip2

dim onepartner
	set onepartner = new CPartnerUser
	onepartner.GetOnePartner session("ssBctId")

dim birthdayarr,telarr,hparr,ziparr
	if onepartner.Fbirthday <> "" then
		birthdayarr = split(left(onepartner.Fbirthday,10),"-")
			if ubound(birthdayarr) >= 2 then
				txBirthday1 = birthdayarr(0)
				txBirthday2 = birthdayarr(1)
				txBirthday3 = birthdayarr(2)
			end if
	end if
	if onepartner.Ftel <> "" then
		telarr = split(onepartner.Ftel,"-")
			if ubound(telarr) >= 3 then
				txPhone1 = telarr(0)
				txPhone2 = telarr(1)
				txPhone3 = telarr(2)
				txPhone4 = telarr(3)
			end if
	end if
	if onepartner.Fmanager_hp <> "" then
		hparr = split(onepartner.Fmanager_hp,"-")
			if ubound(hparr) >= 2 then
				txCell1 = hparr(0)
				txCell2 = hparr(1)
				txCell3 = hparr(2)
			end if
	end if
	if onepartner.Fzipcode <> "" then
		ziparr = split(onepartner.Fzipcode,"-")
			if ubound(ziparr) >= 1 then
				txZip1  = ziparr(0)
				txZip2  = ziparr(1)
			end if
	end if
	if onepartner.Fzipcode <> "" then
		ziparr = split(onepartner.Fzipcode,"-")
			if ubound(ziparr) >= 1 then
				txZip1  = ziparr(0)
				txZip2  = ziparr(1)
			end if
	end if

	if onepartner.fbirth_isSolar = "Y" or onepartner.fbirth_isSolar = "" or isnull(onepartner.fbirth_isSolar) then
		birth_isSolar = "Y"
	else
		birth_isSolar = "N"
	end if
%>

<script language="javascript">

function TnFindZip(frmname){
	window.open('/lib/searchzip2.asp?target=' + frmname, 'findzipcdode', 'width=460,height=250,left=400,top=200,location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no');
}

function TnJoin10x10(frm){
	if (frm.txBirthday1.value == ''){
		alert("��������� �Է��ϼ���");
		frm.txBirthday1.focus();
	return;
	}
	if (frm.txBirthday2.value == ''){
		alert("��������� �Է��ϼ���");
		frm.txBirthday2.focus();
	return;
	}
	if (frm.txBirthday3.value == ''){
		alert("��������� �Է��ϼ���");
		frm.txBirthday3.focus();
	return;
	}
	if (frm.txPhone1.value == ''){
		alert("��ȭ��ȣ�� �Է��ϼ���");
		frm.txPhone1.focus();
	return;
	}
	if (frm.txPhone2.value == ''){
		alert("��ȭ��ȣ�� �Է��ϼ���");
		frm.txPhone2.focus();
	return;
	}
	if (frm.txPhone3.value == ''){
		alert("��ȭ��ȣ�� �Է��ϼ���");
		frm.txPhone3.focus();
	return;
	}
	if (frm.txPhone4.value == ''){
		alert("��ȭ��ȣ�� �Է��ϼ���");
		frm.txPhone4.focus();
	return;
	}
	if (frm.txCell1.value == ''){
		alert("�ڵ�����ȣ�� �Է��ϼ���");
		frm.txCell1.focus();
	return;
	}
	if (frm.txCell2.value == ''){
		alert("�ڵ�����ȣ�� �Է��ϼ���");
		frm.txCell2.focus();
	return;
	}
	if (frm.txCell3.value == ''){
		alert("�ڵ�����ȣ�� �Է��ϼ���");
		frm.txCell3.focus();
	return;
	}
	if (frm.txZip1.value == ''){
		alert("�����ȣ�� �Է��ϼ���");
		frm.txZip1.focus();
	return;
	}
	if (frm.txZip2.value == ''){
		alert("�����ȣ�� �Է��ϼ���");
		frm.txZip2.focus();
	return;
	}
	if (frm.txAddr1.value == ''){
		alert("�ּҸ� �Է��ϼ���");
		frm.txAddr1.focus();
	return;
	}
	if (frm.txAddr2.value == ''){
		alert("�ּҸ� �Է��ϼ���");
		frm.txAddr2.focus();
	return;
	}
	var ret = confirm('���� �Ͻðڽ��ϱ�?');
	if(ret){
	frm.submit();
	}
}

function passwordedit(frm){
	if (frm.txpass2.value != frm.txpass3.value){
		alert("�����Ͻ� ��й�ȣ�� ��ġ���� �ʽ��ϴ�. �ι� ��Ȯ�� �Է����ֽʽÿ�.");
		frm.txpass2.value=""
		frm.txpass3.value=""
		frm.txpass2.focus();
		return ;
	}

	if (frm.txpass1.value == "" || frm.txpass2.value == "" || frm.txpass3.value == ""){
		alert("������й�ȣ�� �����Һ�й�ȣ�� ��Ȯ�� �Է��� �ּ���.");
		frm.txpass1.value=""
		frm.txpass2.value=""
		frm.txpass3.value=""
		frm.txpass1.focus();
		return ;
	}

	var ret = confirm('��й�ȣ�� ���� �Ͻðڽ��ϱ�?');
		if(ret){
			frm.submit();
		}
	}

</script>

<br><br><br><br><font color=red>���� ���� ������</font><br><br><br><br>

<!--�⺻�������� ����-->
<table width="48%" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="myinfoForm" method="post" action="domyinfo.asp">
	<input type="hidden" name="userid" value="<%= onepartner.Fid %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<font color="red"><strong>�⺻����</strong></font>
		</td>
	</tr>
	<tr align="left" height="25">
    	<td width="120" bgcolor="<%= adminColor("tabletop") %>">�̸�</td>
    	<td bgcolor="#FFFFFF">
    		<input name="txName" id="[on,off,2,16][����]" type="text" size="10" class="text" value="<%= onepartner.Fcompany_name %>">
    		���� ���Խ� ������ ���� �Է� �Ͽ� �ֽʽÿ�.(�� : ȫ�浿)
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">���̵�</td>
    	<td bgcolor="#FFFFFF">
    		<%= onepartner.Fid %>
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">�������</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="txBirthday1" id="[on,on,4,4][�¾��]" size="4" maxlength="4" class="text" value="<%= txBirthday1 %>">��
			<input type="text" name="txBirthday2" id="[on,on,2,2][�¾��]" size="4" maxlength="2" class="text" value="<%= txBirthday2 %>">��
			<input type="text" name="txBirthday3" id="[on,on,2,2][�¾��]" size="4" maxlength="2" class="text" value="<%= txBirthday3 %>">��
			&nbsp; &nbsp; &nbsp; &nbsp; ���:<input type="radio" name="birth_isSolar" value="Y" <% if birth_isSolar = "Y" then response.write " checked" %>>
			����:<input type="radio" name="birth_isSolar" value="N" <% if birth_isSolar = "N" then response.write " checked" %>>
    	</td>
    </tr>
	<tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">E-MAIL(�系����)</td>
    	<td bgcolor="#FFFFFF">
    		<input name="txEmail1" id="[on,off,off,off][�系����]" type="text" size="30" class="text" value="<%= onepartner.Femail %>">
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">MSN�޽���</td>
    	<td bgcolor="#FFFFFF">
    		<input name="txEmail2" id="[on,off,off,off][MSN�޽���]" type="text" size="30" class="text" value="<%= onepartner.Fmsn %>">
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">��ȭ��ȣ(����)</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="txPhone1" id="[on,on,2,4][��ȭ��ȣ1]" size="5" class="text" value="<% = txPhone1 %>">-
			<input type="text" name="txPhone2" id="[on,on,2,4][��ȭ��ȣ2]" size="5" class="text" value="<% = txPhone2 %>">-
			<input type="text" name="txPhone3" id="[on,on,2,4][��ȭ��ȣ3]" size="5" class="text" value="<% = txPhone3 %>">&nbsp;&nbsp;����:
			<input type="text" name="txPhone4" id="[on,on,2,4][����]" size="5" class="text" value="<%= txPhone4 %>">
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">070 �����ȣ</td><!-- �ű� -->
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="" id="" size="5" class="text" value="070">-
			<input type="text" name="" id="" size="5" class="text" value="0000">-
			<input type="text" name="" id="" size="5" class="text" value="0000">
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">�ڵ�����ȣ</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="txCell1" id="[on,on,2,4][�ڵ�����ȣ1]" size="5" class="text" value="<% = txCell1 %>">-
			<input type="text" name="txCell2" id="[on,on,2,4][�ڵ�����ȣ2]" size="5" class="text" value="<% = txCell2 %>">-
			<input type="text" name="txCell3" id="[on,on,2,4][�ڵ�����ȣ3]" size="5" class="text" value="<% = txCell3 %>">
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">�����ȣ</td>
    	<td bgcolor="#FFFFFF">
    		<input name="txZip1"  id="[on,on,3,3][�����ȣ1]"  size="5" style="background-color:#EEEEEE;" class="text" readonly  value="<%= txZip1 %>">-
			<input type="text" name="txZip2"  id="[on,on,3,3][�����ȣ2]" size="5" style="background-color:#EEEEEE;" class="text" readonly  value="<% = txZip2 %>">
			<input type="button" class="button_s" value="�ּ��Է�" onClick="javascript:TnFindZip('myinfoForm');">
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">�ּ�</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="txAddr1" id="[on,off,1,64][�ּ�1]"  size="50" style="background-color:#EEEEEE;" class="text" readonly  value="<%= onepartner.Faddress %>">
			<br> <input type="text" name="txAddr2" id="[on,off,1,64][�ּ�2]" size="50" maxlength="128" class="text"  value="<%= onepartner.Fmanager_address %>">
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">�μ�-��Ʈ</td>
    	<td bgcolor="#FFFFFF">
    		<%= onepartner.fpart_name %>
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">���α���</td>
    	<td bgcolor="#FFFFFF">
    		������ / ������ / ��Ʈ���ӱ��� / ��Ʈ���������� ......
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">����</td>
    	<td bgcolor="#FFFFFF">
    		<%= onepartner.fposit_name %>
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">��å</td><!-- �ű� -->
    	<td bgcolor="#FFFFFF">
    		���� / ��Ʈ�� / ��Ʈ������
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">������(ī�װ�)</td><!-- �ű� -->
    	<td bgcolor="#FFFFFF">
    		MD�� ��� ��� ī�װ� ���õ��(�ƴϸ� ���� ����)
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td colspan="2" bgcolor="#FFFFFF">
    		<input type="button" class="button_s" value="�⺻���� ����" onclick="javascript:TnJoin10x10(myinfoForm)">
    	</td>
    </tr>
		<!--	<a href="javascript:TnFindZip('myinfoForm')" ><img src="/images/page_2_3.gif" width="60" height="20" border="0" align="absmiddle"></a>	-->

	</form>
<!--�⺻�������� ��-->

<!--��й�ȣ���� ����-->
	<form name="myinfopassword" method="post" action="domyinfo_password.asp">
	<tr>
		<td valign="bottom" colspan=2 bgcolor="FFFFFF">
			<font color="red"><strong>��й�ȣ ����</strong></font> *��й�ȣ �����ÿ��� �Է��ϼ���.
		</td>
	</tr>
	<tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">������й�ȣ</td>
    	<td bgcolor="#FFFFFF">
    		<input  type="password" name="txpass1" id="[off,off,off,off][������й�ȣ]"  size="16" class="input_01">
    		*���� ��й�ȣ�� �Է��ϼ���.
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">�űԺ�й�ȣ</td>
    	<td bgcolor="#FFFFFF">
    		<input  type="password" name="txpass2" id="[off,off,off,off][��й�ȣ����]"  size="16" class="input_01">
			<input  type="password" name="txpass3" id="[off,off,off,off][��й�ȣ����]"  size="16" class="input_01">
    		*����Ͻ� ��й�ȣ�� �ι� �Է��� �ּ���.
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td colspan="2" bgcolor="#FFFFFF">
			<input type="button" class="button" value="��й�ȣ ����" onclick="javascript:passwordedit(myinfopassword)">
    	</td>
    </tr>
	</form>
<!--��й�ȣ���� ��-->

</table>

<table width="50%" align="right" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<font color="red"><strong>�߰�����</strong></font>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="100">�Ի���</td>
    	<td width="100">�ټӿ���</td>
    	<td width="100"></td>
      	<td width="100"></td>
      	<td width="100"></td>
      	<td></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>2001-08-23</td>
    	<td>9</td>
      	<td></td>
      	<td></td>
      	<td></td>
      	<td></td>
    </tr>

	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<font color="red"><strong>����(�ް�)����</strong></font>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td>�����̿�</td>
    	<td>�ݳ�</td>
    	<td>2010�� �ѿ���</td>
      	<td>����ϼ�</td>
      	<td>�ܿ��ϼ�</td>
      	<td>���</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
    	<td>5</td>
    	<td>12</td>
      	<td>17</td>
      	<td>4</td>
      	<td>13</td>
      	<td><input type="button" class="button" value="�ް���û �� ��������" onclick=""></td>
    </tr>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			*�����̿������� 3�������� ��ȿ�ϸ�, �ް���û�� �����̿��������� �����˴ϴ�.
		</td>
	</tr>
</table>



<%
set onepartner = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->