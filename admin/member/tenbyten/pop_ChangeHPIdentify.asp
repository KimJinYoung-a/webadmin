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
dim sempno, susername, sjuminno, susercell, hp1, hp2, hp3

'//��ŷ������ ���� ����Ű ����
dim C_dumiKey 
C_dumiKey = session.sessionid 

userid = session("ssBctId")
if userid="" then userid=requestCheckVar(Request("uid"),32)

'// ���� �⺻���� ����
Set cMember = new CTenByTenMember
	cMember.Fuserid = userid
	cMember.fnGetScmMyInfo

	sempno   		= cMember.Fempno
	susername      	= cMember.Fusername
	sjuminno		= cMember.FJuminno
	susercell      	= cMember.Fusercell

Set cMember = Nothing	

if sempno="" or isNull(sempno) then
	Call Alert_close("���� ������ �����ϴ�.\n�����ڿ��� ���ǿ��")
	response.End
end if

if trim(sjuminno)="" or isNull(sjuminno) then
	Call Alert_close("��ϵ� �ֹε�Ϲ�ȣ�� �����ϴ�.\n�濵������ �λ�����ڿ��� �����ּ���.")
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

	frm.mode.value ="ActH";
	frm.target ="hidFrm";
	frm.submit();
}

// ������ȣ ����
function actChgHP(){
    var frm = document.frmChkId;
	frm.mode.value ="chgHP";
	frm.target ="hidFrm";
	frm.submit();
}
</script>
<form name="frmChkId" method="post" action="doChangeHPIdentify.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="ActH">
<input type="hidden" name="userid" value="<%=userid%>">
<input type="hidden" name="dumiKey" value="<%= C_dumiKey %>">
<table width="100%" cellpadding="2" cellspacing="1" border="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td colspan="2" bgcolor="#E8F0FF"><b>�޴�����ȣ ���� / ���� Ȯ��</b></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�����ȣ</td>
	<td bgcolor="white"><%=sempno%></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�̸�</td>
	<td bgcolor="white"><input type="text" name="username" value="<%=susername%>" readonly class="text_ro" size="10"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�ֹε�Ϲ�ȣ</td>
	<td bgcolor="white">
		<input type="text" name="jumin1" value="<%=left(sjuminno,6)%>" readonly class="text_ro" size="6">
		<input type="password" name="jumin2" value="*******" readonly class="text_ro" size="7">
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">�޴�����ȣ</td>
	<td bgcolor="white">
		<select name="hpNum0" class="select">
			<option value="SKT">SKT</option>
			<option value="KTF">KTF</option>
			<option value="LGT">LGT</option>
		</select>
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
	</td>
</tr>
<tr>
	<td colspan="2" bgcolor="#FEF8F8">
		�� �Է��� �޴������� �������ڰ� �߼۵Ǹ�, ����Ȯ���� �Ϸ�Ǹ� [�� ����]�� [�޴�����ȣ]�� �����˴ϴ�.<br>
		&nbsp;&nbsp;(����Ȯ���� �޴�����ȣ ���� �� ���� 1ȸ�� Ȯ��)<br>
		�� �̸��� �ֹε�Ϲ�ȣ�� �߸��Ǿ����� ��� <b>�濵������</b> �λ����ڿ��� �������ּ���.
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