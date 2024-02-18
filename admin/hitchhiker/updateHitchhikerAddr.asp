<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	PageName 	: /admin/hitchhiker/downHitchhiker.asp
'	Description : ��ġ����Ŀ
'	History		: 2006.11.30 ������ ����
'				  2016.07.07 �ѿ�� ���� SSL ����
'#############################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/hitchhiker/hitchhikerCls.asp"-->
<%
Dim idx, iHVol, clsUInfo, arrAddr, sZip, sAdd1, sAdd2, sP, sC
	iHVol = request("iHV")
	idx = requestCheckVar(request("idx"),10)

Set clsUInfo = new CUserInfo
	clsUInfo.frectidx = idx
	clsUInfo.FHVol = iHVol
	arrAddr = clsUInfo.updateHitchAddr()

	if isarray(arrAddr) then
		sZip =	arrAddr(4,0)
		sAdd1 = arrAddr(5,0)
		sAdd2 = arrAddr(6,0)
		sP = split(arrAddr(7,0),"-")
		sC = split(arrAddr(8,0),"-")
	end if
Set clsUInfo = nothing
%>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language="JavaScript" src="/js/common.js"></script>
<script type="text/javascript">

function TnFindZip(frmname){
	window.open('<%= getSCMSSLURL %>/lib/newSearchzip.asp?target=' + frmname, 'findzipcdode', 'width=460,height=250,left=400,top=200,location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no');
}

function jsSubmit(frm){
	if(!frm.zipcode.value || !frm.addr1.value || !frm.addr2.value){
		alert("�ּҸ� �Է����ּ���");
		frm.zipcode.focus();
		return false;
	}
	
	if(!frm.userphone1.value || !frm.userphone2.value || !frm.userphone3.value){
		alert("��ȭ��ȣ�� �Է����ּ���");
		frm.userphone1.focus();
		return false;
	}
	
	if(!frm.usercell1.value || !frm.usercell2.value || !frm.usercell3.value){
		alert("�޴��ȣ�� �Է����ּ���");
		frm.usercell1.focus();
		return false;
	}
	
	frm.action="<%= getSCMSSLURL %>/admin/hitchhiker/updateHitchhikerAddrprocess.asp";
}

</script>

<div style="padding:10 10 0 10">
	<img src="/images/icon_star.gif" align="absmiddle"> <font color="red"><strong> ��ġ����Ŀ Vol.<%=arrAddr(0,0)%> �߼۽�û </strong></font><br>
	<hr>
</div>

<table width="98%" border="0" align="center" class="a" cellpadding="1" cellspacing="5" bgcolor="#F4F4F4">
<form name="frmuser" method="post" onSubmit="return jsSubmit(this);">
<input type="hidden" name="iHV" value="<%=iHVol%>">
<tr>
	<td><font color="999999">+</font> ���̵�</td>
	<td><input type="text" name="sUID" value="<%=arrAddr(3,0)%>" readonly style="background-color:#EEEEEE;"></td>
</tr>
<tr>
	<td><font color="999999">+</font>������</td>
	<td><input type="text" name="recevieName" value="<%=arrAddr(10,0)%>"></td>
</tr>
<tr>
	<td colspan="2">		
		<table border="0" class="a" width="100%" cellpadding=1 cellspacing=0>				
			<tr>
	          <td height="20"><font color="999999">+</font> �ּ�</td>
	          <td height="20">
				<font color="#666666">
					<input type="text" name="zipcode" size="7" value="<%=sZip%>" readOnly style="background-color:#EEEEEE;">
				</font>
				<input type="button" class="button" value="�˻�" onClick="FnFindZipNew('frmuser','E')">
				<input type="button" class="button" value="�˻�(��)" onClick="TnFindZipNew('frmuser','E')">
				<% '<input type="button" class="button" value="�˻�(��)" onClick="TnFindZip('frmuser');"> %>
	          </td>
	        <tr>
	          <td class="padding" height="20">&nbsp;</td>
	          <td height="20">
	            <input type="text" name="addr1" size="20" value="<%=sAdd1%>" readOnly style="background-color:#EEEEEE;">
	          </td>
	        </tr>
	        <tr>
	          <td class="padding" height="20">&nbsp;</td>
	          <td height="20">
	            <input type="text" name="addr2" size="40" value="<%=sAdd2%>" style="ime-mode:active">
	           </td>
	        </tr>
	        <tr>
	          <td class="padding" height="20"><font color="999999">+</font>
	            ��ȭ��ȣ<br>
	          </td>
	          <td height="20"><font color="#666666">
	            <input type="text" name="userphone1" size="3" value="<%=sP(0)%>" maxlength="3">
	            -
	            <input type="text" name="userphone2" size="4" value="<%=sP(1)%>" maxlength="4">
	            -
	            <input type="text" name="userphone3" size="4" value="<%=sP(2)%>" maxlength="4">
	            </font></td>
	        </tr>
	        <tr>
	          <td class="padding" height="9"><font color="999999">+</font>
	            �޴���ȭ</td>
	          <td height="9"><font color="#666666">
	            <input type="text" name="usercell1" size="3" value="<%=sC(0)%>" maxlength="3">
	            -
	            <input type="text" name="usercell2" size="4" value="<%=sC(1)%>" maxlength="4">
	            -
	            <input type="text" name="usercell3" size="4" value="<%=sC(2)%>" maxlength="4">
	            </font></td>
	        </tr>
		</table>
	</td>
		
</tr>
<tr>
	<td colspan="2" align="center"><hr>
		<input type="image" src="/images/icon_confirm.gif">
		<a href="javascript:self.close();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>	
</form>
</table>

<!-- #include virtual="/lib/db/dbclose.asp" -->