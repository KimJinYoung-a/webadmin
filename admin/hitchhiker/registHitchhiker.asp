<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	PageName 	: /admin/hitchhiker/downHitchhiker.asp
'	Description : ��ġ����Ŀ ��ûȸ������Ʈ �߼�Ȯ��, �߼۽�û, ��߼۽�û
'	History		: 2006.11.30 ������ ����
'				  2016.07.07 �ѿ�� ���� SSL ����
'#############################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/hitchhiker/hitchhikerCls.asp"-->
<%
Dim idx, pMode, iHVol, clsUInfo, sZip, sAdd1, sAdd2, sP, sC, sChk,sUsername, iAV, sUserID
	idx = requestCheckVar(request("idx"),10)
	pMode	= Request("pMode")
	iHVol = Request("iHV")
	iAV = Request("iAV")
	sUserID = requestCheckVar(request("sUID"),32)

IF idx <> "" or sUserID <> "" THEN
	Set clsUInfo = new CUserInfo
		clsUInfo.frectidx = idx
		clsUInfo.FiHVol = iHVol
		clsUInfo.FUID = sUserID	'/�˾����� �Ķ��Ÿ�� �Ѿ���°����� ���� ���������ο��� ���̵�˻����� ���� ���� �޾ƾ��� 2017.07.17 �ѿ��
		clsUInfo.fnGetUserInfo()
		sChk = clsUInfo.FChk
		sUsername =	clsUInfo.FUsername
		sAdd1 = clsUInfo.FAddress1
		sAdd2 = clsUInfo.FAddress2
		sZip =	clsUInfo.FZipCode
		sUserID = clsUInfo.FUID

		if Not(trim(clsUInfo.FPhone)="" or isnull(clsUInfo.FPhone)) then
			sP =	split(clsUInfo.FPhone,"-")
		else
			sP =	split("--","-")
		end if

		if Not(trim(clsUInfo.FCell)="" or isnull(clsUInfo.FCell)) then
			sC =	split(clsUInfo.FCell,"-")
		else
			sC =	split("--","-")
		end if

		If CInt(clsUInfo.FRegCount) > 0 and pMode = "A" Then
			response.write "<script>"
			response.write "	alert('�̹� Vol."&iHVol&"�� "&clsUInfo.FiApplyVol&"ȸ����  ���̵� ��ϵǾ��ֽ��ϴ�.\n\n��߼� ó�����ּ���');"
			response.write "	window.close();"
			response.write "</script>"
			response.end
		End If

	Set clsUInfo = nothing
END IF	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language="JavaScript" src="/js/common.js"></script>
<script type="text/javascript">

	function jsChkUser(){
		var frm = document.frmuser;
		if (!frm.sUID.value) {
			alert("���̵� �Է����ּ���");
			frm.sUID.focus();
			return;
		}
		
		frm.action = "registHitchhiker.asp";
		frm.submit();
	}
	
	function TnFindZip(frmname){
		window.open('<%= getSCMSSLURL %>/lib/newSearchzip.asp?target=' + frmname, 'findzipcdode', 'width=460,height=250,left=400,top=200,location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no');
	}

	function jsSubmit(frm){
		if(!frm.sUID.value){
			alert("���̵� �Է����ּ���");
			frm.sUID.focus();
			return false;
		}
		
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
			alert("��ȭ��ȣ�� �Է����ּ���");
			frm.userphone1.focus();
			return false;
		}
		
		frm.action="<%= getSCMSSLURL %>/admin/hitchhiker/processHitchhiker.asp";
	}

</script>
</head>
<body leftmargin=0 topmargin=0>
<div style="padding:10 10 0 10">
<img src="/images/icon_star.gif" align="absmiddle"> <font color="red"><strong> ��ġ����Ŀ Vol.<%=iHVol%> �߼۽�û </strong></font><br>
<hr>
 </div>
<table width="98%" border="0" align="center" class="a" cellpadding="1" cellspacing="5" bgcolor="#F4F4F4">
<form name="frmuser" method="post" onSubmit="return jsSubmit(this);">
<input type="hidden" name="pMode" value="<%=pMode%>">
<input type="hidden" name="iHV" value="<%=iHVol%>">
<input type="hidden" name="iAV" value="<%=iAV%>">
	<tr>
		<td><font color="999999">+</font> ���̵�</td>
		<td><input type="text" name="sUID" value="<%=sUserID%>"> <input type="button" value="ȸ������ ��������" class="a" onClick="jsChkUser();"> </td>
	</tr>
	<%IF sUserID <> ""  THEN%>
	<tr>
		<td><font color="999999">+</font> �̸�</td>
		<td><input type="text" name="receviename" value="<%=sUsername%>"></td>
	</tr>
	<% End If %>
	<tr>
		<td colspan="2">		
		<div id="uInfo" <%IF sUserID = ""  THEN%>style="display:none;"<%END IF%>>
		<%IF sChk = 1 THEN%>
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
			<%ELSE%>
				<font color="red"><center>��ϵ��� ���� ȸ���Դϴ�.</center></font>
			<%END IF%>
			</div>
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
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->