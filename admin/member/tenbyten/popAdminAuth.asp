<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ������ѵ��
' History : ������ ����
'			2018.04.17 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%
IF application("Svr_Info")<>"Dev" THEN
	if Not(C_privacyadminuser) or Not(isVPNConnect) then
			response.write "���ε� �������� �ƴմϴ�. ������ ���ǿ�� [���ٱ���:" & C_privacyadminuser & "/VPN:" & isVPNConnect & "]"
			response.end
	end if
end if

Dim sEmpNo,sUsername
Dim suserid,sfrontid, ipart_sn,iposit_sn,ijob_sn ,ilevel_sn,iuserdiv, lv1customerYN, lv2partnerYN, lv3InternalYN, icriticinfouser
Dim cMember, i, mydpID, isdispmember

sEmpNo = requestCheckVar(Request("sEPN"),14)
IF 	sEmpNo = "" THEN
    Alert_return("�߸��� ���԰���Դϴ�.")
    response.end ''�߰�/2014/07/14
END IF

IF application("Svr_Info")="Dev" THEN
	isdispmember = true
else
	' ISMS �ɻ�� ���� �������� ���ٱ��� ����/����/���� Ư������� ���̰�(�ѿ��,������,�̹���)	' 2020.10.12 �ѿ��
	if C_privacyadminuser or C_PSMngPart then
		isdispmember = true
	else
		isdispmember = false
	end if
end if

set cMember = new CTenByTenMember
	cMember.Fempno = sEmpNo
	cMember.fnGetMemberData
	sempno   		= cMember.Fempno
	suserid     	= cMember.Fuserid
	sfrontid    	= cMember.Ffrontid
	sUsername 		= cMember.Fusername
	ipart_sn        = cMember.Fpart_sn
	iposit_sn       = cMember.Fposit_sn
	ijob_sn         = cMember.Fjob_sn
	ilevel_sn       = cMember.Flevel_sn
	iuserdiv        = cMember.Fuserdiv
	icriticinfouser = cMember.Fcriticinfouser
	lv1customerYN = cMember.Flv1customerYN
	lv2partnerYN = cMember.Flv2partnerYN
	lv3InternalYN = cMember.Flv3InternalYN
	mydpID = myDepartmentId(session("ssBctID"))
set 	cMember = nothing

Dim oAddLevel
set oAddLevel = new CPartnerAddLevel
oAddLevel.FRectUserid=suserid
oAddLevel.FRectOnlyAdd = "on"

if (oAddLevel.FRectUserID<>"") then
    oAddLevel.getUserAddLevelList
end if

dim olog
Set olog = new CTenByTenMember
	olog.FPagesize = 50
	olog.FCurrPage = 1
	olog.frectempno=sEmpNo
	olog.getUserTenbytenAdminAuthLog()
%>

<html>
<head>
<title>�������� ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language="javascript" src="/js/common.js"></script>
<script type="text/javascript">

function chRetireUser(){
	if(!document.frmRetireUser.sUI.value) {
		alert("���� ���̵� �����ϴ�");
		return;
	}
	if(!document.frmRetireUser.sFUI.value) {
		alert("�ٹ����ٻ���Ʈ ���̵� �����ϴ�");
		return;
	}
	if(!document.frmRetireUser.sUN.value) {
		alert("�̸��� �Էµ��� �ʾҽ��ϴ�.");
		return;
	}
	if (frm_member.selPN.value!='35' || frm_member.selLN.value!='' || frm_member.lv1customerYN.checked!=false || frm_member.lv2partnerYN.checked!=false || frm_member.lv3InternalYN.checked!=false){
		alert("���� ������ ����߿� �ֽ��ϴ�.\n������ ���ܽ�Ű�ð� �̿��� �ּ���.");
		return;
	}
	if (confirm('�ش� ����Ʈ ���̵� ���ó�� �Ͻðڽ��ϱ�?\n->�������� ȸ��\n->ȸ����� ����')==true){
		frmRetireUser.submit();
	}
}

function jsChkSubmit(){
	<% if not(C_ADMIN_AUTH or C_PSMngPartPower) then %>
		if(!document.frm_member.sUI.value) {
			alert("WEBADMIN ���̵� �Է����ֽʽÿ�.");
			document.frm_member.sUI.focus();
				return ;
		}
	<% else %>
		if(document.frm_member.sUI.value == "") {
			if (confirm("[�����ڱ���] WEBADMIN ���̵� �����ϴ�.\n\n��� �����Ͻðڽ��ϱ�?") != true) {
				alert("WEBADMIN ���̵� �Է����ֽʽÿ�.");
				document.frm_member.sUI.focus();
				return ;
			}
		}
	<% end if %>

	// WEBADMIN ���̵� �ԷµǾ� ������쿡�� üũ��
	if(frm_member.sUI.value != "") {
		if(typeof(document.frm_member.sP) != "undefined"){
			if(!document.frm_member.sP.value) {
				alert("��й�ȣ�� �Է����ֽʽÿ�.");
				document.frm_member.sP.focus();
				return ;
			}

			if (document.frm_member.sP.value.replace(/\s/g, "").length < 6 || document.frm_member.sP.value.replace(/\s/g, "").length > 16){
				alert("��й�ȣ�� ������� 6~16���Դϴ�.");
				document.frm_member.sP.focus();
				return ;
			}

			if ((document.frm_member.sP.value)!=(document.frm_member.sP1.value)){
				alert("��й�ȣ�� ��ġ���� �ʽ��ϴ�.");
				document.frm_member.sP1.focus();
				return;
			}

			if (!fnChkComplexPassword(frm_member.sP.value)) {
				alert('���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)');
				frm_member.sP.focus();
				return;
			}
		}
	}

	if(document.frm_member.hidID.value =="0"){
		alert("���̵� �ߺ�üũ�� ���ּ���");
		return ;
	}

	if(document.frm_member.selPN.value == "") {
		alert("���α���(�μ�)�� �Է����ֽʽÿ�.");
		return ;
	}

	if(document.frm_member.selLN.value == "") {
		if (confirm("!!! ���� ���ѵ�� ���� !!!\n\n�����Ͻðڽ��ϱ�?") == true) {
			//
		} else {
			alert("���α���(���)�� �Է����ֽʽÿ�.");
			return;
		}
	}

	if ((document.frm_member.selUD.value=="")||(document.frm_member.selUD.value*1>111)){
		alert("�������Ѽ������� ");
		return ;
	}

	document.frm_member.submit();
}
			//���̵� �ߺ�üũ
	function jsChkID(){
		var winID;
		var frm = document.frmID;
		if(!document.frm_member.sUI.value){
			alert("���̵� �Է����ּ���");
			document.frm_member.sUI.focus();
			return;
		}
		frmID.sUI.value = document.frm_member.sUI.value;
		frmID.sUN.value = document.frm_member.sUN.value;
		winID = window.open("","popid","width=0, height=0");
		document.frmID.target = "popid";
		document.frmID.submit();
	}
	function jsChkfrontname(){
		var winID;
		var frm = document.frmchk;
		if(!document.frm_member.sUI.value){
			alert("���̵� �Է����ּ���");
			document.frm_member.sUI.focus();
			return;
		}
		frmchk.sUI.value = document.frm_member.sUI.value;
		frmchk.sUN.value = document.frm_member.sUN.value;
		frmchk.sFUI.value = document.frm_member.sFUI.value;
		frmchk.sEN.value = "<%= sEmpNo %>";
		winID = window.open("","popid","width=0, height=0");
		document.frmchk.target = "popid";
		frmchk.mode.value="frontnamewebadmincheck";
		document.frmchk.submit();
		frmchk.mode.value="";
	}
	function jschangefrontname(){
		var winID;
		var frm = document.frmchk;
		if(!document.frm_member.sUI.value){
			alert("���̵� �Է����ּ���");
			document.frm_member.sUI.focus();
			return;
		}
		frmchk.sUI.value = document.frm_member.sUI.value;
		frmchk.sUN.value = document.frm_member.sUN.value;
		frmchk.sFUI.value = document.frm_member.sFUI.value;
		frmchk.sEN.value = "<%= sEmpNo %>";

		var ret = confirm('�������� Ȯ���� ��쿡�� �����ϼž� �մϴ�.\n������������[����->����Ʈ����Ʈ] �Ͻðڽ��ϱ�?');
		if (ret){
			winID = window.open("","popid","width=0, height=0");
			document.frmchk.target = "popid";
			frmchk.mode.value="frontnamewebadminchange";
			document.frmchk.submit();
			frmchk.mode.value="";
		}
	}

	// ���� ���� �˾�
	function popAuthSelect()
	{
		<% if application("Svr_Info")<>"Dev" THEN %>
			<% if Not(C_ADMIN_AUTH or C_PSMngPartPower) then %>
				alert('������ �����ϴ�. ������ ���� ���');
				return;
			<% end if %>
		<% end if %>

		var popwin = window.open("/admin/menu/pop_Menu_auth.asp?userid=<%= suserid %>", "popMenuAuth","width=360,height=200,scrollbars=no");
		popwin.focus();
	}

	// �˾����� ���ñ��� �߰�
	function addAuthItem(psn,pnm,lsn,lnm)
	{
		var lenRow = tbl_auth.rows.length;

		// ������ ���� �ߺ� ��Ʈ ���� �˻�
		if(lenRow>1)	{
			for(l=0;l<document.all.part_sn.length;l++)	{
				if(document.all.part_sn[l].value==psn) {
					alert("�̹� ������ ������ �μ��Դϴ�.\n���� �μ��� �����ϰ� ������ �ٽ� �������ּ���.");
					return;
				}
			}
		}
		else {
			if(lenRow>0) {
				if(document.all.part_sn.value==psn) {
					alert("�̹� ������ ������ �μ��Դϴ�.\n���� �μ��� �����ϰ� ������ �ٽ� �������ּ���.");
					return;
				}
			}
		}

		// ���߰�
		var oRow = tbl_auth.insertRow(lenRow);
		oRow.onmouseover=function(){tbl_auth.clickedRowIndex=this.rowIndex};

		// ���߰� (�μ�,���,������ư)
		var oCell1 = oRow.insertCell(0);
		var oCell2 = oRow.insertCell(1);
		var oCell3 = oRow.insertCell(2);

		oCell1.innerHTML = pnm + "<input type='hidden' name='part_sn' value='" + psn + "'>";
		oCell2.innerHTML = lnm + "<input type='hidden' name='level_sn' value='" + lsn + "'>";
		oCell3.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delAuthItem()' align=absmiddle>";
	}

	// ���ñ��� ����
	function delAuthItem()
	{
		<% if application("Svr_Info")<>"Dev" THEN %>
			<% if not(C_ADMIN_AUTH or C_PSMngPartPower) then %>
				alert('������ �����ϴ�. ������ ���� ���');
				return;
			<% end if %>
		<% end if %>
		if(confirm("������ ������ �����Ͻðڽ��ϱ�?"))
			tbl_auth.deleteRow(tbl_auth.clickedRowIndex);
	}
	//��й�ȣ����
	function jsChangePW(uid){
	    var popwinPass = window.open("/admin/member/tenbyten/pop_ChangPassword.asp?userid="+uid,"popwinPass","width=1024,height=400,scrollbars=yes,resizable=yes");
		popwinPass.focus();
		
	}
</script>
</head>
<body leftmargin="10" topmargin="10">
<form name="frm_member" method="post" action="/admin/member/tenbyten/procAdminAuth.asp" style="margin:0px;">
<input type="hidden" name="hidID" value="1">
<input type="hidden" name="sEN" value="<%=sEmpNo%>">
<input type="hidden" name="selPoN" value="<%=iposit_sn%>">
<input type="hidden" name="selJN" value="<%=ijob_sn%>">
<input type="hidden" name="sUN" value="<%=sUsername%>">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><strong>��� �������� ���</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="3" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>" width="130">���</td>
			<td bgcolor="#FFFFFF">
				<%= sempno %>
			</td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>" width="130">WEBADMIN ���̵�</td>
			<td bgcolor="#FFFFFF">
					<input type="text" name="sUI" class="text" size="20" value="<%=suserid%>" onClick="document.frm_member.hidID.value=0;" onKeypress="document.frm_member.hidID.value=0;"> <input type="button" name="btnChkID" value="���̵� �ߺ�üũ" onClick="jsChkID();" class="input">
			</td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">�ٹ����ٻ���Ʈ ���̵�</td>
			<td bgcolor="#FFFFFF">
				<input type="text" name="sFUI" class="text" size="20" value="<%=sfrontid%>">
				<input type="button" name="btnChkID" value="��������üũ[����,����Ʈ����Ʈ]" onClick="jsChkfrontname();" class="input">

				<% if isdispmember then %>
					<% if C_ADMIN_AUTH or C_PSMngPart then %>
						<input type="button" name="btnChkID" value="������������[����->����Ʈ����Ʈ]" onClick="jschangefrontname();" class="input">
					<% end if %>
				<% end if %>
			</td>
		</tr>
		<% IF isNull(suserid) or suserid = "" THEN %>
			<tr align="left" height="25">
				<td bgcolor="<%= adminColor("tabletop") %>">��й�ȣ</td>
				<td bgcolor="#FFFFFF">
					<input type="password" name="sP" class="text" size="20" maxlength="60" value="">
				</td>
			</tr>
			<tr align="left" height="25">
				<td bgcolor="<%= adminColor("tabletop") %>">��й�ȣȮ��</td>
				<td bgcolor="#FFFFFF">
					<input type="password" name="sP1" class="text" size="20" maxlength="60" value="">
				</td>
			</tr>
		<% ELSE %>
			<tr align="left" height="25">
				<td bgcolor="<%= adminColor("tabletop") %>">�н�����</td>
				<td bgcolor="#FFFFFF">
					<% if isdispmember then %>
						<% If (C_ADMIN_AUTH) or C_PSMngPart OR (mydpID = "88") Then %>
							<input type="button" class="button" value="����(���̵� �α��ο� �н�����)" onClick="jsChangePW('<%=suserid%>');">
						<% End If %>
					<% End If %>
					<br>�� �н����� ����� �ʱ�ȭ �Ǵ� ���
					<br>1. ������ ������ ���� �ΰ��, ��������� �����
					<br>2. ��Ⱓ �̻������ ���� ������ �����, ����� ������.
					<br>3. �н����带 Ʋ���� �����, ����� ���� �˴ϴ�.
				</td>
			</tr>
		<% END IF %>
	    <tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">���� ����(���)</td>
			<td bgcolor="#FFFFFF">
				<%=printPartOption("selPN", ipart_sn)%>
				&nbsp;&nbsp;
				<%=printLevelOption("selLN", ilevel_sn)%>
			</td>
		</tr>
	    <tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">����������ޱ���</td>
			<td bgcolor="#FFFFFF">
				<% 'Call DrawSelectBoxCriticInfoUser("criticinfouser", icriticinfouser) %>
				<input type="hidden" name="criticinfouser" value="<%= icriticinfouser %>">
				<input type="checkbox" name="lv1customerYN" value="Y" <% if lv1customerYN = "Y" then %>checked<% end if %> >LV1(������)
				<input type="checkbox" name="lv2partnerYN" value="Y" <% if lv2partnerYN = "Y" then %>checked<% end if %> >LV2(��Ʈ������)
				<input type="checkbox" name="lv3InternalYN" value="Y" <% if lv3InternalYN = "Y" then %>checked<% end if %> >LV3(��������)
			</td>
		</tr>
	    <tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">�߰� ����</td>
			<td bgcolor="#FFFFFF">
			    <table border="0" cellspacing="0" class="a">
			    <tr>
			        <td >
        			    <table name='tbl_auth' id='tbl_auth' class=a>
        			    <% if (oAddLevel.FResultCount>0) then %>
        			        <% for i=0 to oAddLevel.FResultCount-1 %>
        			        <tr onMouseOver='tbl_auth.clickedRowIndex=this.rowIndex'>
        						<td><%= oAddLevel.FItemList(i).Fpart_name %><input type='hidden' name='part_sn' value='<%= oAddLevel.FItemList(i).Fpart_sn %>'></td>
        						<td><%= oAddLevel.FItemList(i).Flevel_name %><input type='hidden' name='level_sn' value='<%= oAddLevel.FItemList(i).Flevel_sn %>'></td>
        						<td><img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delAuthItem()' align=absmiddle></td>
        					</tr>
        				    <% next %>
        				<% else %>
        				    <tr onMouseOver='tbl_auth.clickedRowIndex=this.rowIndex'>
						    <td><input type='hidden' name='part_sn' value=''></td>
						    <td><input type='hidden' name='level_sn' value=''></td>
						    <td></td>
					        </tr>
        				<% end if %>
        			    </table>
			        </td>
        			<td valign="bottom"><input type="button" class='button' value="�߰�" onClick="popAuthSelect()"></td>
        		</tr>
        		</table>


			</td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
			<td bgcolor="#FFFFFF">
				<% DrawAuthBoxTenUser "selUD",iUserdiv %>
			</td>
		</tr>

		</table>

	</td>
</tr>
<Tr>
	<td align="center">
		<% if isdispmember then %>
			<input type="button" value="����" class="input" onclick="jsChkSubmit();">
			<% if C_ADMIN_AUTH or C_PSMngPart then %>
				<% if sfrontid<>"" and not(isnull(sfrontid)) and sUsername<>"" and not(isnull(sUsername)) then %>
					&nbsp;&nbsp;
					<input type="button" value="����Ʈ���ó��(����ȸ��,��޺���)" class="input" onclick="chRetireUser();">
				<% end if %>
			<% end if %>
		<% end if %>
	</td>
</tr>
</table>
</form>

<% if olog.FResultCount>0 then %>
	<br>
	<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%=olog.FtotalCount%></b>
			&nbsp;&nbsp;�� �ֱ� 50�Ǹ� ǥ�� �˴ϴ�.
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width=60>�α׹�ȣ</td>
		<td>���泻��</td>
		<td width=100>������</td>
	</tr>
	<% for i=0 to olog.FResultCount - 1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= olog.FitemList(i).flogidx %></td>
		<td align="left"><%= olog.FitemList(i).flogmsg %></td>
		<td>
			<%= olog.FitemList(i).fadminid %>
			<Br><%= left(olog.FitemList(i).fregdate,10) %>
			<Br><%= mid(olog.FitemList(i).fregdate,12,22) %>
		</td>
	</tr>
	<% next %>
	</table>
<% end if %>

<!-- ���̵� �ߺ�üũ-->
<form name="frmID" method="post" action="/admin/member/tenbyten/member_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="R">
<input type="hidden" name="sEN" value="">
<input type="hidden" name="sUI" value="">
<input type="hidden" name="sUN" value="">
</form>
<form name="frmRetireUser" method="post" action="/admin/member/tenbyten/member_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="RetireUser">
<input type="hidden" name="sUI" value="<%=suserid%>">
<input type="hidden" name="sFUI" value="<%=sfrontid%>">
<input type="hidden" name="sUN" value="<%=sUsername%>">
</form>
<form name="frmchk" method="post" action="/admin/member/tenbyten/member_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="sEN" value="">
<input type="hidden" name="sUI" value="">
<input type="hidden" name="sUN" value="">
<input type="hidden" name="sFUI" value="">
</form>
