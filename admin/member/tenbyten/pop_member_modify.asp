<%@ language=vbscript %>
<% option explicit %>
<%
session.codePage = 949
Response.CharSet = "EUC-KR"
%>
<%
'###########################################################
' Description :  ������
' History : 2010.12.15 ������ ����
'           2013.03.26 ������ ����; ����Ʈ�� �߰�
'			2016.07.07 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
Dim cMember, sEmpNo, idepartment_id, irank_sn, mydpID, i, menupos, arrjuminno, sjuminno1, sjuminno2
Dim suserid,sfrontid, susername,  dbirthday , blnissolar ,szipcode ,blnsexflag ,szipaddr,suseraddr
Dim blnstatediv ,djoinday ,dretireday ,suserimage ,ipart_sn ,iposit_sn ,ijob_sn ,ilevel_sn ,iuserdiv
dim suserphone , susercell  ,susermail  ,smsnmail , sinterphoneno, sextension , sdirect070 , sjobdetail ,sjuminno
Dim drealjoinday   ,iretirereason, smessenger, mywork,vUserImage, selRD_y, selRD_m, selRD_d
dim changedate, isIdentify, gsshopuserid
dim arrList, intLoop, ipersonalmail, userNameEN, isdispmember
	menupos = requestCheckVar(Request("menupos"),10)
	IF menupos ="" THEN menupos = 1176
	sEmpNo = requestCheckVar(Request("sEPN"),14)

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

IF 	sEmpNo <> "" THEN
set cMember = new CTenByTenMember
	cMember.Fempno = sEmpNo
	cMember.fnGetMemberData

	sempno   		= cMember.Fempno
	suserid     = cMember.Fuserid
	sfrontid    = cMember.Ffrontid
	susername   = cMember.Fusername
	userNameEN   = cMember.fuserNameEN
	sjuminno		= cMember.FJuminno

	IF sjuminno <> "" THEN
		sjuminno1 = trim(left(sjuminno,6))
		sjuminno2 = trim(mid(sjuminno,8,7))
	END IF

	dbirthday   = cMember.Fbirthday
	blnissolar  = cMember.Fissolar
	blnsexflag	= cMember.Fsexflag
	szipcode    = cMember.Fzipcode
	szipaddr		= cMember.Fzipaddr
	suseraddr   = cMember.Fuseraddr
	mywork			= cMember.Fmywork
	suserphone  = cMember.Fuserphone
	susercell   = cMember.Fusercell
	susermail   = cMember.Fusermail
	smsnmail    = cMember.Fmsnmail
	smessenger	= cMember.Fmessenger
	sinterphoneno 	= cMember.Finterphoneno
	sextension      = cMember.Fextension
	sdirect070      = cMember.Fdirect070
	sjobdetail      = cMember.Fjobdetail
	blnstatediv     = cMember.Fstatediv
	djoinday        = cMember.Fjoinday
	dretireday      = cMember.Fretireday
	if dretireday<>"" and not isnull(dretireday) THEN
		selRD_y = Year(dretireday)
		selRD_m = Month(dretireday)
		selRD_d = Day(dretireday)
	end if
	vUserImage    	= cMember.Fuserimage
	ipart_sn        = cMember.Fpart_sn
	iposit_sn       = cMember.Fposit_sn
	ijob_sn         = cMember.Fjob_sn
	ilevel_sn       = cMember.Flevel_sn
	iuserdiv        = cMember.Fuserdiv
	drealjoinday		= cMember.Frealjoinday
	iretirereason		= cMember.Fretirereason

	idepartment_id		= cMember.Fdepartment_id
	irank_sn		= cMember.Frank_sn
	mydpID = myDepartmentId(session("ssBctID"))
	ipersonalmail		= cMember.Fpersonalmail
	isIdentify		= cMember.FisIdentify
	gsshopuserid = cMember.fgsshopuserid

	'�߷ɷα�
	arrList = cMember.fnGetUserModLog
	'������ �߷��� ��������
	IF isArray(arrList) THEN
		changedate = arrList(4,0)
	else
        changedate = date()
    END IF

set 	cMember = nothing
END IF

If irank_sn = "" OR isNull(irank_sn) Then
	irank_sn = "0"
End IF

Dim oAddLevel
set oAddLevel = new CPartnerAddLevel
	oAddLevel.FRectUserid=suserid
	oAddLevel.FRectOnlyAdd = "on"

	if (oAddLevel.FRectUserID<>"") then
	    oAddLevel.getUserAddLevelList
	end if
%>
<html>
<head>
<title>�������� ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<script type="text/javascript">

function chk_form(form){
	if(typeof(document.all.sfImg)=="undefined"){
		form.sUImg.value = "";
	}else{
	form.sUImg.value = document.all.sfImg.value;
}
	/* ����Է��� ������ ������Ʈ ����
	if(form.sEP.value == "")
	{
		alert("��� �α��ο� ��й�ȣ�� �Է����ּ���.");
		form.sEP.focus();
		return false;
	}
	*/

	if(form.sUN.value == "")
	{
		alert("�̸��� �Է����ֽʽÿ�.");
		form.sUN.focus();
		return false;
	}
	if(form.department_id.value == "")
	{
		alert("�μ��� �Է����ֽʽÿ�.");
		form.department_id.focus();
		return false;
	}

	/* �ӽ÷� üũ ����
	if(form.sJN1.value == "")
	{
		alert("�ֹε�Ϲ�ȣ�� �Է����ּ���.");
		form.sJN1.focus();
		return false;
	}

	if(form.sJN2.value == "")
	{
		alert("�ֹε�Ϲ�ȣ�� �Է����ּ���.");
		form.sJN2.focus();
		return false;
	}
	*/

	if(form.selPN.value == "")
	{
		alert("�μ��� �Է����ֽʽÿ�.");
		return false;
	}

	if(form.selPoN.value == "")
	{
		alert("������ �Է����ֽʽÿ�.");
		return false;
	}



	if (form.tmpstatediv[0].checked == true) {
		// ���
		form.statediv.value = "N";
		if (form.tmpretirereason) {
			for (var i = 0; i < form.tmpretirereason.length; i++) {
				if (form.tmpretirereason[i].checked == true) {
					form.retirereason.value = form.tmpretirereason[i].value;
					break;
				}
			}
		}
	} else if (form.tmpstatediv[1].checked) {
		// ������ ��ȯ
		form.statediv.value = "N";
		if (form.tmpretirereason) {
			form.retirereason.value = "99";
		}
	}

	return true;
}

function jsDelUser(){
	if(confirm("��������� �����Ͻðڽ��ϱ�?")){
		document.frmDel.submit();
	}
}

function jumin_format(obj) {
	var tmp;

	tmp = document.frm_member.sJN1.value + document.frm_member.sJN2.value ;
	if(tmp==""){return;}
	tmp = tmp.replace(/\-/g, "");
/*
	if(!jsChkSocialNum(tmp)){
		alert("��ȿ���� ���� �ֹε�Ϲ�ȣ�Դϴ�. Ȯ�� �� �ٽ� ������ּ���");
		document.frm_member.sJN1.value = "";
		document.frm_member.sJN2.value = "";
		document.frm_member.sJN1.focus();
		return;
	}
*/
	if(tmp.substring(6,7)=="1"|| tmp.substring(6,7)=="2"){
	document.frm_member.selBD_y.value = "19"+tmp.substring(0,2);
	}else{
	document.frm_member.selBD_y.value = "20"+tmp.substring(0,2);
	}

	document.frm_member.selBD_m.value = parseInt(tmp.substring(2,4),10);
	document.frm_member.selBD_d.value = parseInt(tmp.substring(4,6),10);

	if(tmp.substring(6,7)=="1"|| tmp.substring(6,7)=="3"){
	document.frm_member.rdoSf[0].checked = true;
	document.frm_member.rdoSf[1].checked = false;
	}else{
	document.frm_member.rdoSf[0].checked = false;
	document.frm_member.rdoSf[1].checked = true;
	}
}

function jsChkSocialNum(sno){
	  var IDAdd = "234567892345";
	  var iDot=0;

	  //����Ȯ��
	  if(!IsDigit(sno)){
		return false;
	}
	  //���ڰ� 13�ڸ� ���� Ȯ��
	  if(sno.length != 13){
		return false;
	 }
	  if (sno.substring(2,3) > 1) return false;
	  if (sno.substring(4,5) > 3) return false;
	  if (sno.substring(0,2) == '00' && (sno.substring(6,7) != 0 || sno.substring(6,7) != 9 || sno.substring(6,7) != 3 || sno.substring(6,7) !=4)) return false;
	  if (sno.substring(0,2) != '00' && (sno.substring(6,7) > 4 || sno.substring(6,7) == 0)) return false;

	  for(var i=0; i < 13; i ++)
		iDot = iDot + sno.substr(i, 1) * IDAdd.substr(i,1);

	  iDot = 11 - (iDot % 11);

	  if(iDot == 10){
		iDot = 0;
	  } else if (iDot == 11){
		iDot = 1;
	  }

	  if(sno.substr(12,1) == iDot){
		return true;
	  } else {
		return false;
	  }
}

function updateIsusingZero(){
	if (confirm('�ش� ����� �����Ͻðڽ��ϱ�?')){
		document.frmDel2.submit();
    }
}
function jsFileDel(){
	$("#dFile").html('');
}


function jsRegPhoto(){
	var winP= window.open('popAddphoto.asp','imageupload','width=380,height=150');
	winP.focus();
}

//�߷����ó��
function jsDelModLog(logidx){
	if(confirm("�߷��� ����Ͻðڽ��ϱ�?")){
	document.frmlogDel.logidx.value = logidx;
	document.frmlogDel.submit();
	}
	}

function jsChangeretire(){
	var empno="<%=sEmpNo%>";
	if(empno==""){
		alert("����� �Է����ּ���");
		return;
	}
	frmchk.sEN.value = empno;
	winID = window.open("","popid","width=0, height=0");
	document.frmchk.target = "popid";
	frmchk.mode.value="changenotretire";
	document.frmchk.submit();
	frmchk.mode.value="";
}

</script>
</head>
<body leftmargin="5" topmargin="5">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><strong>��� ���� ���</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td><font color="red">[�⺻����]</font><br>
		<table width="100%" border="0" cellpadding="3" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<form name="frm_member" method="POST" action="/admin/member/tenbyten/member_process.asp" onsubmit="return chk_form(this)" style="margin:0px;" >
		<input type="hidden" name="mode" value="U">
		<input type="hidden" name="sUImg" value="<%=vUserImage%>">
		<input type="hidden" name="sUI" value="<%=suserid%>">
		<input type="hidden" name="selPN" value="<%= ipart_sn %>">
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>" width="130">���</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sEN" class="text" size="20" maxlength="60" value="<%=sEmpNo%>" readonly style="border:0;"></td>
			<td rowspan="5" bgcolor="#FFFFFF" width="130"  align="center">
				<table border="0" cellpadding="0" cellspacing="0" height="132" class="a">
				<tr >
					<td >
						<div id="dFile">
						<img src="<%=vUserImage%>" width="130" alt="�����̹�������" style="cursor:pointer" onClick="window.open('http://www.10x10.co.kr/common/showimage.asp?img=<%=vUserImage%>', 'imageView', 'width=10,height=10,status=no,resizable=yes,scrollbars=yes');">
						<%if vUserImage <> "" then%>
						<div style="text-align:right;">
						<a href="javascript:jsFileDel('<%=vUserImage%>')" style="font-size:10px;color:blue;">[x]</a>
						</div>
						<%end if%>
						<input type="hidden" name="sfImg" value="<%=vUserImage%>">
						</div>
					</td>
				</tr>
				<tr>
					<td align="center" bgcolor="FFFFFF" valign="bottom"><input type="button" class="button" value="�������" onClick="jsRegPhoto();"></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>" width="130">��й�ȣ</td>
			<td bgcolor="#FFFFFF">
				<input type="password" name="sEP" class="text" size="20" maxlength="60" >
				&nbsp;
				(��� �α��ο�, ����ø� �Է�)
				<div style="font-size:11px;color:gray;">�ּ�8�� �̻�, �������� ����, �������� 3�� ���ӱ���</div>
			</td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>" width="130">�̸�<font color="red">(*)</font></td>
			<td bgcolor="#FFFFFF"><input type="text" name="sUN" class="text" size="20" maxlength="60" value="<%=susername%>"></td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>" width="130">�����̸�<font color="red">(*)</font></td>
			<td bgcolor="#FFFFFF"><input type="text" name="userNameEN" class="text" size="20" maxlength="60" value="<%= userNameEN %>"></td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>" width="130" >�ֹε�Ϲ�ȣ<font color="red">(*)</font></td>
			<td bgcolor="#FFFFFF"  ><input type="text" name="sJN1" class="text" size="6" maxlength="6" value="<%=sjuminno1%>">-
			    <input type="text" name="sJN2" class="text" size="1" maxlength="1" value="<%=LEFT(sjuminno2,1)%>" onFocusOut="jumin_format()">******
			    (���ڸ��� �־��ֽñ� �ٶ��ϴ�.)
			<% if (FALSE) then %><input type="password" name="sJN2" class="text" size="7" maxlength="7" value="<%=sjuminno2%>" onFocusOut="jumin_format()"><% end if %></td>
		</tr>
    		<tr align="left" height="25">
    			<td bgcolor="<%= adminColor("tabletop") %>">�������</td>
    			<td bgcolor="#FFFFFF" colspan="2">
    			    <label><input type="text" name="selBD_y" class="text" value="<%=year(dbirthday)%>" size="4" maxlength="4" />��</label>
					<label><input type="text" name="selBD_m" class="text" value="<%=month(dbirthday)%>" size="2" maxlength="2" />��</label>
					<label><input type="text" name="selBD_d" class="text" value="<%=day(dbirthday)%>" size="2" maxlength="2" />��</label>
    				<label><input type="radio" name="rdoS" value="Y" <%=chkIIF(blnissolar="Y","checked","")%>> ���</label>
					<label><input type="radio" name="rdoS" value="N" <%=chkIIF(blnissolar="N","checked","")%>> ����</label>
    			</td>
    		</tr>
    		<tr align="left" height="25">
    		<td bgcolor="<%= adminColor("tabletop") %>">����</td>
    		<td bgcolor="#FFFFFF" colspan="2"><input type="radio" name="rdoSf" value="M" <%IF blnsexflag ="M" THEN%>checked<%END IF%>> ��  <input type="radio" name="rdoSf" value="F" <%IF blnsexflag ="F" THEN%>checked<%END IF%>> ��</td>
    	</tr>
    		<tr align="left" height="25">
    			<td bgcolor="<%= adminColor("tabletop") %>">�ڵ�����ȣ</td>
    			<td bgcolor="#FFFFFF" colspan="2">
					<input type="text" name="sUC" size="16" class="text" onFocusOut="phone_format(frm_member.sUC)" value="<%=sUsercell%>">
					<font color="<%=chkIIF(isIdentify="Y","red","blue")%>">[<%=chkIIF(isIdentify="Y","�����Ϸ�","��������")%>]</font>
				</td>
    		</tr>
    		<tr align="left" height="25">
    			<td bgcolor="<%= adminColor("tabletop") %>">����ȭ��ȣ</td>
    			<td bgcolor="#FFFFFF" colspan="2"><input type="text" name="sUP" size="16" class="text"  onFocusOut="phone_format(frm_member.sUP)" value="<%=sUserPhone%>"></td>
    		</tr>
    		<tr align="left" height="25">
    			<td bgcolor="<%= adminColor("tabletop") %>">�����ȣ</td>
    			<td bgcolor="#FFFFFF" colspan="2">
    				<input type="text" name="zipcode" size="16" class="text_ro" value="<%=szipcode%>">
    				<input type="button" class="button" value="�˻�" onClick="FnFindZipNew('frm_member','B')">
					<input type="button" class="button" value="�˻�(��)" onClick="TnFindZipNew('frm_member','B')">
    				<% '<input type="button" class="button" value="�˻�(��)" onClick="javascript:PopSearchZipcode('frm_member');"> %>
    			</td>
    		</tr>
    		<tr align="left" height="25">
    			<td bgcolor="<%= adminColor("tabletop") %>">�ּ�</td>
    			<td bgcolor="#FFFFFF" colspan="2">
    				<input type="text" name="zipaddr" size="50" class="text_ro" value="<%=szipaddr%>">
    				<br><input type="text" name="useraddr" size="60" maxlength="60" class="text" value="<%=suseraddr%>">
    			</td>
    		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="3" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<tr align="left" height="25" >
			<td bgcolor="<%= adminColor("tabletop") %>" width="130" >�Ի���(������)<font color="red">(*)</font></td><!-- �ű� -->
			<td bgcolor="#FFFFFF">
					<input type="hidden" name="selJD_y" value="<%=Year(djoinday)%>">
					<input type="hidden" name="selJD_m" value="<%=month(djoinday)%>">
					<input type="hidden" name="selJD_d" value="<%=day(djoinday)%>">
					<%=djoinday%>
			</td>
		</tr>
			<tr align="left" height="25" >
			<td bgcolor="<%= adminColor("tabletop") %>" width="130">�����Ի���</td>
			<td bgcolor="#DDDDFF">
	    		<select name="selRJD_y">
	    			<option value="">-����-</option>
	<% for i = Year(dateadd("yyyy",1,now()))   to  2001 step -1%>
	    			<option value="<%= i %>" <% if (Year(drealjoinday) = i) then %>selected<% end if %>><%= i %></option>
	<% next %>
	    		</select>
	    		<select name="selRJD_m">
	    			<option value="">-����-</option>
	<% for i = 1 to 12 %>
	    			<option value="<%= i %>" <% if (Month(drealjoinday) = i) then %>selected<% end if %>><%= i %></option>
	<% next %>
	    		</select>
	    		<select name="selRJD_d">
	    			<option value="">-����-</option>
	<% for i = 1 to 31 %>
	    			<option value="<%= i %>" <% if (Day(drealjoinday) = i) then %>selected<% end if %>><%= i %></option>
	<% next %>
	    		</select>
			</td>
		</tr>

		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">E-MAIL(�系����)</td>
			<td bgcolor="#FFFFFF">
				<input type="text" name="sUM" class="text" size="30" maxlength="80"  value="<%=susermail%>">
			</td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">E-MAIL(���θ���)</td>
			<td bgcolor="#FFFFFF">
				<input type="text" name="sPM" class="text" size="30" maxlength="80"  value="<%=ipersonalmail%>">
			</td>
		</tr>
	  <tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">��ȭ��ȣ(����)</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sCUP" class="text" size="16" maxlength="16" value="<%=sinterphoneno%>">

				&nbsp;&nbsp;
				����: <input type="text" name="sCE" class="text" size="4" maxlength="16"  value="<%=sextension%>">
			</td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">070 �����ȣ</td>
		    	<td bgcolor="#FFFFFF"><input type="text" name="sD070" class="text" size="16" maxlength="16"  value="<%=sdirect070%>"></td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">�ٹ����ٻ���Ʈ ���̵�</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sFUI" class="text" size="20" maxlength="32" value="<%= sfrontid %>"></td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">GSSHOP���̵�</td>
			<td bgcolor="#FFFFFF"><input type="text" name="gsshopuserid" class="text" size="20" maxlength="32" value="<%= gsshopuserid %>"></td>
		</tr>
    		<!--<tr align="left" height="25">
    			<td bgcolor="<%= adminColor("tabletop") %>">MSN�޽���</td>
    			<td bgcolor="#FFFFFF"><input type="text" name="sMM" class="text" size="30" maxlength="80" value="<%=smsnmail%>"></td>
    		</tr>
    		<tr align="left" height="25">
    			<td bgcolor="<%= adminColor("tabletop") %>">NateOn</td>
    			<td bgcolor="#FFFFFF"><input type="text" name="sNt" class="text" size="30" maxlength="80" value="<%=smessenger%>"></td>
    		</tr>-->
			<tr align="left" height="25">
    			<td bgcolor="<%= adminColor("tabletop") %>">������<br>(ī�װ�or��)</td>
    			<td bgcolor="#FFFFFF">
    				<% SelectBoxBrandCategory "selC", sjobdetail %>
    			</td>
    		</tr>
			<tr align="left" height="25">
				<td bgcolor="<%= adminColor("tabletop") %>">������</td>
				<td bgcolor="#FFFFFF"><input type="text" name="smywork" class="text" size="60" value="<%=mywork%>"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="3" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>" width="130">�μ�</td>
			<td bgcolor="#FFFFFF">
				<%= drawSelectBoxDepartment("department_id", idepartment_id) %>
			</td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">����<font color="red">(*)</font></td>
			<td bgcolor="#FFFFFF">
				<select name="selRank"><%=fnRankInfoSelectBox(irank_sn)%></select>
			</td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">����<font color="red">(*)</font></td>
			<td bgcolor="#FFFFFF">
				<%IF menupos = "1176"  and left(sEmpNo,2) ="10" THEN	'��������ϋ��� ������ü/ ��������������� �������������%>
				<%=printPositOption("selPoN", iposit_sn)%>
				<%ELSE%>
				<%=printPositOptionPartTime("selPoN", iposit_sn)%>
				<%END IF%>
			</td>
		</tr>
		<%IF menupos = "1176" THEN	'��������϶���%>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">��å</td>
			<td bgcolor="#FFFFFF">
				<%=printJobOption("selJN", ijob_sn)%>
			</td>
		</tr>
		<%END IF%>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>" width="130">�߷���</td>
			<td bgcolor="#FFFFFF">
				<input type="text" class="formTxt" id="chdate" name="chdate" style="width:100px" maxlength="10"   value="<%=changedate%>"/>
				<input type="image" name="chdate_trigger" id="chdate_trigger" src="/images/admin_calendar.png" alt="�޷����� �˻�"  onclick="return false;" />
				<script type="text/javascript">
						var CAL_Start = new Calendar({
							inputField : "chdate", trigger    : "chdate_trigger",
							onSelect: function() {
								var date = Calendar.intToDate(this.selection.get());
								CAL_End.args.min = date;
								CAL_End.redraw();
								this.hide();
							}, bottomBar: true, dateFormat: "%Y-%m-%d"
						});

					</script>
			</td>
		</tr>
    		</table>
    	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="3" cellspacing="1" align="center" class="a" bgcolor="#BABABA">
		<tr align="left" height="25" >
			<td bgcolor="<%= adminColor("tabletop") %>" width="130">�����</td>

			<input type="hidden" name="statediv" value="<%= blnstatediv %>">
			<input type="hidden" name="retirereason" value="<%= iretirereason %>">

			<td bgcolor="#DDDDFF">
				<input type="radio" name="tmpstatediv" value="N" <% if blnstatediv = "N" and iretirereason <> "99" then %>checked<% end if %> > ���
				<input type="radio" name="tmpstatediv" value="C" <% if blnstatediv = "N" and iretirereason = "99" then %>checked<% end if %> <% if left(sEmpNo,2) <> "90" then %>disabled<% end if %> > ��������ȯ
				 &nbsp;
				����� : <% DrawOneDateBoxdynamic "selRD_y", selRD_y, "selRD_m", selRD_m, "selRD_d", selRD_d, "", "", "", "" %>
				<% if blnstatediv = "N" and iretirereason <> "99" then %>
					<% If (C_ADMIN_AUTH) or C_PSMngPart Then %>
						<input type="button" class="button" value="����������� ����" onClick="jsChangeretire('<%=suserid%>');">
					<% End If %>
				<% end if %>
			</td>
		</tr>
		<% if left(sEmpNo,2) ="90" then %>
		<tr>
			<td bgcolor="<%= adminColor("tabletop") %>" width="130">������</td>
			<td bgcolor="#DDDDFF" height="25" >
				<input type="radio" name="tmpretirereason" value="1" <% if iretirereason ="1" then %>checked<% end if %>/>���λ���&nbsp;
				<input type="radio" name="tmpretirereason" value="2" <% if iretirereason ="2" then %>checked<% end if %>/>���Ⱓ����&nbsp;
				<input type="radio" name="tmpretirereason" value="3" <% if iretirereason ="3" then %>checked<% end if %>/>�ǰ����&nbsp;
				<input type="radio" name="tmpretirereason" value="4" <% if iretirereason ="4" then %>checked<% end if %>/>�ذ�&nbsp;
				<input type="radio" name="tmpretirereason" value="5" <% if iretirereason ="5" then %>checked<% end if %>/>������ ����&nbsp;
				<input type="radio" name="tmpretirereason" value="6" <% if iretirereason ="6" then %>checked<% end if %>/>��Ÿ
			</td>
		</tr>
		<% end if %>
	</td>
	</table>
</tr>

<tr align="center" height="25">
	<td >
	<!-- �ּ�ó�� : ������
		<input type="button" class="button" value="����" onClick="jsDelUser()">&nbsp;&nbsp;&nbsp;
	-->

		<% 'if isdispmember then %>
			<input type="submit" class="button" value="Ȯ��">&nbsp;&nbsp;
			<input type="button" class="button" value="���" onClick="self.close()">
			<% If (C_ADMIN_AUTH) OR (C_PSMngPart) Then %>
			&nbsp;&nbsp;<input type="button" class="button" value="����" onclick="updateIsusingZero()" style=color:red;font-weight:bold>
			<% End If %>
		<% 'End If %>
	</td>
</tr>
</table>
</form>
<form name="frmchk" method="post" action="/admin/member/tenbyten/member_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="sEN" value="">
</form>
<form name="frmDel" method="post" action="member_process.asp" style="margin:0px;">
<input type="hidden" name="sEN" value="<%=sEmpNo%>">
<input type="hidden" name="mode" value="D">
</form>
<form name="frmDel2" method="post" action="member_process.asp" style="margin:0px;">
<input type="hidden" name="sEN" value="<%=sEmpNo%>">
<input type="hidden" name="mode" value="S">
</form>
<form name="frmlogDel" method="post" action="member_process.asp" style="margin:0px;">
<input type="hidden" name="sEN" value="<%=sEmpNo%>">
<input type="hidden" name="logidx" value="">
<input type="hidden" name="mode" value="LD">
</form>
<div style="padding:10px;">
<div>+�߷�����<hr width="100%"></div>
<table width="100%" border="0" cellpadding="3" cellspacing="1" align="center" class="a" bgcolor="#BABABA">
<tr bgcolor="<%= adminColor("tabletop") %>" height="25" align="center">
	<Td>�߷���</td>
	<Td>�μ�</td>
	<Td>��å</td>
	<Td>����</td>
    <td>ó��</td>
</tr>
<%IF isArray(arrList) THEN
	For intLoop = 0 To ubound(arrList,2)
%>
<tr bgcolor="#ffffff">
	<td align="center"><%=arrList(4,intLoop)%></td>
	<td><%=arrList(7,intLoop)%></td>
	<td align="center"><%=arrList(6,intLoop)%></td>
	<td align="center"><%=arrList(5,intLoop)%></td>
	<td align="center">
		<% 'if isdispmember then %>
			<%if intLoop <> 0 and (C_PSMngPart or C_ADMIN_AUTH) then%>
				<input type="button" class="button" value="�߷����" onClick="jsDelModLog(<%=arrList(8,intLoop)%>);">
			<%end if%>
		<% 'end if %>
	</td>
</tr>
<%	Next
END IF%>
</table>
</div>
</body>
</html>
<%
set oAddLevel = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
