<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ��Ʈ��� �������� ����
' History : 2007.07.30 �ѿ�� ����
'		2011.01.12 ������ ���ν��� ����
'       2011.05.30 ������ ����Ȯ���� ����� �޴�����ȣ �������� ����
'###########################################################
%>
<!-- #include virtual="/tenmember/incSessionTenMember.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/tenmember/lib/header.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member/10x10staffcls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVacationCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVitaminCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenAgitCls.asp" -->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<%
dim userid, empno, susercell, susermail, smsnmail, sinterphoneno, sextension, sdirect070, sjobdetail, sjuminno
Dim sfrontid, susername, dbirthday , blnissolar, szipcode, blnsexflag, szipaddr,suseraddr,  suserphone
Dim blnstatediv, djoinday, dretireday, suserimage, ipart_sn, iposit_sn, ijob_sn, ilevel_sn, iuserdiv
Dim spart_name, sposit_name, sjob_name, slevel_name, sMessenger,sdepartmentNameFull
Dim smywork, isIdentify,sbizsection_Cd, arrList, intLoop, cMember, i, userNameEN
dim isISViewValid : isISViewValid = (LCASE(session("ssBctId"))<>"iiitester")
	empno = session("ssBctSn")

'����������ȣ�� ���� �н������ �ѹ��� Check
dim checkedPass, param, refUrl
checkedPass = session("chkSCMMyInfoPass")
param = request.queryString
refUrl = request.ServerVariables("HTTP_REFERER")

If InStr(refUrl,"/tenmember/member/modify_myinfo_process.asp") > 0 then
	checkedPass = now()
End If

If (checkedPass="") Then
	response.redirect getSCMSSLURL & "/tenmember/member/confirmuser.asp?" & param
	response.end
Else
	session("chkSCMMyInfoPass") = ""
End If

Set cMember = new CTenByTenMember
	cMember.Fempno = empno
	cMember.fnGetMemberData

	empno   		= cMember.Fempno
	userid			= cMember.Fuserid
	sfrontid       	= cMember.Ffrontid
	susername      	= cMember.Fusername
	userNameEN      = cMember.FuserNameEN
	sjuminno		= cMember.FJuminno
	dbirthday       = cMember.Fbirthday
	blnissolar      = cMember.Fissolar
	blnsexflag		= cMember.Fsexflag
	szipcode        = cMember.Fzipcode
	szipaddr		= cMember.Fzipaddr
	suseraddr      	= cMember.Fuseraddr
	suserphone     	= cMember.Fuserphone
	susercell      	= cMember.Fusercell
	susermail      	= cMember.Fusermail
	smsnmail       	= cMember.Fmsnmail
	sinterphoneno 	= cMember.Finterphoneno
	sextension    	= cMember.Fextension
	sdirect070     	= cMember.Fdirect070
	sjobdetail     	= cMember.Fjobdetail
	blnstatediv    	= cMember.Fstatediv
	djoinday       	= cMember.Fjoinday
	dretireday    	= cMember.Fretireday
	suserimage  	= cMember.Fuserimage
	smywork    		= cMember.Fmywork
	ipart_sn       	= cMember.Fpart_sn
	iposit_sn     	= cMember.Fposit_sn
	ijob_sn        	= cMember.Fjob_sn
	ilevel_sn      	= cMember.Flevel_sn
	iuserdiv        = cMember.Fuserdiv
	spart_name     	= cMember.Fpart_name
	sposit_name     = cMember.Fposit_name
	sjob_name      	= cMember.Fjob_name
	slevel_name		= cMember.Flevel_name
	sdepartmentNameFull	= cMember.FdepartmentNameFull

	isIdentify		= cMember.FisIdentify
	sMessenger		= cMember.Fmessenger
	sbizsection_Cd	= cMember.Fbizsection_cd
	IF isNull(sbizsection_Cd) THEN sbizsection_Cd = ""
	cMember.Fempno = empno
	arrList = cMember.fnGetUserBizSection
Set cMember = nothing

dim birthday_yyyy, birthday_mm, birthday_dd

if (Not IsNull(dbirthday) and (dbirthday <> "")) then
	birthday_yyyy = Year(dbirthday)
	birthday_mm = Month(dbirthday)
	birthday_dd = Day(dbirthday)
end if

dim joinday_yyyy, joinday_mm, joinday_dd

if ((djoinday) and (djoinday <> "")) then
	joinday_yyyy = Year(djoinday)
	joinday_mm = Month(djoinday)
	joinday_dd = Day(djoinday)
end if

dim blnView,blnSale,clsBS,arrBiz, totalvacationday, expiredday, requestedday, usedvacationday
 	blnView = "Y"
	blnSale = "N"

Set clsBS = new CBizSection
	clsBS.FUSE_YN = "Y"
	clsBS.FOnlySub = "Y"
	clsBS.FView		= blnView
	clsBS.FSale		= blnSale
	arrBiz = clsBS.fnGetBizSectionList
Set clsBS = nothing

dim IS_SHOW_SCM_INFO
IS_SHOW_SCM_INFO = False

if userid <> "" then
	IS_SHOW_SCM_INFO = True
end if

Dim oAddLevel
set oAddLevel = new CPartnerAddLevel
	oAddLevel.FRectUserid=userid
	oAddLevel.FRectOnlyAdd = "on"

	if (oAddLevel.FRectUserID<>"") then
	    oAddLevel.getUserAddLevelList
	end if
%>

<script type="text/javascript">

function SaveBaseInfo() {
	var frm = document.frm_base;

	frm.birthday.value = frm.birthday_yyyy.value + "-" + frm.birthday_mm.value + "-" + frm.birthday_dd.value;

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret) {
		frm.submit();
	}
}

function OpenVacationList()
{
	var win = window.open("pop_tenbyten_vacation_list.asp","OpenVacationList","width=750,height=500,scrollbars=yes");
	win.focus();
}

function OpenVacationListAdmin()
{
	var win = window.open("/admin/member/tenbyten/pop_tenbyten_vacation_list_admin.asp","OpenVacationListAdmin","width=900,height=500,scrollbars=yes");
	win.focus();
}

function SaveAddressInfo() {
	var frm = document.frm_addr;

	// ========================================================================
	if (frm.usercell.value == ''){
		alert("�޴�����ȣ�� �����ϴ�. �޴�����ȣ ���� ��ư�� ���� �������ּ���.");
		return;
	}

	if (frm.userphone.value == ''){
		alert("����ȭ��ȣ�� �Է��ϼ���");
		frm.userphone.focus();
		return;
	}

	if ((frm.zipcode.value == '') || (frm.useraddr.value == '')) {
		alert("�ּҸ� �Է��ϼ���");
		frm.useraddr.focus();
		return;
	}
	// ========================================================================

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret) {
		frm.submit();
	}
}

function SaveAuthInfo() {
	var frm = document.frm_auth;

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret) {
		frm.submit();
	}
}

function SavePassInfo() {
	var frm = document.frm_mypass;

	if (frm.olduserpass.value == ''){
		alert("������й�ȣ�� �Է��ϼ���");
		frm.olduserpass.focus();
		return;
	}

	if (frm.newuserpass.value == ''){
		alert("�űԺ�й�ȣ�� �Է��ϼ���");
		frm.newuserpass.focus();
		return;
	}

	if (frm.newuserpass.value != frm.newuserpass1.value){
		alert("�űԺ�й�ȣ�� ���� ��ġ���� �ʽ��ϴ�.");
		frm.newuserpass.focus();
		return;
	}

	if (!fnChkComplexPassword(frm.newuserpass.value)) {
		alert('���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)');
		frm.newuserpass.focus();
		return;
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret) {
		frm.submit();
	}
}

function SaveEmpPassInfo() {
	var frm = document.frm_myemppass;

	if (frm.olduserpass.value == ''){
		alert("������й�ȣ�� �Է��ϼ���");
		frm.olduserpass.focus();
		return;
	}

	if (frm.newuserpass.value == ''){
		alert("�űԺ�й�ȣ�� �Է��ϼ���");
		frm.newuserpass.focus();
		return;
	}

	if (frm.newuserpass.value != frm.newuserpass1.value){
		alert("�űԺ�й�ȣ�� ���� ��ġ���� �ʽ��ϴ�.");
		frm.newuserpass.focus();
		return;
	}

	if (!fnChkComplexPassword(frm.newuserpass.value)) {
		alert('���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)');
		frm.newuserpass.focus();
		return;
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret) {
		frm.submit();
	}
}


function SaveMoreInfo() {
	var frm = document.frm_moreinfo;

	frm.joinday.value = frm.joinday_yyyy.value + "-" + frm.joinday_mm.value + "-" + frm.joinday_dd.value;

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret) {
		frm.submit();
	}
}


function SaveUserImage()
{
	//alert(frm_base.userimage.value);
	var frm = document.frm_base;

	frm.birthday.value = frm.birthday_yyyy.value + "-" + frm.birthday_mm.value + "-" + frm.birthday_dd.value;

	frm.submit();
}


function PopSearchZipcode(frmname) {
	var popwin = window.open("/lib/searchzip3.asp?target=" + frmname,"PopSearchZipcode","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}


function CopyZip(frmname, post1, post2, addr, dong) {
    eval(frmname + ".zipcode").value = post1 + "-" + post2;

    eval(frmname + ".zipaddr").value = addr;
    eval(frmname + ".useraddr").value = dong;
}

// �޴�����ȣ ����/����Ȯ�� �˾�
function PopChgHPNum() {
	var popwin = window.open("pop_ChangeHPIdentify.asp","PopChgHPNum","width=400 height=270 scrollbars=yes");
	popwin.focus();
}

//�������� ���
function jsSetUserBiz(sDate){
	var winBiz = window.open("pop_member_bizsection_reg.asp?sD="+sDate,"popBiz","width=630 height=800 scrollbars=yes");
	winBiz.focus();
}

//��Ÿ�ν�û
function OpenVitamin(){
//var winVM = window.open("/admin/approval/eapp/regeapp.asp?iAidx=351&ieidx=33","popNew","width=880, height=600,scrollbars=yes, resizable=yes");
var winVM = window.open("<%=manageUrl%>/tenmember/member/vitamin/popRegVitamin.asp","popNew","width=400, height=400,scrollbars=yes, resizable=yes");
	winVM.focus();
}

//��Ÿ�� ��������
function listVitamin(){
	var winVM = window.open("<%=manageUrl%>/tenmember/member/vitamin/popListVitamin.asp","popNew","width=880, height=600,scrollbars=yes, resizable=yes");
	winVM.focus();
}

//����Ʈ��û
function OpenAgit(){
	var winap = window.open("<%=manageUrl%>/admin/member/tenbyten/Agit/tenbyten_agit_calendar.asp","popAgit","");
	winap.focus();
}


//����Ʈ��û
function listAgit(){
	var winap = window.open("<%=manageUrl%>/admin/member/tenbyten/Agit/pop_agit_mylist.asp","popAList","width=1024, height=600,scrollbars=yes, resizable=yes");
	winap.focus();
}
</script>

<!--�⺻�������� ����-->
<table border="0" width="100%" cellpadding="10" cellspacing="0" class="a">
	<tr>
		<td valign="top">
		<table border=0 width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm_base" method="post" action="modify_myinfo_process.asp">
			<input type="hidden" name="mode" value="base">
			<input type="hidden" name="userimage" value="<%= sUserImage%>">
			<tr height="25" bgcolor="FFFFFF">
				<td colspan="2" height="35">
					<font color="red"><strong>[�⺻����]</strong></font>
				</td>
			</tr>
			<tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">���</td>
		    	<td bgcolor="#FFFFFF">
		    		<%= empno %>
		    	</td>
		    </tr>
			<tr align="left" height="25">
		    	<td width="120" bgcolor="<%= adminColor("tabletop") %>">�̸�</td>
		    	<td bgcolor="#FFFFFF">
		    		<%=susername %>
		    	</td>
		    </tr>
			<tr align="left" height="25">
		    	<td width="120" bgcolor="<%= adminColor("tabletop") %>">�����̸�</td>
		    	<td bgcolor="#FFFFFF">
		    		<input type="text" name="userNameEN" class="text" size="20" maxlength="60" value="<%= userNameEN %>">
		    	</td>
		    </tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">���ξ��̵�</td>
		    	<td bgcolor="#FFFFFF">
		    		<%=userid %>
		    	</td>
		    </tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">�ٹ����� ���̵�</td>
		    	<td bgcolor="#FFFFFF">
		    		<%=sfrontid %>
		    	</td>
		    </tr>
			<tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">E-MAIL(�系����)</td>
		    	<td bgcolor="#FFFFFF">
		    		<%= susermail %>
		    	</td>
		    </tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">��ȭ��ȣ(����)</td>
		    	<td bgcolor="#FFFFFF">
		    		<%= sinterphoneno %>
		    		<input type="hidden" name="interphoneno" value="<%= sinterphoneno %>">
		    		&nbsp;&nbsp;
		    		����: <input type="text" name="extension" size="5" class="text_ro" value="<%= sextension %>" readonly>
		    	</td>
		    </tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">070 �����ȣ</td>
		    	<td bgcolor="#FFFFFF">
		    		<input type="text" name="direct070" id="" size="16" class="text_ro" value="<%= sdirect070 %>" readonly>
		    	</td>
		    </tr>
		    <input type="hidden" name="birthday" value="">
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">�������</td>
		   	<td bgcolor="#FFFFFF">
		    		<select name=birthday_yyyy>
		<% for i = 1960 to year(date)-14 %>
		    			<option value="<%= i %>" <% if (birthday_yyyy = i) then %>selected<% end if %>><%= i %></option>
		<% next %>
		    		</select>
		    		<select name=birthday_mm>
		<% for i = 1 to 12 %>
		    			<option value="<%= i %>" <% if (birthday_mm = i) then %>selected<% end if %>><%= i %></option>
		<% next %>
		    		</select>
		    		<select name=birthday_dd>
		<% for i = 1 to 31 %>
		    			<option value="<%= i %>" <% if (birthday_dd = i) then %>selected<% end if %>><%= i %></option>
		<% next %>
		    		</select>
					&nbsp; &nbsp; &nbsp; &nbsp;
					<input type="radio" name="issolar" value="Y" <% if  blnissolar = "Y" then response.write "checked" %>> ���
					<input type="radio" name="issolar" value="N" <% if blnissolar= "N" then response.write "checked" %>> ����
		    	</td>
		    </tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		    	<td bgcolor="#FFFFFF">
					<input type="radio" name="sexflag" value="M" <% if blnsexflag = "M" then response.write "checked" %>> ����
					<input type="radio" name="sexflag" value="F" <% if blnsexflag = "F" then response.write "checked" %>> ����
		    	</td>
		    </tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">MSN�޽���</td>
		    	<td bgcolor="#FFFFFF">
		    		<input name="msnmail" type="text" size="45" class="text" value="<%= smsnmail %>">
		    	</td>
		    </tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">NateOn</td>
		    	<td bgcolor="#FFFFFF">
		    		<input name="messenger" type="text" size="45" class="text" value="<%= sMessenger %>">
		    	</td>
		    </tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">������(ī�װ�)</td>
		    	<td bgcolor="#FFFFFF">
		    		<%= sjobdetail %>
		    	</td>
		    </tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">������</td>
		    	<td bgcolor="#FFFFFF">
		    		<input name="mywork" type="text" size="45" class="text" maxlength="80" value="<%= smywork %>">
		    	</td>
		    </tr>
		    </form>
		    <tr align="left" height="35">
		    	<td colspan="2" bgcolor="#FFFFFF" align=center>
					<input type="button" class="button" value="�⺻���� ����" onclick="javascript:SaveBaseInfo()">
					&nbsp;&nbsp;&nbsp;
					<input type="button" class="button" value="����<% If sUserImage = "" Then %>���<% Else %>����<% End If %>" onclick="javascript:window.open('popAddImage.asp?sF=<%=session("ssAdminPsn")%>','myimageupload','width=380,height=150');">
		    	</td>
		    </tr>
			<tr>
				<td valign="bottom" colspan=2 bgcolor="FFFFFF" height="35">
					<font color="red"><strong>[����ó ����]</strong></font>
				</td>
			</tr>
			<form name="frm_addr" method="post" action="modify_myinfo_process.asp">
			<input type="hidden" name="mode" value="addr">
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">�޴�����ȣ</td>
		    	<td bgcolor="#FFFFFF">
		    		<input type="text" name="usercell" size="16" class="text_ro" value="<%= susercell %>" readonly onFocusOut="phone_format(frm_addr.usercell)">
		    		<input type="button" class="button_s" value="�޴�����ȣ ����" onClick="javascript:PopChgHPNum();" style="width:100px;">
		    		<% IF isIdentify="Y" then %>
		    		 <font color=darkred>����Ȯ�� ��</font>
		    		 <% ELSE%>
		    		  �޴�����ȣ ���� ��ư�� �����ż� ����Ȯ���� ���ּ���
		    		 <%END IF%>
		    	</td>
		    </tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">����ȭ��ȣ</td>
		    	<td bgcolor="#FFFFFF">
		    		<input type="text" name="userphone" size="16" class="text" value="<%= suserphone %>" onFocusOut="phone_format(frm_addr.userphone)">
		    	</td>
		    </tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">�����ȣ</td>
		    	<td bgcolor="#FFFFFF">
		    		<input type="text" name="zipcode" size="16" class="text_ro" value="<%= szipcode %>">
					<input type="button" class="button_s" value="�˻�" onClick="FnFindZipNew('frm_addr','B')">
					<input type="button" class="button_s" value="�˻�(��)" onClick="TnFindZipNew('frm_addr','B')">
					<% '<input type="button" class="button_s" value="�˻�(��)" onClick="javascript:PopSearchZipcode('frm_addr');"> %>
		    	</td>
		    </tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">�ּ�</td>
		    	<td bgcolor="#FFFFFF">
		    		<input type="text" name="zipaddr" size="50" class="text_ro" value="<%= szipaddr %>"><br>
		    		<input type="text" name="useraddr" size="50" maxlength="128" class="text"  value="<%= suseraddr %>">
		    	</td>
		    </tr>
		    </form>
		    <tr align="left" height="35">
		    	<td colspan="2" bgcolor="#FFFFFF" align=center>
					<input type="button" class="button" value="����ó ����" onclick="javascript:SaveAddressInfo()">
		    	</td>
		    </tr>
			<tr>
				<td valign="bottom" colspan=2 bgcolor="FFFFFF" height="35">
					<font color="red"><strong>[��������]</strong></font>
				</td>
			</tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">�μ�-��Ʈ</td>
		    	<td bgcolor="#FFFFFF">
		    		<%= sdepartmentNameFull %>
		    	</td>
		    </tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">���α���</td>
		    	<td bgcolor="#FFFFFF">
		    			<div style="padding:1px;"><%= spart_name %>&nbsp;<%= slevel_name %> </div>
		    		<% if (oAddLevel.FResultCount>0) then %>
        			  <% for i=0 to oAddLevel.FResultCount-1 %>
        						 <div style="color:Gray;padding:1px;"><%= oAddLevel.FItemList(i).Fpart_name %> &nbsp;<%= oAddLevel.FItemList(i).Flevel_name %> </div>
        				<% next %>
        		<%end if%>
        		</div>
		    	</td>
		    </tr>
		    <!--
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		    	<td bgcolor="#FFFFFF">
		    		<%= sposit_name %>
		    	</td>
		    </tr>
		    -->
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">��å</td>
		    	<td bgcolor="#FFFFFF">
		    		<%= sjob_name %>
		    	</td>
		    </tr>

			<!--��й�ȣ���� ����-->
			<% if IS_SHOW_SCM_INFO then %>
			<form name="frm_mypass" method="post" action="modify_myinfo_process.asp">
			<input type="hidden" name="mode" value="mypass">
			<tr>
				<td valign="bottom" colspan=2 bgcolor="FFFFFF" height="35">
					<font color="red"><strong>[���� ���̵� �α��� ��й�ȣ]</strong></font>
				</td>
			</tr>
			<tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">������й�ȣ</td>
		    	<td bgcolor="#FFFFFF">
		    		<input  type="password" name="olduserpass" size="16" class="input_01">
		    	</td>
		    </tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">�űԺ�й�ȣ</td>
		    	<td bgcolor="#FFFFFF">
		    		�Է� : <input  type="password" name="newuserpass" size="16" class="input_01"><br>
					Ȯ�� : <input  type="password" name="newuserpass1" size="16" class="input_01">
					<br>
					���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)
		    	</td>
		    </tr>
		    </form>
		    <tr align="center" height="35">
		    	<td colspan="2" bgcolor="#FFFFFF">
		    		<input type="button" class="button_s" value="��й�ȣ ����" onclick="javascript:SavePassInfo()">
		    	</td>
		    </tr>
			<% end if %>

			<form name="frm_myemppass" method="post" action="modify_myinfo_process.asp">
			<input type="hidden" name="mode" value="myemppass">
			<tr>
				<td valign="bottom" colspan=2 bgcolor="FFFFFF" height="35">
					<font color="red"><strong>[���� ��� �α��� ��й�ȣ]</strong></font>
				</td>
			</tr>
			<tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">������й�ȣ</td>
		    	<td bgcolor="#FFFFFF">
		    		<input  type="password" name="olduserpass" size="16" class="input_01">
		    	</td>
		    </tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">�űԺ�й�ȣ</td>
		    	<td bgcolor="#FFFFFF">
		    		�Է� : <input  type="password" name="newuserpass" size="16" class="input_01"><br>
					Ȯ�� : <input  type="password" name="newuserpass1" size="16" class="input_01">
					<br>
					���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)
		    	</td>
		    </tr>
		    </form>
		    <tr align="center" height="35">
		    	<td colspan="2" bgcolor="#FFFFFF">
		    		<input type="button" class="button_s" value="��й�ȣ ����" onclick="javascript:SaveEmpPassInfo()">
		    	</td>
		    </tr>
			<!--��й�ȣ���� ��-->

		</table>
	</td>
	<td valign="top">
		<table width="100%"  cellpadding="0" cellspacing="0" class="a" bgcolor="#FFFFFF">
		<tr>
			<td>
				<table width="100%"   cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
					<tr height="25" bgcolor="FFFFFF">
						<td colspan="4">
							<font color="red"><strong>�߰�����</strong></font>
						</td>
					</tr>
					<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
				    	<td width="150">�Ի���</td>
				    	<td width="100">�ټӿ���</td>
				    	<td></td>
				      	<td></td>
				    </tr>
				    <input type="hidden" name="joinday" value="">
				    <tr height="25" align="center" bgcolor="#FFFFFF">
				    	<td>
				    		<%= Left(djoinday, 10) %>
				    	</td>
				    	<td><%= GetYearDiff(djoinday) %></td>
				      	<td></td>
				      	<td></td>
				    </tr>
				</table>
			</td>
		</tr>
		<tr>
			<td style="padding-top:20px;">
				<table width="100%"   cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
					<tr height="25" bgcolor="FFFFFF">
						<td colspan="15">
							<font color="red"><strong>����(�ް�)����</strong></font>
						</td>
					</tr>
					<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
				    	<td>����</td>
				    	<td>�� �� ��</td>
				      	<td>����ϼ�</td>
				      	<td>���δ��</td>
				      	<td>�ܿ��ϼ�</td>
				      	<td>�����ϼ�</td>
				    </tr>
				<%
				if iposit_sn = 12 or iposit_sn = 13 then
					i = GetPartYearVacationDayByEmpno(empno, totalvacationday, usedvacationday, requestedday, expiredday)
				%>
				 	<tr align="center" bgcolor="#FFFFFF">
				 		<td>�ް�</td>
				 		<td>
							<%= GetDayOrHourWithPositSN(iposit_sn, totalvacationday) %> <%= GetDayOrHourNameWithPositSN(iposit_sn) %>
						</td>
				      	<td>
							<%= GetDayOrHourWithPositSN(iposit_sn, usedvacationday) %> <%= GetDayOrHourNameWithPositSN(iposit_sn) %>
						</td>
				      	<td>
							<%= GetDayOrHourWithPositSN(iposit_sn, requestedday) %> <%= GetDayOrHourNameWithPositSN(iposit_sn) %>
						</td>
				      	<td>
							<b>
								<%= GetDayOrHourWithPositSN(iposit_sn, (totalvacationday - (usedvacationday + requestedday+expiredday))) %> <%= GetDayOrHourNameWithPositSN(iposit_sn) %>
				       </b>
				      	</td>
				      	<td>
							<%= GetDayOrHourWithPositSN(iposit_sn, expiredday) %> <%= GetDayOrHourNameWithPositSN(iposit_sn) %>
						</td>
				<%
				else
				i = GetPrevYearVacationDayByEmpno(empno, totalvacationday, usedvacationday, requestedday, expiredday)

				%>
				    <tr align="center" bgcolor="#FFFFFF">
				    	<td>�۳� �ް�</td>
				    	<td>
							<%= GetDayOrHourWithPositSN(iposit_sn, totalvacationday) %> <%= GetDayOrHourNameWithPositSN(iposit_sn) %>
						</td>
				      	<td>
							<%= GetDayOrHourWithPositSN(iposit_sn, usedvacationday) %> <%= GetDayOrHourNameWithPositSN(iposit_sn) %>
						</td>
				      	<td>
							<%= GetDayOrHourWithPositSN(iposit_sn, requestedday) %> <%= GetDayOrHourNameWithPositSN(iposit_sn) %>
						</td>
				      	<td>
							<b>
								<%= GetDayOrHourWithPositSN(iposit_sn, (totalvacationday - (usedvacationday + requestedday+expiredday))) %> <%= GetDayOrHourNameWithPositSN(iposit_sn) %>
				       </b>
				      	</td>
				      	<td>
							<%= GetDayOrHourWithPositSN(iposit_sn, expiredday) %> <%= GetDayOrHourNameWithPositSN(iposit_sn) %>
						</td>
				    </tr>
				<%

				i = GetCurrYearVacationDayByEmpno(empno, totalvacationday, usedvacationday, requestedday, expiredday)

				%>
				    <tr align="center" bgcolor="#FFFFFF">
				    	<td>�ݳ� �ް�</td>
				    	<td>
							<%= GetDayOrHourWithPositSN(iposit_sn, totalvacationday) %> <%= GetDayOrHourNameWithPositSN(iposit_sn) %>
						</td>
				      	<td>
							<%= GetDayOrHourWithPositSN(iposit_sn, usedvacationday) %> <%= GetDayOrHourNameWithPositSN(iposit_sn) %>
						</td>
				      	<td>
							<%= GetDayOrHourWithPositSN(iposit_sn, requestedday) %> <%= GetDayOrHourNameWithPositSN(iposit_sn) %>
						</td>
				      	<td>
							<b>
								<%= GetDayOrHourWithPositSN(iposit_sn, (totalvacationday - (usedvacationday + requestedday+expiredday))) %> <%= GetDayOrHourNameWithPositSN(iposit_sn) %>
				      </b>
				      	</td>
				      	<td>
							<%= GetDayOrHourWithPositSN(iposit_sn, expiredday) %> <%= GetDayOrHourNameWithPositSN(iposit_sn) %>
						</td>
				    </tr>
				 <%end if%>
				    <tr height="25" bgcolor="FFFFFF">
						<td colspan="15">
							<br>
							<input type="button" class="button" value="�ް���û �� ��������" onclick="OpenVacationList()">
							<% if ((session("ssAdminPOsn") = "1") or (session("ssAdminPOsn") = "2") or (session("ssAdminPOsn") = "3") or (session("ssAdminPOsn") = "4") or session("ssAdminPsn")=7) then %>
							<input type="button" class="button" value="��Ʈ�� �ް���û ��������" onclick="OpenVacationListAdmin()" <% if Not C_IS_SCM_LOGIN then %>disabled<% end if %>>
							<% end if %>
							<br><br>

							* �߻������� �ų� �������� ��ȿ�ϸ�, ������� ���� ������ �������� �̿����� �ʽ��ϴ�.
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<!-- ��Ÿ�� -->
		<%
		if iposit_sn <= 11   then
		dim clsVM
		dim totvm, usevm, reqvm,payvm
		set clsVM = new CMyVitamin
		clsVM.FRectempno = empno
		clsVM.fnGetMyVitamin
		totvm = clsVM.Ftotvm
		usevm = clsVM.Fusevm
		set clsVM = nothing
		%>
		<tr>
			<td style="padding-top:20px;">
				<table width="100%"  border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
						<tr height="25" bgcolor="FFFFFF">
							<td  colspan="5">
								<font color="red" >&nbsp;<strong>��Ÿ�� ��Ȳ</strong></font>
							</td>
						</tr>
						<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
					    	<td>�⵵</td>
					    	<td>�ο��ݾ�</td>
					      <td>���ݾ�</td>
					      <td>�ܾ�</td>
					  </tr>
					  <tr align="center" bgcolor="FFFFFF">
					  	<td>�ݳ�</td>
					  	<td><%=formatnumber(totvm,0)%></td>
					  	<td><%=formatnumber(usevm,0)%></td>
					  	<td><%=formatnumber(totvm-usevm,0)%></td>
					  </tr>
					 <tr height="25" bgcolor="FFFFFF">
						<td colspan="5">
							<br>
							<input type="button" class="button" value="��Ÿ�� ��û�ϱ�" onclick="OpenVitamin()">
							<input type="button" class="button" value="��Ÿ�� ��û����" onclick="listVitamin()">

						</td>
					</tr>
				</table>
			</td>
		</tr>
	<%end if%>
	<!--// ��Ÿ�� -->
	<!-- ����Ʈ -->
		<%

		dim clsap
		dim totap, useap, reqap,payap,pkind,psdate,pedate, pstate
		dim arrap , intap
		set clsap = new CMyAgit
		clsap.FRectEmpno = empno
		clsap.FRectChkStart = year(date())
		arrap = clsap.fnGetMyInfoAgit

		set clsap = nothing


		%>
		<tr>
			<td style="padding-top:20px;">
				<table width="100%"  border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
						<tr height="25" bgcolor="FFFFFF">
							<td  colspan="5">
								<font color="red" >&nbsp;<strong>����Ʈ ����</strong></font>
							</td>
						</tr>
						<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
					    	<td>�⵵</td>
					    	<td>������Ʈ</td>
					      <td>�������Ʈ</td>
					      <td>�ܿ�����Ʈ</td>
					      <td>����</td>
					  </tr>
					  <%if isArray(arrap) then
					  			for intap = 0 To Ubound(arrap,2)
					  	 %>
					  <tr align="center" bgcolor="FFFFFF">
					  	<td><%=arrap(1,intap)%></td>
					  	<td><%=formatnumber(arrap(2,intap),0)%></td>
					  	<td><%=formatnumber(arrap(3,intap),0)%></td>
					  	<td><%=formatnumber(arrap(2,intap)-arrap(3,intap),0)%></td>
					  	<td>
					  		<%
					  		if (arrap(6,intap) ="0" or isNull(arrap(6,intap))  ) and (arrap(2,intap)-arrap(3,intap))>0 then
									pstate = "��밡��"
								else
									pstate = "���Ұ�"
								end if
					  		%>
					  		<%=pstate%>

					  	</td>
					  </tr>
					<%next %>
					 <tr height="25" bgcolor="FFFFFF">
						<td colspan="5">
							<br>
							<input type="button" class="button" value="����Ʈ ��û�ϱ�" onclick="OpenAgit()">
							<input type="button" class="button" value="����Ʈ ��û����" onclick="listAgit()">

						</td>
					</tr>
					<%
					else%>
					<tr>
						<td colspan="5" align="center">��û������ ����Ʈ�� �����ϴ�.</td>
					</tr>
					<%end if%>
				</table>
			</td>
		</tr>

		<!--2017.03.06 ������ �ּ�ó��-->
		<tr>
			<td  style="padding-top:20px;">
				<table width="100%"  border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
					<tr height="25" bgcolor="FFFFFF">
						<td  colspan="3">
							<font color="red" >&nbsp;<strong>�μ�����������</strong></font>
						</td>
					</tr>

<% if IS_SHOW_SCM_INFO then %>
					<tr height="25" bgcolor="FFFFFF">
						<td colspan="3">
							<form name="frmBiz" method="post" action="modify_myinfo_process.asp" style="margin:0px;">
								<input type="hidden" name="mode" value="biz">
							ERP�μ�����:
							<select name="selBiz">
								<option value="">--����--</option>
								<%IF isArray(arrBiz) THEN
										For intLoop =0 To UBound(arrBiz,2)
									%>
									<option value="<%=arrBiz(0,intLoop)%>" <%IF Cstr(sbizsection_Cd) = Cstr(arrBiz(0,intLoop)) THEN%>selected<%END IF%>><%=arrBiz(1,intLoop)%></option>
								<%
									Next
								END IF%>
							</select>
							<input type="button" value="���" class="button" onClick="document.frmBiz.submit();">
						</form>
						</td>
					</tr>
				  	<tr bgcolor="<%= adminColor("tabletop") %>"align="center">
						<td valign="top"  >
							 �μ� �� ��¥
						</td>
					   <td width="30%"><%=left(date(),7)%></td>
					   <td width="30%"><%=left(dateadd("m",-1,date()),7)%></td>
					</tr>
	<%IF isArray(arrList) THEN
		For intLoop = 0 To UBound(arrList,2)
			IF  arrList(2,intLoop) = ""   THEN
	%>
					<Tr bgcolor="#FFFFFF">
						<td><%=arrList(1,intLoop)%></td>
						<td></td>
						<td></td>
					</tr>
			<%	ELSE%>
					<Tr bgcolor="#FFFFFF">
						<td> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							 ��  <%IF not isNull(arrList(3,intLoop)) or arrList(3,intLoop)>0  or not isNull(arrList(4,intLoop)) or arrList(4,intLoop)> 0 THEN %><font color="blue"><%END IF%><%=arrList(1,intLoop)%></td>
						<td align="center"><%IF isNull(arrList(3,intLoop)) or arrList(3,intLoop)= 0 THEN %>0<%ELSE%><font color="blue"><%=arrList(3,intLoop)%></font><%END IF%> %</td>
						<td align="center"><%IF isNull(arrList(4,intLoop)) or arrList(4,intLoop)= 0 THEN %>0<%ELSE%><font color="blue"><%=arrList(4,intLoop)%></font><%END IF%> %</td>
					</tr>
			<%	END IF %>
		<% Next %>
	<% END IF%>
					<tr bgcolor="FFFFFF" align="center">
						<td></td>
						<td><input type="button" class="button" value="<%=left(date(),7)%> ���/����" onClick="jsSetUserBiz('<%=left(date(),7)%>');"></td>
						<td><%IF day(date())<=10 THEN%><input type="button" class="button" value="<%=left(dateadd("m",-1,date()),7)%> ���/����" onClick="jsSetUserBiz('<%=left(dateadd("m",-1,date()),7)%>');"><%END IF%></td>
					</tr>
<% end if %>
				</table>
			</td>
		</tr>
	</table>
</td>
</tr>
</table>
<%
	Dim vUserImage
	If sUserImage <> "" Then
		vUserImage = sUserImage
	Else
		vUserImage = "https://fiximage.10x10.co.kr/web2010/mytenbyten/grade_left_7.gif"
	End If
%>
<div id="drag" style="position:absolute; top:68px; left:343px; width:110px; height:132px; background-color:#FFF;">
<table border="1" cellpadding="0" cellspacing="0" height="132">
<tr style="cursor:pointer" onClick="window.open('https://www.10x10.co.kr/common/showimage.asp?img=<%=vUserImage%>', 'imageView', 'width=10,height=10,status=no,resizable=yes,scrollbars=yes');">
	<td><img src="<%=vUserImage%>" width="110" alt="�����̹�������"></td>
</tr>
<tr onmouseover="style.cursor='move'" onmousedown="start_drag('drag');">
	<td align="center" bgcolor="FFFFFF" valign="bottom"><font size="2">[�̵��ϱ�]</font></td>
</tr>
</table>
</div>

<script type="text/javascript">

var mouseDown;
var startDrag= false;

function move(){
	if(startDrag){
		mouseDown.style.left = x + event.clientX - pre_x;
		mouseDown.style.top  = y + event.clientY - pre_y;
		return false;
	}//if
}//drag_move

function start_drag(drag){
	mouseDown = document.getElementById(drag);
	//x,y
	x = parseInt(mouseDown.style.left);
	y = parseInt(mouseDown.style.top);
	pre_x = event.clientX;
	pre_y = event.clientY;

	//drag flag
	startDrag = true;
	//move
	mouseDown.onmousemove = move;
	//stop
	mouseDown.onmouseup = stop;
}

function stop(){
	startDrag=false;
}// drag_release

</script>

<%
set oAddLevel = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
