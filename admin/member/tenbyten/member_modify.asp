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
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member/10x10staffcls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVacationCls.asp" -->
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<%
dim cMember
dim userid
Dim sempno,sfrontid, susername,  dbirthday , blnissolar ,szipcode ,blnsexflag ,szipaddr,suseraddr , suserphone , susercell  ,susermail  ,smsnmail , sinterphoneno, sextension   , sdirect070   , sjobdetail   ,sjuminno
Dim blnstatediv   ,djoinday     ,dretireday   ,suserimage   ,ipart_sn     ,iposit_sn    ,ijob_sn      ,ilevel_sn    ,iuserdiv     
Dim spart_name, sposit_name, sjob_name, slevel_name
Dim smywork, isIdentify,sbizsection_Cd
Dim arrList, intLoop
userid = session("ssBctId") 

Set cMember = new CTenByTenMember
	cMember.Fuserid = userid
	cMember.fnGetScmMyInfo

	sempno   		= cMember.Fempno           	
	sfrontid       	= cMember.Ffrontid     
	susername      	= cMember.Fusername    
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
	isIdentify		= cMember.FisIdentify
	sbizsection_Cd	= cMember.Fbizsection_cd 
	IF isNull(sbizsection_Cd) THEN sbizsection_Cd = ""
	cMember.Fempno = sempno
	arrList = cMember.fnGetUserBizSection
Set cMember = nothing
'==============================================================================
dim birthday_yyyy, birthday_mm, birthday_dd

if (Not IsNull(dbirthday) and (dbirthday <> "")) then
	birthday_yyyy = Year(dbirthday)
	birthday_mm = Month(dbirthday)
	birthday_dd = Day(dbirthday)
end if
 

'==============================================================================
dim joinday_yyyy, joinday_mm, joinday_dd

if ((djoinday) and (djoinday <> "")) then
	joinday_yyyy = Year(djoinday)
	joinday_mm = Month(djoinday)
	joinday_dd = Day(djoinday)
end if
'==============================================================================
dim i

dim totalvacationday
dim usedvacationday
dim requestedday
dim expiredday
dim blnView,blnSale,clsBS,arrBiz
 	blnView = "Y"
	blnSale = "N" 
 Set clsBS = new CBizSection  
	clsBS.FUSE_YN = "Y"
	clsBS.FOnlySub = "Y"
	clsBS.FView		= blnView
	clsBS.FSale		= blnSale
	arrBiz = clsBS.fnGetBizSectionList   
Set clsBS = nothing	  
%>

<script language="javascript">

//document.domain = "10x10.co.kr";  //searchzip �ȵ�;

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
	var win = window.open("pop_tenbyten_vacation_list_admin.asp","OpenVacationListAdmin","width=900,height=500,scrollbars=yes");
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
	var winBiz = window.open("pop_member_bizsection_reg.asp?sEN=<%=sempno%>&sD="+sDate,"popBiz","width=630 height=800 scrollbars=yes");
	winBiz.focus();
}
</script>


<!--�⺻�������� ����-->
<table border="0" width="100%" cellpadding="10" cellspacing="0" class="a">
	<tr>
		<td>
		<table border=0 width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm_base" method="post" action="domodifymemberinfo.asp">
			<input type="hidden" name="mode" value="base">
			<input type="hidden" name="isadmin" value="N">
			<input type="hidden" name="userid" value="<%= userid %>">
			<input type="hidden" name="userimage" value="<%= sUserImage%>">
			<tr height="25" bgcolor="FFFFFF">
				<td colspan="2">
					<font color="red"><strong>[�⺻����]</strong></font>
				</td>
			</tr>
			<tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">���</td>
		    	<td bgcolor="#FFFFFF">
		    		<%=sempno %>
		    	</td>
		    </tr>
			<tr align="left" height="25">
		    	<td width="120" bgcolor="<%= adminColor("tabletop") %>">�̸�</td>
		    	<td bgcolor="#FFFFFF">
		    		<%=susername %>
		    	</td>
		    </tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">���� ���̵�</td>
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
		    		&nbsp;&nbsp;
		    		����: <input type="text" name="extension" size="5" class="text" value="<%= sextension %>">
		    	</td>
		    </tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">070 �����ȣ</td>
		    	<td bgcolor="#FFFFFF">
		    		<input type="text" name="direct070" id="" size="16" class="text" value="<%= sdirect070 %>">
		    	</td>
		    </tr>
		    <input type="hidden" name="birthday" value="">
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">�������</td>
		   	<td bgcolor="#FFFFFF">
		    		<select name=birthday_yyyy>
		<% for i = 1960 to 1995 %>
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
		    	<td bgcolor="<%= adminColor("tabletop") %>">������</td>
		    	<td bgcolor="#FFFFFF">
		    		<input name="mywork" type="text" size="45" class="text" maxlength="80" value="<%= smywork %>">
		    	</td>
		    </tr>
		    </form>
		    <tr align="left" height="50">
		    	<td colspan="2" bgcolor="#FFFFFF" align=center>
					<input type="button" class="button" value="�⺻���� ����" onclick="javascript:SaveBaseInfo()">
					&nbsp;&nbsp;&nbsp;
				<input type="button" class="button" value="����<% If sUserImage = "" Then %>���<% Else %>����<% End If %>" onclick="javascript:window.open('popAddImage.asp?sF=<%=session("ssAdminPsn")%>','myimageupload','width=380,height=150');">
		    	</td>
		    </tr>
			<tr>
				<td valign="bottom" colspan=2 bgcolor="FFFFFF">
					<font color="red"><strong>[����ó ����]</strong></font>
				</td>
			</tr>
			<form name="frm_addr" method="post" action="domodifymemberinfo.asp">
			<input type="hidden" name="mode" value="addr">
			<input type="hidden" name="isadmin" value="N">
			<input type="hidden" name="userid" value="<%= userid %>">
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">�޴�����ȣ</td>
		    	<td bgcolor="#FFFFFF">
		    		<input type="text" name="usercell" size="16" class="text_ro" value="<%= susercell %>" readonly onFocusOut="phone_format(frm_addr.usercell)">
		    		<input type="button" class="button_s" value="�޴�����ȣ ����" onClick="javascript:PopChgHPNum();" style="width:100px;">
		    		<%=chkIIF(isIdentify="Y","<font color=darkred>����Ȯ�� ��</font>","")%>
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
					<input type="button" class="button_s" value="�ּ��Է�" onClick="javascript:PopSearchZipcode('frm_addr');">
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
		    <tr align="left" height="50">
		    	<td colspan="2" bgcolor="#FFFFFF" align=center>
					<input type="button" class="button" value="����ó ����" onclick="javascript:SaveAddressInfo()">
		    	</td>
		    </tr>
			<tr>
				<td valign="bottom" colspan=2 bgcolor="FFFFFF">
					<font color="red"><strong>[��������]</strong></font>
				</td>
			</tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">�μ�-��Ʈ</td>
		    	<td bgcolor="#FFFFFF">
		    		<%= spart_name %>
		    	</td>
		    </tr>  
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">���α���</td>
		    	<td bgcolor="#FFFFFF">
		    		<%= slevel_name %>
		    	</td>
		    </tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">����</td>
		    	<td bgcolor="#FFFFFF">
		    		<%= sposit_name %>
		    	</td>
		    </tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">��å</td>
		    	<td bgcolor="#FFFFFF">
		    		<%= sjob_name %>
		    	</td>
		    </tr>
		    <tr align="left" height="25">
		    	<td bgcolor="<%= adminColor("tabletop") %>">������(ī�װ�)</td>
		    	<td bgcolor="#FFFFFF">
		    		<%= sjobdetail %>
		    	</td>
		    </tr>
		
		<!--��й�ȣ���� ����-->
			<form name="frm_mypass" method="post" action="domodifymemberinfo.asp">
			<input type="hidden" name="mode" value="mypass">
			<input type="hidden" name="isadmin" value="N">
			<input type="hidden" name="userid" value="<%= userid %>">
			<tr>
				<td valign="bottom" colspan=2 bgcolor="FFFFFF">
					<font color="red"><strong>[��й�ȣ]</strong></font>
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
		    	</td>
		    </tr>
		    </form>
		    <tr align="center" height="50">
		    	<td colspan="2" bgcolor="#FFFFFF">
		    		<input type="button" class="button_s" value="��й�ȣ ����" onclick="javascript:SavePassInfo()">
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
				
				i = GetPrevYearVacationDay(userid, totalvacationday, usedvacationday, requestedday, expiredday)
				
				%>
				    <tr align="center" bgcolor="#FFFFFF">
				    	<td>�۳� �ް�</td>
				    	<td><%= totalvacationday %></td>
				      	<td><%= usedvacationday %></td>
				      	<td><%= requestedday %></td>
				      	<td>
				      		<% if (expiredday = 0) then %>
				      		<b><%= (totalvacationday - (usedvacationday + requestedday)) %></b>
				      		<% else %>
				      		<b><%= (totalvacationday - expiredday) %></b>
				      		<% end if %>
				      	</td>
				      	<td><%= expiredday %></td>
				    </tr>
				<%
				
				i = GetCurrYearVacationDay(userid, totalvacationday, usedvacationday, requestedday, expiredday)
				
				%>
				    <tr align="center" bgcolor="#FFFFFF">
				    	<td>�ݳ� �ް�</td>
				    	<td><%= totalvacationday %></td>
				      	<td><%= usedvacationday %></td>
				      	<td><%= requestedday %></td>
				      	<td>
				      		<% if (expiredday = 0) then %>
				      		<b><%= (totalvacationday - (usedvacationday + requestedday)) %></b>
				      		<% else %>
				      		<b><%= (totalvacationday - expiredday) %></b>
				      		<% end if %>
				      	</td>
				      	<td><%= expiredday %></td>
				    </tr>
				    <tr height="25" bgcolor="FFFFFF">
						<td colspan="15">
							<br>
							<input type="button" class="button" value="�ް���û �� ��������" onclick="OpenVacationList()">
							<% if ((session("ssAdminPOsn") = "1") or (session("ssAdminPOsn") = "2") or (session("ssAdminPOsn") = "3") or (session("ssAdminPOsn") = "4") or session("ssAdminPsn")=7) then %>
							<input type="button" class="button" value="��Ʈ�� �ް���û ��������" onclick="OpenVacationListAdmin()">
							<% end if %>
							<br><br>
				
							* �����̿������� 3�������� ��ȿ�ϸ�, �ް���û�� �����̿��������� �����˴ϴ�.<br>
						</td>
					</tr>
				</table>
			</td>
		</tr> 
		<tr>
			<td  style="padding-top:20px;">
				<table width="100%"  border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
					<tr height="25" bgcolor="FFFFFF">
						<td  colspan="3">
							<font color="red" >&nbsp;<strong>�μ�����������</strong></font>
						</td>
					</tr> 
					<tr height="25" bgcolor="FFFFFF">
						<td colspan="3">
							<form name="frmBiz" method="post" action="domodifymemberinfo.asp" style="margin:0px;">
								<input type="hidden" name="mode" value="biz">
								<input type="hidden" name="empno" value="<%=sempno%>">
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
			<%	END IF
					Next
					END IF%>
						<tr bgcolor="FFFFFF" align="center">
							<td></td>  
							<td><input type="button" class="button" value="<%=left(date(),7)%> ���/����" onClick="jsSetUserBiz('<%=left(date(),7)%>');"></td>
							<td><%IF day(date())<=10 THEN%><input type="button" class="button" value="<%=left(dateadd("m",-1,date()),7)%> ���/����" onClick="jsSetUserBiz('<%=left(dateadd("m",-1,date()),7)%>');"><%END IF%></td>
						</tr>
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
		vUserImage = "http://fiximage.10x10.co.kr/web2010/mytenbyten/grade_left_7.gif"
	End If
%>
<div id="drag" style="position:absolute; top:68px; left:343px; width:110px; height:132px; background-color:#FFF;">
<table border="1" cellpadding="0" cellspacing="0" height="132">
<tr style="cursor:pointer" onClick="window.open('http://www.10x10.co.kr/common/showimage.asp?img=<%=vUserImage%>', 'imageView', 'width=10,height=10,status=no,resizable=yes,scrollbars=yes');">
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
 

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->