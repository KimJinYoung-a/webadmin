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
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenVacationCls.asp" -->
<!-- #include virtual="/lib/classes/offshop/offshopchargecls.asp"-->
<%

'==============================================================================
dim userid

userid = requestCheckvar(request("userid"),32)

dim oMember
Set oMember = new CTenByTenMember

oMember.FRectUserId = userid

oMember.GetMember



'==============================================================================
dim birthday_yyyy, birthday_mm, birthday_dd

if ((Not IsNull(oMember.FitemList(1).Fbirthday)) and (oMember.FitemList(1).Fbirthday <> "")) then
	birthday_yyyy = Year(oMember.FitemList(1).Fbirthday)
	birthday_mm = Month(oMember.FitemList(1).Fbirthday)
	birthday_dd = Day(oMember.FitemList(1).Fbirthday)
end if



'==============================================================================
dim joinday_yyyy, joinday_mm, joinday_dd

if ((Not IsNull(oMember.FitemList(1).Fjoinday)) and (oMember.FitemList(1).Fjoinday <> "")) then
	joinday_yyyy = Year(oMember.FitemList(1).Fjoinday)
	joinday_mm = Month(oMember.FitemList(1).Fjoinday)
	joinday_dd = Day(oMember.FitemList(1).Fjoinday)
end if



'==============================================================================
dim i

dim totalvacationday
dim usedvacationday
dim requestedday
dim expiredday

%>

<script language="javascript">

document.domain = "10x10.co.kr";

function SaveBaseInfo() {
	var frm = document.frm_base;

	// ========================================================================
	if (frm.username.value == ''){
		alert("�̸��� �Է��ϼ���");
		frm.username.focus();
		return;
	}
	
	// ========================================================================

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret) {
		frm.submit();
	}
}



function SaveAddressInfo() {
	var frm = document.frm_addr;

	// ========================================================================
	if (frm.usercell.value == ''){
		alert("�ڵ�����ȣ�� �Է��ϼ���");
		frm.usercell.focus();
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

	if ((frm.part_sn.value == 6) || (frm.part_sn.value == 13)) {
		// 6  : �������� - ��ȭ��
		// 13 : ����������
	} else {
		if (frm.bigo.value != 0) {
			alert("�ش�μ�(��ȭ��,����������) �� ��缥�� ������ �� �ֽ��ϴ�.");
			return;
		}
	}

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



function SubmitDelete() {
	var frm = document.frm_isusing;

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret) {
		frm.isusing.value = "N";
		frm.submit();
	}
}



function SubmitUndelete() {
	var frm = document.frm_isusing;

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret) {
		frm.isusing.value = "Y";
		frm.submit();
	}
}


function SaveUserImage()
{
	//alert(frm_base.userimage.value);
	var frm = document.frm_base;

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


</script>

<!--�⺻�������� ����-->
<table width="48%" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm_base" method="post" action="domodifymemberinfo.asp">
	<input type="hidden" name="mode" value="base">
	<input type="hidden" name="userid" value="<%= oMember.FitemList(1).Fuserid %>">
	<input type="hidden" name="userimage" value="<%= oMember.FItemList(1).FUserImage%>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="2">
			<font color="red"><strong>[�⺻����]</strong></font>
		</td>
	</tr>
	<tr align="left" height="25">
    	<td width="120" bgcolor="<%= adminColor("tabletop") %>">�̸�</td>
    	<td bgcolor="#FFFFFF">
    		<input name="username" type="text" size="16" class="text" value="<%= oMember.FitemList(1).Fusername %>">
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">���� ���̵�</td>
    	<td bgcolor="#FFFFFF">
    		<%= oMember.FitemList(1).Fuserid %>
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">�ٹ����� ���̵�</td>
    	<input name="frontid" type="hidden" value="<%= oMember.FitemList(1).Ffrontid %>">
    	<td bgcolor="#FFFFFF">
    		<%= oMember.FitemList(1).Ffrontid %>
    	</td>
    </tr>
	<tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">E-MAIL(�系����)</td>
    	<td bgcolor="#FFFFFF">
    		<input name="usermail" type="text" size="45" class="text" value="<%= oMember.FitemList(1).Fusermail %>">
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">ȸ����ȭ(����)</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="interphoneno" size="16" class="text" value="<%= oMember.FitemList(1).Finterphoneno %>">
    		&nbsp;&nbsp;
    		����: <input type="text" name="extension" size="5" class="text" value="<%= oMember.FitemList(1).Fextension %>">
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">070 �����ȣ</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="direct070" id="" size="16" class="text" value="<%= oMember.FitemList(1).Fdirect070 %>">
    	</td>
    </tr>
    <input type="hidden" name="birthday" value="">
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">�������</td>
    	<td bgcolor="#FFFFFF">
    		<%
    		if (Not IsNull(oMember.FitemList(1).Fbirthday)) then
    			response.write Left(oMember.FitemList(1).Fbirthday, 10)
    		end if
    		%>
			&nbsp; &nbsp; &nbsp; &nbsp;
			[
			<% if (oMember.FitemList(1).Fissolar = "Y") then response.write "���" end if %>
			<% if (oMember.FitemList(1).Fissolar = "N") then response.write "����" end if %>
			]
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">����</td>
    	<td bgcolor="#FFFFFF">
			<% if (oMember.FitemList(1).Fsexflag = "M") then response.write "����" end if %>
			<% if (oMember.FitemList(1).Fsexflag = "F") then response.write "����" end if %>
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">MSN�޽���</td>
    	<td bgcolor="#FFFFFF">
    		<%= oMember.FitemList(1).Fmsnmail %>
    	</td>
    </tr>
    </form>
    <tr align="left" height="50">
    	<td colspan="2" bgcolor="#FFFFFF" align=center>
			<input type="button" class="button" value="�⺻���� ����" onclick="javascript:SaveBaseInfo()">
			&nbsp;&nbsp;&nbsp;
		<input type="button" class="button" value="����<% If oMember.FItemList(1).FUserImage = "" Then %>���<% Else %>����<% End If %>" onclick="javascript:window.open('popAddImage.asp?sF=<%=oMember.FItemList(1).Fpart_sn%>','myimageupload','width=380,height=150');">
    	</td>
    </tr>
	<tr>
		<td valign="bottom" colspan=2 bgcolor="FFFFFF">
			<font color="red"><strong>[��󿬶��� ����]</strong></font>
		</td>
	</tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">�ڵ�����ȣ</td>
    	<td bgcolor="#FFFFFF">
    		<%= oMember.FitemList(1).Fusercell %>
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">����ȭ��ȣ</td>
    	<td bgcolor="#FFFFFF">
    		<%= oMember.FitemList(1).Fuserphone %>
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">�ּ�</td>
    	<td bgcolor="#FFFFFF">
    		[<%= oMember.FitemList(1).Fzipcode %>] <%= oMember.FitemList(1).Fzipaddr %>&nbsp;<%= oMember.FitemList(1).Fuseraddr %>
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
    		<%= oMember.FitemList(1).Fpart_name %>
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">������̵�</td>
    	<td bgcolor="#FFFFFF">
<%
if ((oMember.FitemList(1).Fpart_sn = "6") or (oMember.FitemList(1).Fpart_sn = "13")) then
	'6  : �������� - ��ȭ��
	'13 : ����������
%>
    		<%= oMember.FitemList(1).Fbigo %>
<% end if %>
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">���α���</td>
    	<td bgcolor="#FFFFFF">
    		<%= oMember.FitemList(1).Flevel_name %>
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">����</td>
    	<td bgcolor="#FFFFFF">
    		<%= oMember.FitemList(1).Fposit_name %>
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">��å</td>
    	<td bgcolor="#FFFFFF">
    		<%= oMember.FitemList(1).Fjob_name %>
    	</td>
    </tr>
    <tr align="left" height="25">
    	<td bgcolor="<%= adminColor("tabletop") %>">������(ī�װ�)</td>
    	<td bgcolor="#FFFFFF">
    		<%= oMember.FitemList(1).Fjobdetail %>
    	</td>
    </tr>






</table>

<table width="50%" align="right" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<font color="red"><strong>�߰�����</strong></font>
		</td>
	</tr>
	<form name="frm_moreinfo" method="post" action="domodifymemberinfo.asp">
	<input type="hidden" name="mode" value="moreinfo">
	<input type="hidden" name="userid" value="<%= oMember.FitemList(1).Fuserid %>">
	<tr height="25" align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="150">�Ի���</td>
    	<td width="100">�ټӿ���</td>
    	<td colspan=3></td>
      	<td></td>
    </tr>
    <input type="hidden" name="joinday" value="">
    <tr height="25" align="center" bgcolor="#FFFFFF">
    	<td>
    		<select name=joinday_yyyy>
<% for i = 2001 to Year(now())+1 %>
    			<option value="<%= i %>" <% if (joinday_yyyy = i) then %>selected<% end if %>><%= i %></option>
<% next %>
    		</select>
    		<select name=joinday_mm>
<% for i = 1 to 12 %>
    			<option value="<%= i %>" <% if (joinday_mm = i) then %>selected<% end if %>><%= i %></option>
<% next %>
    		</select>
    		<select name=joinday_dd>
<% for i = 1 to 31 %>
    			<option value="<%= i %>" <% if (joinday_dd = i) then %>selected<% end if %>><%= i %></option>
<% next %>
    		</select>
    	</td>
    	<td><%= oMember.FitemList(1).GetYearDiff %></td>
      	<td colspan=3></td>
      	<td></td>
    </tr>
    </form>
    <tr align="center" height="50">
    	<td colspan="6" bgcolor="#FFFFFF">
    		<input type="button" class="button_s" value="�߰����� ����" onclick="javascript:SaveMoreInfo()">
    	</td>
    </tr>
</table>
<br><br><br><br><br><br><br><br><br>
<table width="50%" align="right" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
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
			* �����̿������� 3�������� ��ȿ�ϸ�, �ް���û�� �����̿��������� �����˴ϴ�.<br>
		</td>
	</tr>
</table>


<%
	Dim vUserImage
	If oMember.FItemList(1).FUserImage <> "" Then
		vUserImage = oMember.FItemList(1).FUserImage
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
	<td align="center" valign="bottom"><font size="2">[�̵��ϱ�]</font></td>
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
set oMember = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->