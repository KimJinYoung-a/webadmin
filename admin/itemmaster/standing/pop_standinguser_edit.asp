<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���ⱸ�� ��ǰ
' History : 2016.06.16 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/standing/item_standing_cls.asp"-->
<%
dim uidx, i, menupos, orgitemid, orgitemoption, reserveidx, orderserial, userid, itemno, sendstatus, senddate, username
dim zipcode, reqzipaddr, useraddr, userphone, usercell, isusing, regdate, regadminid, lastupdate, lastadminid
dim userphone1, userphone2, userphone3, usercell1, usercell2, usercell3, tmpuserphone, tmpusercell, editmode
dim jukyogubun, ostanding, ouser, itemid, itemoption
dim reqname_u, reqzipcode_u, reqzipaddr_u, reqaddress_u, reqphone_u, reqhp_u
	uidx = getNumeric(requestcheckvar(request("uidx"),10))
	menupos = requestcheckvar(request("menupos"),10)
	editmode = requestcheckvar(request("editmode"),32)
	reserveidx = getNumeric(requestcheckvar(request("reserveidx"),10))
	orgitemid = getNumeric(requestcheckvar(request("itemid"),10))
	orgitemoption = requestcheckvar(request("itemoption"),10)

if editmode="RE" or editmode="EDIT" then
	if uidx="" or isnull(uidx) then
		response.write "<script type='text/javascript'>alert('�Ϸù�ȣ�� �����ϴ�.');</script>"
		dbget.close() : response.end
	end if
else
	if orgitemid="" or orgitemoption="" then
		response.write "<script type='text/javascript'>alert('�Ǹſ��ǰ�ڵ峪 �Ǹſ�ɼ��ڵ尡 �����ϴ�.');</script>"
		dbget.close() : response.end
	end if
end if

set ouser = new Citemstanding
	ouser.FRectuidx = uidx

if uidx<>"" then
	ouser.fitemstanding_user_one

	if ouser.ftotalcount > 0 then
		uidx = ouser.FOneItem.fuidx
		orgitemid = ouser.FOneItem.forgitemid
		orgitemoption = ouser.FOneItem.forgitemoption
		reserveidx = ouser.FOneItem.freserveidx
		jukyogubun = ouser.FOneItem.fjukyogubun
		orderserial = ouser.FOneItem.forderserial
		userid = ouser.FOneItem.fuserid
		itemno = ouser.FOneItem.fitemno
		sendstatus = ouser.FOneItem.fsendstatus
		senddate = ouser.FOneItem.fsenddate
		username = trim(ouser.FOneItem.fusername)
		zipcode = trim(ouser.FOneItem.fzipcode)
		reqzipaddr = trim(ouser.FOneItem.freqzipaddr)
		useraddr = trim(ouser.FOneItem.fuseraddr)
		userphone = trim(ouser.FOneItem.fuserphone)
		if userphone<>"" then
			tmpuserphone = split(trim(userphone),"-")
			if ubound(tmpuserphone) >= 2 then
				userphone1 = trim(tmpuserphone(0))
				userphone2 = trim(tmpuserphone(1))
				userphone3 = trim(tmpuserphone(2))
			end if
		end if
		usercell = trim(ouser.FOneItem.fusercell)
		if usercell<>"" then
			tmpusercell = split(trim(usercell),"-")
			if ubound(tmpusercell) >= 2 then
				usercell1 = trim(tmpusercell(0))
				usercell2 = trim(tmpusercell(1))
				usercell3 = trim(tmpusercell(2))
			end if
		end if
		isusing = ouser.FOneItem.fisusing
		regdate = ouser.FOneItem.fregdate
		regadminid = ouser.FOneItem.fregadminid
		lastupdate = ouser.FOneItem.flastupdate
		lastadminid = ouser.FOneItem.flastadminid

		reqname_u = trim(ouser.FOneItem.freqname_u)
		reqzipcode_u = trim(ouser.FOneItem.freqzipcode_u)
		reqzipaddr_u = trim(ouser.FOneItem.freqzipaddr_u)
		reqaddress_u = trim(ouser.FOneItem.freqaddress_u)
		reqphone_u = trim(ouser.FOneItem.freqphone_u)
		reqhp_u = trim(ouser.FOneItem.freqhp_u)
	end if

' else
' 	if editmode="SUDONG" then
' 		set ostanding = new Citemstanding
' 			ostanding.FRectItemID = itemid
' 			ostanding.FRectitemoption = itemoption

' 			if itemid<>"" and itemoption<>"" then
' 				ostanding.fitemstanding_one

' 				if ostanding.ftotalcount > 0 then
' 					senddate = ostanding.FOneItem.freserveDlvDate
' 					orgitemid = itemid
' 					orgitemoption = itemoption
' 				end if
' 			end if
' 	end if
end if

if isusing="" then isusing="Y"
%>
<script type="text/javascript">

function TnFindZip(frmname){
	window.open('<%= getSCMSSLURL %>/lib/newSearchzip.asp?target=' + frmname, 'findzipcdode', 'width=460,height=250,left=400,top=200,location=no,menubar=no,resizable=no,scrollbars=yes,status=no,toolbar=no');
}

function editstandinguser(editmode){
	if(!frmstanding.jukyogubun.value){
		alert("���並 �Է����ּ���");
		frmstanding.jukyogubun.focus();
		return false;
	}
	if(!frmstanding.username.value){
		alert("�̸��� �Է����ּ���");
		frmstanding.username.focus();
		return false;
	}
	if(!frmstanding.zipcode.value){
		alert("�����ȣ�� �Է����ּ���");
		frmstanding.zipcode.focus();
		return false;
	}
	if(!frmstanding.addr1.value){
		alert("�ּ�1�� �Է����ּ���");
		frmstanding.addr1.focus();
		return false;
	}
	if(!frmstanding.addr2.value){
		alert("���ּҸ� �Է����ּ���");
		frmstanding.addr2.focus();
		return false;
	}
	if(!frmstanding.userphone1.value || !frmstanding.userphone2.value || !frmstanding.userphone3.value){
		alert("��ȭ��ȣ�� �Է����ּ���");
		frm.userphone1.focus();
		return false;
	}
	if(!frmstanding.usercell1.value || !frmstanding.usercell2.value || !frmstanding.usercell3.value){
		alert("�ڵ��� ��ȣ�� �Է����ּ���");
		frm.userphone1.focus();
		return false;
	}
	if(!frmstanding.isusing.value){
		alert("��뿩�θ� �������ּ���");
		frmstanding.isusing.focus();
		return false;
	}
	if (frmstanding.itemno.value!=""){
		if (!IsDouble(frmstanding.itemno.value)){
			alert('������ ���ڸ� �Է� �����մϴ�.');
			frmstanding.itemno.focus();
			return;
		}
	}else{
		alert("������ �Է��ϼ���.");
		frmstanding.isusing.focus();
		return false;
	}

	if(confirm("���� �Ͻðڽ��ϱ�?")) {
		frmstanding.mode.value="standinguser_sudong";
		frmstanding.action="<%= getSCMSSLURL %>/admin/itemmaster/standing/standinguser_process.asp";
		frmstanding.submit();
	}
}

</script>

<form name="frmstanding" method="POST" action="" style="margin:0;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="itemgubun" value="10">
<input type="hidden" name="itemid" value="<%= orgitemid %>">
<input type="hidden" name="itemoption" value="<%= orgitemoption %>">
<input type="hidden" name="reserveidx" value="<%= reserveidx %>">

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left" bgcolor="#FFFFFF">
	<td height="30" colspan="4">
		���ⱸ�� �߼� ���� ����
	</td>
</tr>
<tr align="left">
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">idx :</td>
	<td bgcolor="#FFFFFF" width="40%">
		<% if uidx <> "" then %>
			<%= uidx %>
			<input type="hidden" name="uidx" value="<%= uidx %>">
		<% else %>
			�ű�
		<% end if %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">�߼��� :</td>
	<td bgcolor="#FFFFFF" width="40%"><%= senddate %></td>
</tr>
<tr align="left">
	<td bgcolor="<%= adminColor("tabletop") %>">�ֹ���ȣ :</td>
	<td bgcolor="#FFFFFF">
		<% if editmode="SUDONG" then %>
			<input type="text" name="orderserial" value="">
			<br>�ʿ�ÿ��� �Է��ϼ���.
		<% else %>
			<%= orderserial %>
			<input type="hidden" name="orderserial" value="<%= orderserial %>">
		<% end if %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">���̵� :</td>
	<td bgcolor="#FFFFFF">
		<% if editmode="SUDONG" then %>
			<input type="text" name="userid" value="">
		<% else %>
			<%= userid %>
			<input type="hidden" name="userid" value="<%= userid %>">
		<% end if %>
	</td>
</tr>
<tr align="left">
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">���ʵ�� :</td>
	<td bgcolor="#FFFFFF" width="40%">
        <%= regadminid %>
        <br><%= regdate %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">�������� :</td>
	<td bgcolor="#FFFFFF" width="40%">
        <%= lastadminid %>
        <br><%= lastupdate %>
	</td>
</tr>
<tr align="left">
	<td bgcolor="<%= adminColor("tabletop") %>">���� :</td>
	<td bgcolor="#FFFFFF">
		<font color="red"><%= getsendstatusname(sendstatus) %></font>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">���� :</td>
	<td bgcolor="#FFFFFF">
		<% if jukyogubun <> "" then %>
			<%= getjukyoname(jukyogubun) %>
			<input type="hidden" name="jukyogubun" size="16" value="<%= jukyogubun %>">
		<% else %>
			<% drawSelectBoxjukyo "jukyogubun", "EVENT", "" %>
		<% end if %>
	</td>
</tr>
</table>

<br>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="left">
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">�����ȣ :</td>
	<td bgcolor="#FFFFFF" width="40%">
		<input type="text" name="zipcode" size="7" value="<%= zipcode %>" readOnly class="text_ro">
		<input type="button" class="button" value="�˻�" onClick="FnFindZipNew('frmstanding','E')">
		<input type="button" class="button" value="�˻�(��)" onClick="TnFindZipNew('frmstanding','E')">
		<% '<input type="button" onclick="TnFindZip('frmstanding');" value="�˻�(��)" class="button"> %>
		<% if reqzipcode_u<>"" then response.write "<br><font color='red'>ȸ������ : " & reqzipcode_u & "</font>" %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">�̸� :</td>
	<td bgcolor="#FFFFFF" width="40%">
		<input type="text" name="username" size="20" value="<%= username %>">
		<% if reqname_u<>"" then response.write "<br><font color='red'>ȸ������ : " & reqname_u & "</font>" %>
	</td>
</tr>
<tr align="left">
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">�ּ�1 :</td>
	<td bgcolor="#FFFFFF" width="40%">
		<input type="text" name="addr1" size="40" value="<%= reqzipaddr %>" readOnly class="text_ro">
		<% if reqzipaddr_u<>"" then response.write "<br><font color='red'>ȸ������ : " & reqzipaddr_u & "</font>" %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">���ּ� :</td>
	<td bgcolor="#FFFFFF" width="40%">
		<input type="text" name="addr2" size="40" value="<%= useraddr %>">
		<% if reqaddress_u<>"" then response.write "<br><font color='red'>ȸ������ : " & reqaddress_u & "</font>" %>
	</td>
</tr>
<tr align="left">
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">��ȭ��ȣ :</td>
	<td bgcolor="#FFFFFF" width="40%">
        <input type="text" name="userphone1" size="3" value="<%= userphone1 %>" maxlength="3">
        -
        <input type="text" name="userphone2" size="4" value="<%= userphone2 %>" maxlength="4">
        -
        <input type="text" name="userphone3" size="4" value="<%= userphone3 %>" maxlength="4">
		<% if reqphone_u<>"" then response.write "<br><font color='red'>ȸ������ : " & reqphone_u & "</font>" %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">�ڵ��� :</td>
	<td bgcolor="#FFFFFF" width="40%">
        <input type="text" name="usercell1" size="3" value="<%= usercell1 %>" maxlength="3">
        -
        <input type="text" name="usercell2" size="4" value="<%= usercell2 %>" maxlength="4">
        -
        <input type="text" name="usercell3" size="4" value="<%= usercell3 %>" maxlength="4">
		<% if reqhp_u<>"" then response.write "<br><font color='red'>ȸ������ : " & reqhp_u & "</font>" %>
	</td>
</tr> 
<tr align="left">
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">��뿩�� :</td>
	<td bgcolor="#FFFFFF" width="40%">
        <% drawSelectBoxisusingYN "isusing", isusing, " onchange='frmsubmit("""");'" %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">���� :</td>
	<td bgcolor="#FFFFFF" width="40%">
		<input type="text" name="itemno" value="<%= itemno %>" size="7" maxlength="7" maxlength="3">
	</td>
</tr>
<tr align="center">
	<td bgcolor="#FFFFFF" colspan=4>
		<% if (sendstatus=0 or sendstatus=5) and editmode="EDIT" then %>
        	<input type="button" onClick="editstandinguser('editstandinguser');" value="����" class="button">
        <% end if %>

		<% if (sendstatus=3 or sendstatus=7) and editmode="RE" then %>
        	<input type="button" onClick="editstandinguser('standinguser_re');" value="��߼�" class="button">
        <% end if %>

        <% if editmode="SUDONG" then %>
        	<input type="button" onClick="editstandinguser('standinguser_sudong');" value="�����Է�" class="button">
        <% end if %>
	</td>
</tr>
</table>

</form>

<%
set ouser = nothing
set ostanding = nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->