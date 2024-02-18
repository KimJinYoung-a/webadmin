<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 정기구독 상품
' History : 2016.06.16 한용민 생성
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
		response.write "<script type='text/javascript'>alert('일련번호가 없습니다.');</script>"
		dbget.close() : response.end
	end if
else
	if orgitemid="" or orgitemoption="" then
		response.write "<script type='text/javascript'>alert('판매용상품코드나 판매용옵션코드가 없습니다.');</script>"
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
		alert("적요를 입력해주세요");
		frmstanding.jukyogubun.focus();
		return false;
	}
	if(!frmstanding.username.value){
		alert("이름을 입력해주세요");
		frmstanding.username.focus();
		return false;
	}
	if(!frmstanding.zipcode.value){
		alert("우편번호를 입력해주세요");
		frmstanding.zipcode.focus();
		return false;
	}
	if(!frmstanding.addr1.value){
		alert("주소1을 입력해주세요");
		frmstanding.addr1.focus();
		return false;
	}
	if(!frmstanding.addr2.value){
		alert("상세주소를 입력해주세요");
		frmstanding.addr2.focus();
		return false;
	}
	if(!frmstanding.userphone1.value || !frmstanding.userphone2.value || !frmstanding.userphone3.value){
		alert("전화번호를 입력해주세요");
		frm.userphone1.focus();
		return false;
	}
	if(!frmstanding.usercell1.value || !frmstanding.usercell2.value || !frmstanding.usercell3.value){
		alert("핸드폰 번호를 입력해주세요");
		frm.userphone1.focus();
		return false;
	}
	if(!frmstanding.isusing.value){
		alert("사용여부를 선택해주세요");
		frmstanding.isusing.focus();
		return false;
	}
	if (frmstanding.itemno.value!=""){
		if (!IsDouble(frmstanding.itemno.value)){
			alert('수량은 숫자만 입력 가능합니다.');
			frmstanding.itemno.focus();
			return;
		}
	}else{
		alert("수량을 입력하세요.");
		frmstanding.isusing.focus();
		return false;
	}

	if(confirm("저장 하시겠습니까?")) {
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
		정기구독 발송 정보 수정
	</td>
</tr>
<tr align="left">
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">idx :</td>
	<td bgcolor="#FFFFFF" width="40%">
		<% if uidx <> "" then %>
			<%= uidx %>
			<input type="hidden" name="uidx" value="<%= uidx %>">
		<% else %>
			신규
		<% end if %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">발송일 :</td>
	<td bgcolor="#FFFFFF" width="40%"><%= senddate %></td>
</tr>
<tr align="left">
	<td bgcolor="<%= adminColor("tabletop") %>">주문번호 :</td>
	<td bgcolor="#FFFFFF">
		<% if editmode="SUDONG" then %>
			<input type="text" name="orderserial" value="">
			<br>필요시에만 입력하세요.
		<% else %>
			<%= orderserial %>
			<input type="hidden" name="orderserial" value="<%= orderserial %>">
		<% end if %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">아이디 :</td>
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
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">최초등록 :</td>
	<td bgcolor="#FFFFFF" width="40%">
        <%= regadminid %>
        <br><%= regdate %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">최종수정 :</td>
	<td bgcolor="#FFFFFF" width="40%">
        <%= lastadminid %>
        <br><%= lastupdate %>
	</td>
</tr>
<tr align="left">
	<td bgcolor="<%= adminColor("tabletop") %>">상태 :</td>
	<td bgcolor="#FFFFFF">
		<font color="red"><%= getsendstatusname(sendstatus) %></font>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">적요 :</td>
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
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">우편번호 :</td>
	<td bgcolor="#FFFFFF" width="40%">
		<input type="text" name="zipcode" size="7" value="<%= zipcode %>" readOnly class="text_ro">
		<input type="button" class="button" value="검색" onClick="FnFindZipNew('frmstanding','E')">
		<input type="button" class="button" value="검색(구)" onClick="TnFindZipNew('frmstanding','E')">
		<% '<input type="button" onclick="TnFindZip('frmstanding');" value="검색(구)" class="button"> %>
		<% if reqzipcode_u<>"" then response.write "<br><font color='red'>회원정보 : " & reqzipcode_u & "</font>" %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">이름 :</td>
	<td bgcolor="#FFFFFF" width="40%">
		<input type="text" name="username" size="20" value="<%= username %>">
		<% if reqname_u<>"" then response.write "<br><font color='red'>회원정보 : " & reqname_u & "</font>" %>
	</td>
</tr>
<tr align="left">
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">주소1 :</td>
	<td bgcolor="#FFFFFF" width="40%">
		<input type="text" name="addr1" size="40" value="<%= reqzipaddr %>" readOnly class="text_ro">
		<% if reqzipaddr_u<>"" then response.write "<br><font color='red'>회원정보 : " & reqzipaddr_u & "</font>" %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">상세주소 :</td>
	<td bgcolor="#FFFFFF" width="40%">
		<input type="text" name="addr2" size="40" value="<%= useraddr %>">
		<% if reqaddress_u<>"" then response.write "<br><font color='red'>회원정보 : " & reqaddress_u & "</font>" %>
	</td>
</tr>
<tr align="left">
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">전화번호 :</td>
	<td bgcolor="#FFFFFF" width="40%">
        <input type="text" name="userphone1" size="3" value="<%= userphone1 %>" maxlength="3">
        -
        <input type="text" name="userphone2" size="4" value="<%= userphone2 %>" maxlength="4">
        -
        <input type="text" name="userphone3" size="4" value="<%= userphone3 %>" maxlength="4">
		<% if reqphone_u<>"" then response.write "<br><font color='red'>회원정보 : " & reqphone_u & "</font>" %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">핸드폰 :</td>
	<td bgcolor="#FFFFFF" width="40%">
        <input type="text" name="usercell1" size="3" value="<%= usercell1 %>" maxlength="3">
        -
        <input type="text" name="usercell2" size="4" value="<%= usercell2 %>" maxlength="4">
        -
        <input type="text" name="usercell3" size="4" value="<%= usercell3 %>" maxlength="4">
		<% if reqhp_u<>"" then response.write "<br><font color='red'>회원정보 : " & reqhp_u & "</font>" %>
	</td>
</tr> 
<tr align="left">
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">사용여부 :</td>
	<td bgcolor="#FFFFFF" width="40%">
        <% drawSelectBoxisusingYN "isusing", isusing, " onchange='frmsubmit("""");'" %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" width="10%">수량 :</td>
	<td bgcolor="#FFFFFF" width="40%">
		<input type="text" name="itemno" value="<%= itemno %>" size="7" maxlength="7" maxlength="3">
	</td>
</tr>
<tr align="center">
	<td bgcolor="#FFFFFF" colspan=4>
		<% if (sendstatus=0 or sendstatus=5) and editmode="EDIT" then %>
        	<input type="button" onClick="editstandinguser('editstandinguser');" value="수정" class="button">
        <% end if %>

		<% if (sendstatus=3 or sendstatus=7) and editmode="RE" then %>
        	<input type="button" onClick="editstandinguser('standinguser_re');" value="재발송" class="button">
        <% end if %>

        <% if editmode="SUDONG" then %>
        	<input type="button" onClick="editstandinguser('standinguser_sudong');" value="수동입력" class="button">
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