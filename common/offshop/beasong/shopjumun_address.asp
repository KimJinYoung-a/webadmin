<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : 오프라인 배송
' Hieditor : 2011.02.23 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/checkPoslogin.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<% '<!-- #include virtual="/lib/checkAllowIPWithLog.asp" --> %>
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->

<%
dim i , orderno , itemgubunarr ,itemoptionarr, itemidarr, mode , shopidarr, reqhp,comment
dim buyname,buyphone, buyhp, buyemail, reqname, reqzipcode, reqzipaddr, reqaddress, reqphone
dim buyphone1, buyphone2, buyphone3 ,buyhp1 ,buyhp2 ,buyhp3 ,reqphone1 ,reqphone2 ,reqphone3
dim reqhp1, reqhp2 ,reqhp3 ,buyemail1 ,buyemail2 , reqaddress1 ,reqaddress2
dim masteridx ,oedit ,odlvTypearr
	orderno = requestcheckvar(request("orderno"),16)
	masteridx = requestcheckvar(request("masteridx"),10)
	itemgubunarr = request("itemgubunarr")
	itemidarr = request("itemidarr")
	itemoptionarr = request("itemoptionarr")
	mode = requestcheckvar(request("mode"),32)
	shopidarr = request("shopidarr")
	odlvTypearr = request("odlvTypearr")

set oedit = new cupchebeasong_list
	oedit.frectmasteridx = masteridx
	oedit.frectorderno = orderno

	if masteridx <> "" or orderno <> "" then
		oedit.fshopjumun_edit()

		if oedit.ftotalcount > 0 then
			buyname = oedit.FOneItem.fbuyname
			buyphone = oedit.FOneItem.fbuyphone
				if buyphone<>"" then
					if instr(buyphone,"-") = 0 then
						buyphone = left(buyphone,3)
						buyphone = mid(buyphone,4,len(buyphone)-3-4)
						buyphone = right(buyphone,4)
					else
						buyphone1 = split(buyphone,"-")(0)
						buyphone2 = split(buyphone,"-")(1)
						buyphone3 = split(buyphone,"-")(2)
					end if
				end if
			buyhp = oedit.FOneItem.fbuyhp
				if buyhp<>"" then
					buyhp1 = split(buyhp,"-")(0)
					buyhp2 = split(buyhp,"-")(1)
					buyhp3 = split(buyhp,"-")(2)
				end if
			buyemail = oedit.FOneItem.fbuyemail
				if buyemail<>"" then
					buyemail = split(buyemail,"@")
					buyemail1 = buyemail(0)
					buyemail2 = buyemail(1)
				end if
			reqname = oedit.FOneItem.freqname
			reqzipcode = oedit.FOneItem.freqzipcode
			reqzipaddr = oedit.FOneItem.freqzipaddr
			reqaddress = oedit.FOneItem.freqaddress
			reqphone = oedit.FOneItem.freqphone
				if reqphone<>"" then
					reqphone = split(reqphone,"-")
					reqphone1 = reqphone(0)
					reqphone2 = reqphone(1)
					reqphone3 = reqphone(2)
				end if
			reqhp = oedit.FOneItem.freqhp
				if reqhp<>"" then
					if instr(reqhp,"-") = 0 then
						reqhp1 = left(reqhp,3)
						reqhp2 = mid(reqhp,4,len(reqhp)-3-4)
						reqhp3 = right(reqhp,4)
						'response.write reqhp1 & "/" & reqhp2 & "/" & reqhp3
					else
						reqhp1 = split(reqhp,"-")(0)
						reqhp2 = split(reqhp,"-")(1)
						reqhp3 = split(reqhp,"-")(2)
					end if
				end if
			comment = oedit.FOneItem.fcomment
		end if
	end if
%>

<script language="javascript">

//주문자 이메일 선택
function NewEmailChecker_buy(){
	var frm = document.frminfo;

	if( frm.txEmail_buy.value == "etc")  {
		frm.buyemail2.style.display = '';
		frm.buyemail2.focus();
	}else{
		frm.buyemail2.value = frm.txEmail_buy.value;
	}

	return;
}

//주소 검색 타고 들어옴 opener
function CopyZip(frmname, post1, post2, addr, dong) {
    eval(frmname + ".reqzipcode").value = post1 + "-" + post2;

    eval(frmname + ".reqzipaddr").value = addr;
    eval(frmname + ".reqaddress").value = dong;
}

// 폼전송
function frmsubmit(){
	if(frminfo.reqname.value!=''){
		/*
		if (IsDouble(frminfo.reqname.value)){
			alert('수령인 성함을 정확히 입력해주세요');
			frminfo.reqname.focus();
			return;
		}
		*/
	}else{
		alert('수령인 성함을 입력 하세요');
		frminfo.reqname.focus();
		return;
	}

	if (frminfo.reqhp1.value==''){
		alert('수령인 휴대전화 번호를 입력 하세요');
		frminfo.reqhp1.focus();
		return;
	}
	if (frminfo.reqhp2.value==''){
		alert('수령인 휴대전화 번호를 입력 하세요');
		frminfo.reqhp2.focus();
		return;
	}
	if (frminfo.reqhp3.value==''){
		alert('수령인 휴대전화 번호를 입력 하세요');
		frminfo.reqhp3.focus();
		return;
	}

	if(frminfo.reqzipcode.value=="") {
		alert('우편번호를 입력하세요.');
		return;
	}

//	for(var i=0; i < frminfo.reqzipcode.value.length;i++){
//		if(i==3){
//			if(frminfo.reqzipcode.value.charAt(i)!='-'){
//				alert('우편번호 중간에 - 이없습니다.\n정확히 입력해 주세요');
//				return;
//			}
//		}else{
//			if(isNaN(parseInt(frminfo.reqzipcode.value.charAt(i)))){
//				alert('우편번호에 '+frminfo.reqzipcode.value.charAt(i)+' 값이 잘못 입력 되었습니다.\n정확히 입력해 주세요');
//				return;
//			}
//		}
//	}

	if (frminfo.reqzipaddr.value==''){
		alert('주소를 입력 하세요');
		frminfo.reqzipaddr.focus();
		return;
	}
	if (frminfo.reqaddress.value==''){
		alert('주소를 입력 하세요');
		frminfo.reqaddress.focus();
		return;
	}

	frminfo.action='/common/offshop/beasong/shopbeasong_process.asp';
	frminfo.submit();
}

</script>

<form name="frminfo" method="post">
<input type="hidden" name="itemgubunarr" value="<%= itemgubunarr %>">
<input type="hidden" name="itemidarr" value="<%= itemidarr %>">
<input type="hidden" name="itemoptionarr" value="<%= itemoptionarr %>">
<input type="hidden" name="shopidarr" value="<%= shopidarr %>">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="orderno" value="<%= orderno %>">
<input type="hidden" name="masteridx" value="<%= masteridx %>">
<input type="hidden" name="odlvTypearr" value="<%= odlvTypearr %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<!--<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan=4>
		주문자
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td>
		성함
	</td>
	<td>
		<input type="text" size=16 maxlength=16 name="buyname" value="<%=buyname%>">
	</td>
	<td>
		이메일
	</td>
	<td>
		<input type="text" name="buyemail1" value="<%=buyemail1%>" class="input_02" style="width:95px;height:20px;" maxlength="32">
		@ <input type="text" name="buyemail2" maxlength="80" style="width:95px;height:20px;" value="<%=buyemail2%>">
		<select name="txEmail_buy" onchange="NewEmailChecker_buy()" style="width:95px;height:20px;">
			<option value="etc">직접입력</option>
			<option value="hanmail.net" >hanmail.net</option>
			<option value="naver.com" >naver.com</option>
			<option value="hotmail.com" >hotmail.com</option>
			<option value="yahoo.co.kr" >yahoo.co.kr</option>
			<option value="hanmir.com" >hanmir.com</option>
			<option value="paran.com" >paran.com</option>
			<option value="lycos.co.kr" >lycos.co.kr</option>
			<option value="nate.com" >nate.com</option>
			<option value="dreamwiz.com" >dreamwiz.com</option>
			<option value="korea.com" >korea.com</option>
			<option value="empal.com" >empal.com</option>
			<option value="netian.com" >netian.com</option>
			<option value="freechal.com" >freechal.com</option>
			<option value="msn.com" >msn.com</option>
			<option value="gmail.com" >gmail.com</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td>
		전화번호
	</td>
	<td>
		<input type="text" size=4 maxlength=4 name="buyphone1" value="<%=buyphone1%>">
		- <input type="text" size=4 maxlength=4 name="buyphone2" value="<%=buyphone2%>">
		- <input type="text" size=4 maxlength=4 name="buyphone3" value="<%=buyphone3%>">
	</td>
	<td>
		휴대전화
	</td>
	<td>
		<input type="text" size=4 maxlength=4 name="buyhp1" value="<%=buyhp1%>">
		- <input type="text" size=4 maxlength=4 name="buyhp2" value="<%=buyhp2%>">
		- <input type="text" size=4 maxlength=4 name="buyhp3" value="<%=buyhp3%>">
	</td>
</tr>-->
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan=4>
		주소입력
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td>
		* 성함
	</td>
	<td>
		<input type="text" size=16 maxlength=16 name="reqname" value="<%=reqname%>">
	</td>
	<td>
		이메일
	</td>
	<td>
		<input type="text" name="buyemail1" value="<%=buyemail1%>" class="input_02" style="width:95px;height:20px;" maxlength="32">
		@ <input type="text" name="buyemail2" maxlength="80" style="width:95px;height:20px;" value="<%=buyemail2%>">
		<select name="txEmail_buy" onchange="NewEmailChecker_buy()" style="width:95px;height:20px;">
			<option value="etc">직접입력</option>
			<option value="hanmail.net" >hanmail.net</option>
			<option value="naver.com" >naver.com</option>
			<option value="hotmail.com" >hotmail.com</option>
			<option value="yahoo.co.kr" >yahoo.co.kr</option>
			<option value="hanmir.com" >hanmir.com</option>
			<option value="paran.com" >paran.com</option>
			<option value="lycos.co.kr" >lycos.co.kr</option>
			<option value="nate.com" >nate.com</option>
			<option value="dreamwiz.com" >dreamwiz.com</option>
			<option value="korea.com" >korea.com</option>
			<option value="empal.com" >empal.com</option>
			<option value="netian.com" >netian.com</option>
			<option value="freechal.com" >freechal.com</option>
			<option value="msn.com" >msn.com</option>
			<option value="gmail.com" >gmail.com</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td>
		전화번호
	</td>
	<td>
		<input type="text" size=4 maxlength=4 name="reqphone1" value="<%=reqphone1%>">
		- <input type="text" size=4 maxlength=4 name="reqphone2" value="<%=reqphone2%>">
		- <input type="text" size=4 maxlength=4 name="reqphone3" value="<%=reqphone3%>">
	</td>
	<td>
		* 휴대전화
	</td>
	<td>
		<input type="text" size=4 maxlength=4 name="reqhp1" value="<%=reqhp1%>">
		- <input type="text" size=4 maxlength=4 name="reqhp2" value="<%=reqhp2%>">
		- <input type="text" size=4 maxlength=4 name="reqhp3" value="<%=reqhp3%>">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td>* 우편번호</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="reqzipcode" size="12" value="<%= reqzipcode %>">
		<input type="button" class="button" value="검색" onClick="FnFindZipNew('frminfo','A')">
		<input type="button" class="button" value="검색(구)" onClick="TnFindZipNew('frminfo','A')">
		<% '<input type="button" class="button" value="검색(구)" onClick="javascript:PopSearchZipcode('frminfo');"> %>
	</td>
	<td>* 주소</td>
	<td bgcolor="#FFFFFF">
		<input type="text" name="reqzipaddr" size="50" class="text_ro" value="<%= reqzipaddr %>">
		<br><input type="text" name="reqaddress" size="50" maxlength="128" class="text"  value="<%= reqaddress %>">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td>주문유의사항</td>
	<td bgcolor="#FFFFFF" colspan=3>
		<textarea rows=5 cols=100 name="comment"><%= comment %></textarea>
	</td>
</tr>
<tr align="center" bgcolor="ffffff">
	<td colspan=4>
		<input type="button" class="button" value="저장" onclick="frmsubmit();">
		<!--<input type="button" class="button" value="이전페이지로" onclick="location.href='/common/offshop/beasong/shopjumun_list.asp?orderno=<%'=orderno%>'">-->
	</td>
</tr>
</table>
</form>

<%
set oedit = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->