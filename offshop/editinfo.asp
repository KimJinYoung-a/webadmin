<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionOffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopchargecls.asp"-->
<%
dim opartner,i

set opartner = new CPartnerUser
opartner.FCurrpage = 1
opartner.FRectDesignerID = session("ssBctID")
opartner.FPageSize = 1
opartner.GetOnePartnerNUser 

Dim GroupIdExists 
if IsNULL(opartner.FOneItem.FGroupid) or (opartner.FOneItem.FGroupid="") then
    ''response.write "그룹코드가 지정되지 않았습니다. 수정불가."
    ''response.end
    GroupIdExists = FALSE
ELSE
    GroupIdExists = TRUE
end if

dim ogroup
set ogroup = new CPartnerGroup
ogroup.FRectGroupid = opartner.FOneItem.FGroupid
ogroup.GetOneGroupInfo

    
dim ochargeuser
set ochargeuser = new COffShopChargeUser
ochargeuser.FRectShopID = session("ssBctID")
ochargeuser.GetOffShopList


%>
<script language='javascript'>
function ModiInfo(frm){
		if (frm.groupid.value.length<1){
		alert('그룹코드가 설정되 있지 않습니다.- 관리자에게 문의하세요.');
		frm.groupid.focus();
		return;
	}

//	if (frm.company_name.value.length<1){
//		alert('사업자 등록상의 회사명을 입력하세요.');
//		frm.company_name.focus();
//		return;
//	}

//	if (frm.ceoname.value.length<1){
//		alert('사업자 등록상의 대표자명을 입력하세요.');
//		frm.ceoname.focus();
//		return;
//	}

//	if (frm.company_no.value.length<1){
//		alert('사업자 등록 번호를 입력하세요.');
//		frm.company_no.focus();
//		return;
//	}

	//if (frm.jungsan_gubun.value.length<1){
	//	alert('과세구분을 선택하세요.');
	//	frm.jungsan_gubun.focus();
	//	return;
	//}

	if (frm.company_zipcode.value.length<1){
		alert('우편번호를 선택하세요.');
		frm.company_zipcode.focus();
		return;
	}

	if (frm.company_address.value.length<1){
		alert('사업자 등록상의 주소1을 입력하세요.');
		frm.company_address.focus();
		return;
	}

	if (frm.company_address2.value.length<1){
		alert('사업자 등록상의 주소2를 입력하세요.');
		frm.company_address2.focus();
		return;
	}

	if (frm.company_uptae.value.length<1){
		alert('사업자 등록상의 업태를 입력하세요.');
		frm.company_uptae.focus();
		return;
	}

	if (frm.company_upjong.value.length<1){
		alert('사업자 등록상의 업종을 입력하세요.');
		frm.company_upjong.focus();
		return;
	}
    
    if (frm.shopname.value.length<1){
		alert('매장명을 입력하세요.');
		return;
	}
	
	if (frm.return_zipcode.value.length<1){
		alert('사무실주소(반품주소) 우편번호를 선택하세요.');
		frm.return_zipcode.focus();
		return;
	}

	if (frm.return_address.value.length<1){
		alert('사무실주소1(반품 주소1)을 입력하세요.');
		frm.return_address.focus();
		return;
	}

	if (frm.return_address2.value.length<1){
		alert('사무실주소2(반품 주소2)를 입력하세요.');
		frm.return_address2.focus();
		return;
	}

	if (frm.company_tel.value.length<1){
		alert('업체 전화번호를 입력하세요.');
		frm.company_tel.focus();
		return;
	}

	if (frm.manager_name.value.length<1){
		alert('담당자 성함을 입력하세요.');
		frm.manager_name.focus();
		return;
	}

	if (frm.manager_phone.value.length<1){
		alert('담당자 전화번호를 입력하세요.');
		frm.manager_phone.focus();
		return;
	}

	if (frm.manager_email.value.length<1){
		alert('담당자 이메일을 입력하세요.');
		frm.manager_email.focus();
		return;
	}

	if (frm.manager_hp.value.length<1){
		alert('담당자 핸드폰을 입력하세요.');
		frm.manager_hp.focus();
		return;
	}

	var ret = confirm('업체 정보를 저장 하시겠습니까?');

	if (ret){
		frm.submit();
	}

}


function CopyZip(flag,post1,post2,add,dong){
	if (flag=="s"){
		frmupche.company_zipcode.value= post1 + "-" + post2;
		frmupche.company_address.value= add;
		frmupche.company_address2.value= dong;
	}else if(flag=="m"){
		frmupche.return_zipcode.value= post1 + "-" + post2;
		frmupche.return_address.value= add;
		frmupche.return_address2.value= dong;
	}
}

function popZip(flag){
	var popwin = window.open("/lib/searchzip3.asp?target=" + flag,"searchzip3","width=460 height=240 scrollbars=yes resizable=yes");
	popwin.focus();
}

// 패스워드 복잡도 검사
function fnChkComplexPassword(pwd) {
    var aAlpha = /[a-z]|[A-Z]/;
    var aNumber = /[0-9]/;
    var aSpecial = /[!|@|#|$|%|^|&|*|(|)|-|_]/;
    var sRst = true;

    if(pwd.length < 8){
        sRst=false;
        return sRst;
    }

    var numAlpha = 0;
    var numNums = 0;
    var numSpecials = 0;
    for(var i=0; i<pwd.length; i++){
        if(aAlpha.test(pwd.substr(i,1)))
            numAlpha++;
        else if(aNumber.test(pwd.substr(i,1)))
            numNums++;
        else if(aSpecial.test(pwd.substr(i,1)))
            numSpecials++;
    }

    if((numAlpha>0&&numNums>0)||(numAlpha>0&&numSpecials>0)||(numNums>0&&numSpecials>0)) {
    	sRst=true;
    } else {
    	sRst=false;
    }
    return sRst;
}

function EditPass(frm){
	if (!frm.txoldpassword.value){
		alert('기존 비밀번호를 입력하세요.');
		frm.txoldpassword.focus();
		return;
	}
	
	if (!frm.txnewpassword1.value){
		alert('변경하실 1차 비밀번호를 입력하세요.');
		frm.txnewpassword1.focus();
		return;
	}

	if (frm.txnewpassword1.value.length < 8 || frm.txnewpassword1.value.length > 16){
			alert("비밀번호는 공백없이 8~16자입니다.");
			frm.txnewpassword1.focus();
			return ;
	}
	
	var uid = "<%=session("ssBctId")%>";
	
	if(frm.txnewpassword1.value==uid) {
		alert("아이디와 다른 비밀번호를 사용해주세요.");
		frm.txnewpassword1.focus();
		return  ;
	}
	
	if (!fnChkComplexPassword(frm.txnewpassword1.value)) {
			alert('패스워드는 영문/숫자/특수문자 중 두 가지 이상의 조합으로 입력해주세요.');
			frm.txnewpassword1.focus();
			return;
	}

 	if(!frm.txnewpassword2.value) {
			alert("비밀번호 확인을 입력해주세요.");
			frm.txnewpassword2.focus();
			return  ;
		}
		
		
	if (frm.txnewpassword1.value!=frm.txnewpassword2.value){
		alert('1차 비밀번호 확인이 일치하지 않습니다.');
		frm.txnewpassword2.focus();
		return;
	}
	
	if (!frm.txnewpasswordS1.value){
		alert('변경하실 2차 비밀번호를 입력하세요.');
		frm.txnewpasswordS1.focus();
		return;
	}

    if (frm.txnewpasswordS1.value.length < 8 || frm.txnewpasswordS1.value.length > 16){
			alert("비밀번호는 공백없이 8~16자입니다.");
			frm.txnewpasswordS1.focus();
			return ;
	}
	
	if (!fnChkComplexPassword(frm.txnewpasswordS1.value)) {
			alert('패스워드는 영문/숫자/특수문자 중 두 가지 이상의 조합으로 입력해주세요.');
			frm.txnewpasswordS1.focus();
			return;
		}
			
	if(!frm.txnewpasswordS2.value) {
			alert("2차 비밀번호 확인을 입력해주세요.");
			frm.txnewpasswordS2.focus();
			return  ;
		}
		 
	
	if (frm.txnewpasswordS1.value!=frm.txnewpasswordS2.value){
		alert('2차 비밀번호 확인이 일치하지 않습니다.');
		frm.txnewpasswordS2.focus();
		return;
	}

    if (frm.txnewpassword1.value==frm.txnewpasswordS1.value){
		alert('1차 비밀번호와  다른 비밀번호를 사용해주세요.');
		frm.txnewpasswordS1.focus();
		return;
	}
	
	var ret = confirm('비밀 번호를 수정 하시겠습니까?\r\nPOS 로그인 비밀번호와 SCM로그인 비밀번호가 동시에 변경됩니다.');
	if (ret){
		frm.submit();
	}

}

function SameReturnAddr(bool){
	if (bool){
		frmupche.return_zipcode.value = frmupche.company_zipcode.value;
		frmupche.return_address.value = frmupche.company_address.value;
		frmupche.return_address2.value = frmupche.company_address2.value;
	}else{
		frmupche.return_zipcode.value = "";
		frmupche.return_address.value = "";
		frmupche.return_address2.value = "";
	}
}
</script>

<table width="600" cellspacing="1" class="a" bgcolor=#3d3d3d>
<% if opartner.FresultCount >0 then %>


<table width="600" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmupche" method="post" action="shopinfoedit_process.asp">
	<input type="hidden" name="mode" value="groupedit">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<input type="button" class="icon" value="#">
			<font color="red"><b>사업자 정보</b></font>
		</td>
	</tr>
	<% IF (Not GroupIdExists) THEN %>
	<tr height="25" bgcolor="FFFFFF">
	    <td> 그룹 코드가 지정되지 않았습니다. 매장정보 수정 불가</td>
	</tr>
	<% ELSE %>
	<tr>
		<td width="120" bgcolor="<%= adminColor("tabletop") %>">업체코드</td>
		<td bgcolor="#FFFFFF" width="180">
			<input type="text" class="text_ro" name="groupid" value="<%= ogroup.FOneItem.FGroupId %>" size="7" maxlength="5" readonly>
		</td>
		<td width="120" bgcolor="<%= adminColor("tabletop") %>">업체명</td>
		<td bgcolor="#FFFFFF" width="180">
			<%= ogroup.FOneItem.FCompany_name %>
		</td>
	</tr>
	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**사업자등록정보**</td>
	</tr>

	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">회사명(상호)</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text_ro" name="company_name" value="<%= ogroup.FOneItem.FCompany_name %>" size="30" readonly>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>">대표자</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text_ro" name="ceoname" value="<%= ogroup.FOneItem.Fceoname %>" size="30" readonly>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">사업자번호</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text_ro" name="company_no" value="<%= ogroup.FOneItem.Fcompany_no %>" size="30" readonly>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>">과세구분</td>
		<td bgcolor="#FFFFFF">
			<input type="text" class="text_ro" name="" value="<%= ogroup.FOneItem.Fjungsan_gubun %>" size="30" readonly>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">사업장소재지</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<input type="text" class="text" name="company_zipcode" value="<%= ogroup.FOneItem.Fcompany_zipcode %>" size="7" maxlength="7">
			<input type="button" class="button" value="우편번호검색" onclick="javascript:popZip('s');"><br>
			<input type="text" class="text" name="company_address" value="<%= ogroup.FOneItem.Fcompany_address %>" size="26" maxlength="64">&nbsp;
			<input type="text" class="text" name="company_address2" value="<%= ogroup.FOneItem.Fcompany_address2 %>" size="38" maxlength="64">
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">업태</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_uptae" value="<%= ogroup.FOneItem.Fcompany_uptae %>" size="28" maxlength="32"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">업종</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_upjong" value="<%= ogroup.FOneItem.Fcompany_upjong %>" size="28" maxlength="32"></td>
	</tr>

	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**매장 기본정보** &nbsp;&nbsp;</td>
	</tr>

    <tr>
		<td bgcolor="<%= adminColor("tabletop") %>">매장명</td>
		<td bgcolor="#FFFFFF" colspan="3"><input type="text" class="text" name="shopname" value="<%= ochargeuser.FItemList(0).Fshopname %>" size="20" maxlength="64"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">대표전화</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_tel" value="<%= ogroup.FOneItem.Fcompany_tel %>" size="16" maxlength="16"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">팩스</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_fax" value="<%= ogroup.FOneItem.Fcompany_fax %>" size="16" maxlength="16"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">사무실주소<br>(배송지주소)</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<input type="text" class="text" name="return_zipcode" value="<%= ogroup.FOneItem.Freturn_zipcode %>" size="7" maxlength="7">
			<input type="button" class="button" value="우편번호검색" onclick="javascript:popZip('m');">
			<input type="checkbox" class="checkbox" name=samezip onclick="SameReturnAddr(this.checked)">상동<br>
			<input type="text" class="text" name="return_address" value="<%= ogroup.FOneItem.Freturn_address %>" size="26" maxlength="64">&nbsp;
			<input type="text" class="text" name="return_address2" value="<%= ogroup.FOneItem.Freturn_address2 %>" size="38" maxlength="64">
		</td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">거래택배사</td>
		<td colspan="3" bgcolor="#FFFFFF"><% drawSelectBoxDeliverCompany "defaultsongjangdiv" , opartner.FOneItem.Fdefaultsongjangdiv %>
		</td>
	</tr>
	<% if (FALSE) then %>
    <!--
	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**결제계좌정보** &nbsp;&nbsp;(결제계좌 정보 수정시 담당MD에게 연락하시기 바랍니다.)</td>
	</tr>

	<tr height="26">
		<td bgcolor="<%= adminColor("tabletop") %>">거래은행</td>
		<td colspan="3" bgcolor="#FFFFFF" >
		<input type="text" class="text_ro" value="<%= ogroup.FOneItem.Fjungsan_bank %>" size="30" readonly>
		<% if (IsNULL(ogroup.FOneItem.Fjungsan_acctno)) or (ogroup.FOneItem.Fjungsan_acctno="") then %>
		(결제 계좌를 등록하시려면  담당 MD에게 Fax로 통장 사본을 보내주시기 바랍니다.)
		<% end if %>
		</td>
	</tr>
	<tr height="26">
		<td bgcolor="<%= adminColor("tabletop") %>">계좌번호</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<input type="text" class="text_ro" value="<%= ogroup.FOneItem.Fjungsan_acctno %>" size="30" readonly>
		</td>
	</tr>
	<tr height="26">
		<td bgcolor="<%= adminColor("tabletop") %>">예금주명</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<input type="text" class="text_ro" value="<%= ogroup.FOneItem.Fjungsan_acctname %>" size="30" readonly>
		</td>
	</tr>
    -->
    <% end if %>
	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**담당자정보**</td>
	</tr>

	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">담당자명</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_name" value="<%= ogroup.FOneItem.Fmanager_name %>" size="30" maxlength="32"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">일반전화</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_phone" value="<%= ogroup.FOneItem.Fmanager_phone %>" size="30" maxlength="16"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_email" value="<%= ogroup.FOneItem.Fmanager_email %>" size="30" maxlength="64"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">핸드폰</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_hp" value="<%= ogroup.FOneItem.Fmanager_hp %>" size="30" maxlength="16"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">배송담당자명</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_name" value="<%= ogroup.FOneItem.Fdeliver_name %>" size="30" maxlength="32"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">일반전화</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_phone" value="<%= ogroup.FOneItem.Fdeliver_phone %>" size="30" maxlength="16"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_email" value="<%= ogroup.FOneItem.Fdeliver_email %>" size="30" maxlength="64"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">핸드폰</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_hp" value="<%= ogroup.FOneItem.Fdeliver_hp %>" size="30" maxlength="16"></td>
	</tr>

	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">정산담당자명</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_name" value="<%= ogroup.FOneItem.Fjungsan_name %>" size="30" maxlength="32"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">일반전화</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_phone" value="<%= ogroup.FOneItem.Fjungsan_phone %>" size="30" maxlength="16"></td>
	</tr>
	<tr>
		<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_email" value="<%= ogroup.FOneItem.Fjungsan_email %>" size="30" maxlength="64"></td>
		<td bgcolor="<%= adminColor("tabletop") %>">핸드폰</td>
		<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_hp" value="<%= ogroup.FOneItem.Fjungsan_hp %>" size="30" maxlength="16"></td>
	</tr>
	
	<tr align="center" height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<input type="button" class="button" value="업체정보 저장" onclick="ModiInfo(frmupche);">
		</td>
	</tr>
	<% END IF %>
	</form>
</table>

<br>

<p>

<table width="600" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmpass" method=post action="shopinfoedit_process.asp">
	<input type="hidden" name="mode" value="editpass">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="2">
			<input type="button" class="icon" value="#">
			<font color="red"><b>비밀번호 변경</b></font>
			&nbsp;
			(비밀번호를 변경하시려면 아래 란을 채워 주시기바랍니다.)
		</td>
	</tr>
	<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">기존비밀번호</td>
    	<td width="480" bgcolor="#FFFFFF">
    		1차 :<input type="password" class="text" name="txoldpassword" size="12" value="" maxlength="16">
    		2차 :<input type="password" class="text" name="txoldpasswordS" size="12" value="" maxlength="16">
    	</td>
    </tr>
	<tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">변경비밀번호 1차 </td>
    	<td width="480" bgcolor="#FFFFFF">
    		입력: <input type="password" class="text" name="txnewpassword1" size="12" value="" maxlength="16"><br>
    		확인: <input type="password" class="text" name="txnewpassword2" size="12" value="" maxlength="16">
    	</td>
    </tr> 
    <tr>
    	<td bgcolor="<%= adminColor("tabletop") %>">변경비밀번호 2차 </td>
    	<td width="480" bgcolor="#FFFFFF">
    		입력: <input type="password" class="text" name="txnewpasswordS1" size="12" value="" maxlength="16"><br>
    		확인: <input type="password" class="text" name="txnewpasswordS2" size="12" value="" maxlength="16">
    	</td>
    </tr>
	</form>
	
	<tr align="center" height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<input type="button" class="button" value="비밀번호변경" onclick="EditPass(frmpass);">
		</td>
	</tr>
</table>


<% end if %>
<%
set ochargeuser = Nothing
set ogroup = Nothing
set opartner = Nothing
%>
<!-- #include virtual="/offshop/lib/offshopbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->