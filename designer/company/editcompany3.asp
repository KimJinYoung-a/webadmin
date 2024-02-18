<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 업체정보
' History : 2009.04.17 최초생성자 모름
'			2016.06.30 한용민 수정
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp"-->
<%
dim i,page

dim opartner
set opartner = new CPartnerUser
	opartner.FCurrpage = 1
	opartner.FRectDesignerID = session("ssBctId")
	opartner.FPageSize = 1
	opartner.GetOnePartnerNUser

dim ogroup
set ogroup = new CPartnerGroup
	ogroup.FRectGroupid = opartner.FOneItem.FGroupid
	ogroup.GetOneGroupInfo

dim ooffontract
set ooffontract = new COffContractInfo
	ooffontract.FRectDesignerID = session("ssBctId")
	ooffontract.GetPartnerOffContractInfo

dim OReturnAddr
set OReturnAddr = new CCSReturnAddress
	OReturnAddr.FRectMakerid = session("ssBctId")
	OReturnAddr.GetBrandReturnAddress

%>
<script type="text/javascript">

function SaveUpcheInfo(frm){
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

	if (frm.return_zipcode.value.length<1){
		alert('사무실주소 우편번호를 선택하세요.');
		frm.return_zipcode.focus();
		return;
	}

	if (frm.return_address.value.length<1){
		alert('사무실주소1 을 입력하세요.');
		frm.return_address.focus();
		return;
	}

	if (frm.return_address2.value.length<1){
		alert('사무실주소2를 입력하세요.');
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

function SaveBrandReturnInfo(frm){
	var ret = confirm('브랜드 반품 정보를 저장 하시겠습니까?');

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
	}else if(flag=="b"){
		frmbrand.return_zipcode.value= post1 + "-" + post2;
		frmbrand.return_address.value= add;
		frmbrand.return_address2.value= dong;
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
	
	var ret = confirm('수정 하시겠습니까?');
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

function SaveBrandEtcInfo(frm){
	if (frm.socname_kor.value.length<1){
		alert('브랜드명 한글을 입력하세요.');
		frm.socname_kor.focus();
		return;
	}

	if (frm.socname.value.length<1){
		alert('브랜드명 영문을 입력하세요.');
		frm.socname.focus();
		return;
	}

//	if (!FileCheck(frm.logoimg,150000,160,110)){
//		frm.file1.focus();
//		return;
//	}

//	if (!FileCheck(frm.titleimg,150000,610,300)){
//		frm.file2.focus();
//		return;
//	}

	if (confirm('저장 하시겠습니까?')){
		frm.submit();
	}

}

function ChangeTitle(comp,imgcomp){
	imgcomp.src = comp.value;
}

function ChangeLogo(comp,imgcomp){
	imgcomp.src = comp.value;
}

function FileCheck(comp,maxfilesize,maxwidth,maxheight){
	if(comp.fileSize > maxfilesize){
		alert("파일사이즈는 "+ maxfilesize + "byte를 넘기실 수 없습니다...");
		return false;
	}

	if ((comp.src!="")&&(comp.width <1)){
		alert('이미지만 가능합니다.');
		return false;
	}

	//if(comp.width > maxwidth){
	//	alert("가로폭은 " + maxwidth + " 픽셀을 넘기실 수 없습니다...");
	//	return false;
	//}
	//if(comp.height > maxheight){
	//	alert("세로폭은 " + maxheight + " 픽셀을 넘기실 수 없습니다...");
	//	return false;
	//}

	return true;
}

function PopUpcheReturnAddrOnly(){
	var popwin = window.open("popupchereturnaddronly.asp","popupchereturnaddronly","width=1100 height=450 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>

<table width="600" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmupche" method="post" action="doupcheedit.asp">
<input type="hidden" name="mode" value="groupedit">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<input type="button" class="icon" value="#">
		<font color="red"><b>업체 사업자 정보</b></font>
	</td>
</tr>

<tr>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">업체코드</td>
	<td bgcolor="#FFFFFF" width="200">
		<input type="text" class="text_ro" name="groupid" value="<%= ogroup.FOneItem.FGroupId %>" size="7" maxlength="5" readonly>
	</td>
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">업체명</td>
	<td bgcolor="#FFFFFF" >
		<%= ogroup.FOneItem.FCompany_name %>
	</td>
</tr>

<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">입점브랜드ID</td>
	<td colspan="3" bgcolor="#FFFFFF"><%= DdotFormat(stripHTML(ogroup.FOneItem.getBrandList),100) %></td>
</tr>

<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">**업체 사업자등록정보**</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">회사명(상호)</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="company_name" value="<%= ogroup.FOneItem.FCompany_name %>" size="20" readonly>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">대표자</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="ceoname" value="<%= ogroup.FOneItem.Fceoname %>" size="20" readonly>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">사업자번호</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="company_no" value="<%= socialnoReplace(ogroup.FOneItem.Fcompany_no) %>" size="20" readonly>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">과세구분</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text_ro" name="" value="<%= ogroup.FOneItem.Fjungsan_gubun %>" size="20" readonly>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">사업장소재지</td>
	<td colspan="3" bgcolor="#FFFFFF" >
		<input type="text" class="text" name="company_zipcode" value="<%= ogroup.FOneItem.Fcompany_zipcode %>" size="7" maxlength="7">
		<input type="button" class="button" value="검색" onClick="TnFindZipNewdesigner('frmupche','C')">
		<input type="button" class="button" value="검색(구)" onclick="javascript:popZip('s');"><br>
		<input type="text" class="text" name="company_address" value="<%= ogroup.FOneItem.Fcompany_address %>" size="26" maxlength="64">&nbsp;
		<input type="text" class="text" name="company_address2" value="<%= ogroup.FOneItem.Fcompany_address2 %>" size="38" maxlength="64">
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">업태</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_uptae" value="<%= ogroup.FOneItem.Fcompany_uptae %>" size="20" maxlength="32"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">업종</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_upjong" value="<%= ogroup.FOneItem.Fcompany_upjong %>" size="20" maxlength="32"></td>
</tr>

<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">**업체 기본정보** &nbsp;&nbsp;(반품정보는 브랜드별로 입력할 수 있습니다.)</td>
</tr>

<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">대표전화</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_tel" value="<%= ogroup.FOneItem.Fcompany_tel %>" size="16" maxlength="16"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">팩스</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_fax" value="<%= ogroup.FOneItem.Fcompany_fax %>" size="16" maxlength="16"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">사무실주소</td>
	<td colspan="3" bgcolor="#FFFFFF" >
		<input type="text" class="text" name="return_zipcode" value="<%= ogroup.FOneItem.Freturn_zipcode %>" size="7" maxlength="7">
		<input type="button" class="button" value="검색" onClick="TnFindZipNewdesigner('frmupche','D')">
		<input type="button" class="button" value="검색(구)" onclick="javascript:popZip('m');">
		<input type="checkbox" class="checkbox" name=samezip onclick="SameReturnAddr(this.checked)">상동<br>
		<input type="text" class="text" name="return_address" value="<%= ogroup.FOneItem.Freturn_address %>" size="26" maxlength="64">&nbsp;
		<input type="text" class="text" name="return_address2" value="<%= ogroup.FOneItem.Freturn_address2 %>" size="38" maxlength="64">
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">브랜드 반품주소</td>
	<td colspan="3" bgcolor="#FFFFFF" >
		<input type="button" class="button" value="브랜드별 배송&CS 담당자 및 반품주소 설정" onclick="PopUpcheReturnAddrOnly()">
	</td>
</tr>

<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">**업체 결제계좌정보** &nbsp;&nbsp;(결제계좌 정보 수정시 담당MD에게 연락하시기 바랍니다.)</td>
</tr>

<tr height="26">
	<td bgcolor="<%= adminColor("tabletop") %>">거래은행</td>
	<td colspan="3" bgcolor="#FFFFFF" >
	<input type="text" class="text_ro" value="<%= ogroup.FOneItem.Fjungsan_bank %>" size="30" readonly>
	<% if (IsNULL(ogroup.FOneItem.Fjungsan_acctno)) or (ogroup.FOneItem.Fjungsan_acctno="") then %>
	<br>(결제 계좌를 등록하시려면  담당 MD에게 Fax로 통장 사본을 보내주시기 바랍니다.)
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

<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">**업체 담당자정보**</td>
</tr>

<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">담당자명</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_name" value="<%= ogroup.FOneItem.Fmanager_name %>" size="20" maxlength="32"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">일반전화</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_phone" value="<%= ogroup.FOneItem.Fmanager_phone %>" size="20" maxlength="16"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_email" value="<%= ogroup.FOneItem.Fmanager_email %>" size="25" maxlength="64"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">핸드폰</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_hp" value="<%= ogroup.FOneItem.Fmanager_hp %>" size="20" maxlength="16"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" height="30">배송&CS 담당자명</td>
	<td bgcolor="#FFFFFF" colspan="3">배송담당자 또는 CS담당자 정보는 아래 브랜드 반품정보에서 수정 가능합니다.</td>
</tr>

<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">정산담당자명</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_name" value="<%= ogroup.FOneItem.Fjungsan_name %>" size="20" maxlength="32"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">일반전화</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_phone" value="<%= ogroup.FOneItem.Fjungsan_phone %>" size="20" maxlength="16"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_email" value="<%= ogroup.FOneItem.Fjungsan_email %>" size="25" maxlength="64"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">핸드폰</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_hp" value="<%= ogroup.FOneItem.Fjungsan_hp %>" size="20" maxlength="16"></td>
</tr>
</form>

<tr align="center" height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<input type="button" class="button" value="업체정보 저장" onclick="SaveUpcheInfo(frmupche);">
	</td>
</tr>

</table>

<br>
<% if (opartner.FOneItem.Fuserdiv="14") then %>
<table width="600" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<input type="button" class="icon" value="#">
		<font color="red"><b>브랜드 계약관련 정보</b></font>
		(상품별로 마진율은 달라질 수 있습니다.)
	</td>
</tr>
<tr height="25">
	<td width="100" bgcolor="<%= adminColor("pink") %>">브랜드ID</td>
	<td bgcolor="#FFFFFF" width="200"><%= opartner.FOneItem.FID %></td>
	<td width="100" bgcolor="<%= adminColor("pink") %>"></td>
	<td bgcolor="#FFFFFF"></td>
</tr>
<% if (opartner.FOneItem.Fdiy_yn="Y") then %>
<tr height="25">
	<td bgcolor="<%= adminColor("pink") %>" >작품기본마진</td>
	<td bgcolor="#FFFFFF" >
		<%= opartner.FOneItem.GetMWUName %>&nbsp;
		<%= opartner.FOneItem.Fdiy_margin %> %
		&nbsp;&nbsp;(부가세포함)
	</td>
	<td bgcolor="<%= adminColor("pink") %>" >정산일 </td>
	<td bgcolor="#FFFFFF" >
	<%= opartner.FOneItem.Fjungsan_date %>
	</td>
</tr>
<% end if %>
<% if (opartner.FOneItem.Flec_yn="Y") then %>
<tr height="25">
	<td bgcolor="<%= adminColor("pink") %>" >강좌기본마진</td>
	<td bgcolor="#FFFFFF" >
		강좌 : <%= opartner.FOneItem.Flec_margin %> %
		재료비 : <%= opartner.FOneItem.Fmat_margin %> %
	</td>
	<td bgcolor="<%= adminColor("pink") %>" >정산일 </td>
	<td bgcolor="#FFFFFF" >
	<%= opartner.FOneItem.Fjungsan_date %>
	</td>
</tr>
<% end if %>
</table>
<% else %>
<table width="600" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<input type="button" class="icon" value="#">
		<font color="red"><b>브랜드 계약관련 정보</b></font>
		(상품별로 마진율은 달라질 수 있습니다.)
	</td>
</tr>
<tr height="25">
	<td width="100" bgcolor="<%= adminColor("pink") %>">브랜드ID</td>
	<td bgcolor="#FFFFFF" width="200"><%= opartner.FOneItem.FID %></td>
	<td width="100" bgcolor="<%= adminColor("pink") %>">상품수</td>
	<td bgcolor="#FFFFFF"><%= opartner.FOneItem.FTotalitemcount %></td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("pink") %>" >온라인기본마진</td>
	<td bgcolor="#FFFFFF" >
		<%= opartner.FOneItem.GetMWUName %>&nbsp;
		<%= opartner.FOneItem.Fdefaultmargine %> %
	</td>
	<td bgcolor="<%= adminColor("pink") %>" >정산일 </td>
	<td bgcolor="#FFFFFF" >
	<%= opartner.FOneItem.Fjungsan_date %>
	</td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("pink") %>" >오프라인(직영점)</td>
	<td bgcolor="#FFFFFF" >
		<table border=0 cellspacing=0 cellpadding=0 class=a>
			<tr>
				<td width="90"><b>직영점대표</b></td>
				<td width="80"><%= ooffontract.GetSpecialChargeDivName("streetshop000") %></td>
				<td><%= ooffontract.GetSpecialDefaultMargin("streetshop000") %> %</td>
			</tr>
			<% for i=0 to ooffontract.FResultCount-1 %>
			<% if (ooffontract.FItemList(i).Fshopdiv="1")  then %>
			<tr>
				<td><%= ooffontract.FItemList(i).Fshopname %></td>
				<td><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
				<td><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
			</tr>
			<% end if %>
			<% next %>
		</table>
	</td>
	<td bgcolor="<%= adminColor("pink") %>" >정산일 </td>
	<td bgcolor="#FFFFFF"><%= opartner.FOneItem.Fjungsan_date_off %></td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("pink") %>" >오프라인(가맹점)</td>
	<td bgcolor="#FFFFFF" >
		<table border=0 cellspacing=0 cellpadding=0 class=a>
			<tr>
				<td width="90"><b>가맹점점대표</b></td>
				<td width="80"><%= ooffontract.GetSpecialChargeDivName("streetshop800") %></td>
				<td><%= ooffontract.GetSpecialDefaultMargin("streetshop800") %> %</td>
			</tr>
			<% for i=0 to ooffontract.FResultCount-1 %>
			<% if (ooffontract.FItemList(i).Fshopdiv="3")  then %>
			<tr>
				<td><%= ooffontract.FItemList(i).Fshopname %></td>
				<td><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
				<td><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
			</tr>
			<% end if %>
			<% next %>

			<% for i=0 to ooffontract.FResultCount-1 %>
			<% if (ooffontract.FItemList(i).Fshopdiv="5")  then %>
			<tr>
				<td><%= ooffontract.FItemList(i).Fshopname %></td>
				<td><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
				<td><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
			</tr>
			<% end if %>
			<% next %>
		</table>
	</td>
	<td bgcolor="<%= adminColor("pink") %>">정산일 </td>
	<td bgcolor="#FFFFFF"><%= opartner.FOneItem.Fjungsan_date_frn %></td>
</tr>
</table>
<% end if %>
<br>

<table width="600" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmbrand" method="post" action="doupcheedit.asp">
<input type="hidden" name="mode" value="brandedit">
<input type="hidden" name="groupid" value="<%= ogroup.FOneItem.FGroupId %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<input type="button" class="icon" value="#">
		<font color="red"><b>브랜드 배송담당 및 반품주소</b></font>
		(업체배송상품의 경우 아래 반품주소를 고객님께 안내해 드립니다.)
	</td>
</tr>
<tr height="25">
	<td width="100" bgcolor="<%= adminColor("pink") %>">배송담당자</td>
	<td bgcolor="#FFFFFF" width="200"><input type="text" class="text" name="deliver_name" value="<%= OReturnAddr.FreturnName %>" size="16" maxlength="16"></td>
	<td width="100" bgcolor="<%= adminColor("pink") %>">전화번호</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_phone" value="<%= OReturnAddr.FreturnPhone %>" size="16" maxlength="16"></td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("pink") %>">핸드폰번호</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_hp" value="<%= OReturnAddr.Freturnhp %>" size="16" maxlength="16"></td>
	<td bgcolor="<%= adminColor("pink") %>">이메일주소</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="deliver_email" value="<%= OReturnAddr.FreturnEmail %>" size="16" maxlength="128"></td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("pink") %>" >반품주소</td>
	<td bgcolor="#FFFFFF" colspan=3>
		<input type="text" class="text" name="return_zipcode" value="<%= OReturnAddr.FreturnZipcode %>" size="7" maxlength="7">
		<input type="button" class="button" value="검색" onClick="TnFindZipNewdesigner('frmbrand','D')">
		<input type="button" class="button" value="검색(구)" onclick="javascript:popZip('b');"><br>
		<input type="text" class="text" name="return_address" value="<%= OReturnAddr.FreturnZipaddr %>" size="26" maxlength="64">&nbsp;
		<input type="text" class="text" name="return_address2" value="<%= OReturnAddr.FreturnEtcaddr %>" size="38" maxlength="64">
	</td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("pink") %>" >택배사</td>
	<td bgcolor="#FFFFFF" colspan=3>
		<% drawSelectBoxDeliverCompany "defaultsongjangdiv" , OReturnAddr.Fsongjangdiv %>
	</td>
</tr>
<!--
<tr height="25">
	<td bgcolor="<%= adminColor("pink") %>" ></td>
	<td bgcolor="#FFFFFF" colspan=3>
		<input type="checkbox" class="checkbox" name=applyallbrand value="Y"> 업체내 모든 브랜드 일괄 수정
	</td>
</tr>
-->
</form>

<tr align="center" height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<input type="button" class="button" value="브랜드 반품정보 저장" onclick="SaveBrandReturnInfo(frmbrand);">
	</td>
</tr>
</table>

<br>
<% if (opartner.FOneItem.Fuserdiv="14") then %>
<!-- 표시 않함 2016/08/22-->
<% else %>
<!-- <table width="600" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>"> -->
<!-- <form name="frmetc" method="post" action="<%= uploadImgUrl %>/linkweb/partner_info/doprofileimageadmin.asp" enctype="multipart/form-data"> -->
<!-- <input type="hidden" name="designerid" value="<%= opartner.FOneItem.FID %>"> -->
<!-- <tr height="25" bgcolor="FFFFFF"> -->
<!-- 	<td colspan="4"> -->
<!-- 		<input type="button" class="icon" value="#"> -->
<!-- 		<font color="red"><b>브랜드 관련정보(웹표시정보)</b></font> -->
<!-- 	</td> -->
<!-- </tr> -->
<!-- <tr> -->
<!-- 	<td width="120" bgcolor="<%= adminColor("sky") %>">브랜드명(한글)</td> -->
<!-- 	<td width="180" bgcolor="#FFFFFF"> -->
<!-- 		<input type="text" class="text" name="socname_kor" value="<%= opartner.FOneItem.Fsocname_kor %>" size=20 maxlength="20"> -->
<!-- 	</td> -->
<!-- 	<td width="120" bgcolor="<%= adminColor("sky") %>">브랜드명(영문)</td> -->
<!-- 	<td bgcolor="#FFFFFF"> -->
<!-- 		<input type="text" class="text" name="socname" value="<%= opartner.FOneItem.Fsocname %>" size=20 maxlength="20"> -->
<!-- 	</td> -->
<!-- </tr> -->
<!-- <tr> -->
<!-- 	<td bgcolor="<%= adminColor("sky") %>">로고</td> -->
<!-- 	<td bgcolor="#FFFFFF" colspan="3"> -->
<!-- 		<img name="logoimg" src="<%= opartner.FOneItem.getSocLogoUrl %>" width=150 height=100><br> -->
<!-- 		(브랜드 로고는 150x100 픽셀로 업로드 해주시기 바랍니다.)<br> -->
<!-- 		<input type="file" class="file" name="file1" size="40"><!--  onchange="ChangeLogo(this,frmetc.logoimg);" -->
<!-- 	</td> -->
<!-- </tr> -->
<!-- <tr> -->
<!-- 	<td bgcolor="<%= adminColor("sky") %>" >배너</td> -->
<!-- 	<td bgcolor="#FFFFFF" colspan="3"> -->
<!-- 	<img name=titleimg src="<%= opartner.FOneItem.getTitleImgUrl %>" width=300 height=75><br> -->
<!-- 	(이미지는 720x220 픽셀로 업로드 해주시기 바랍니다.)<br> -->
<!-- 	<input type=file name=file2 size=40><!--  onchange="ChangeTitle(this,frmetc.titleimg);" --> 
<!-- 	</td> -->
<!-- </tr> -->
<!-- <tr> -->
<!-- 	<td bgcolor="<%= adminColor("sky") %>">브랜드<br>코멘트</td> -->
<!-- 	<td bgcolor="#FFFFFF" colspan="3"> -->
<!-- 	<textarea class="textarea" name="dgncomment" cols="80" rows="6"><%= opartner.FOneItem.Fdgncomment %></textarea> -->
<!-- 	</td> -->
<!-- </tr> -->
<!-- </form> -->
<!--  -->
<!-- <tr align="center" height="25" bgcolor="FFFFFF"> -->
<!-- 	<td colspan="15"> -->
<!-- 		<input type="button" class="button" value="브랜드 정보 수정" onclick="SaveBrandEtcInfo(frmetc);"> -->
<!-- 	</td> -->
<!-- </tr> -->
<!-- </table> -->
<table width="600" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmetc" method="post" action="<%= uploadImgUrl %>/linkweb/partner_info/doprofileimageadmin.asp" enctype="multipart/form-data">
<input type=hidden name=designerid value="<%= opartner.FOneItem.FID %>">
<tr>
	<td width="120" bgcolor="<%= adminColor("sky") %>">브랜드명(한글)</td>
	<td width="180" bgcolor="#FFFFFF">
		<input type="text" class="text" name="socname_kor" value="<%= opartner.FOneItem.Fsocname_kor %>" size=20 maxlength="20">
	</td>
	<td width="120" bgcolor="<%= adminColor("sky") %>">브랜드명(영문)</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="socname" value="<%= opartner.FOneItem.Fsocname %>" size=20 maxlength="20">
	</td>
</tr>
<% if (FALSE) then %>
<tr>
	<td bgcolor="<%= adminColor("sky") %>" >프로필 이미지</td>
	<td bgcolor="#FFFFFF" colspan="3">
	<img name=brandimg src="<%= opartner.FOneItem.getBrandImgUrl("") %>" width=<%=600/2%> height=<%=600/2%>><br>
	(프로필 이미지는 600X600 픽셀이상으로 지정되어 있습니다.)<br>
	<img src="<%= opartner.FOneItem.getBrandImgUrl("1") %>" width=<%=400/2%> height=<%=400/2%>>&nbsp;
	<img src="<%= opartner.FOneItem.getBrandImgUrl("2") %>" width=<%=200/2%> height=<%=200/2%>>&nbsp;
	<img src="<%= opartner.FOneItem.getBrandImgUrl("3") %>" width=<%=100/2%> height=<%=100/2%>>&nbsp;<br/>
	<input type="file" class="button" name="file4" size="60" onclick="ChangeTitle(this,frmetc.brandimg);">
	<% If opartner.FOneItem.getBrandImgUrl("") <> "http://webimage.10x10.co.kr/image/brandlogo/" Then %>
		<input type="checkbox" name="deltitleimg" size="60" value="Y">삭제
	<% End If %>
	</td>
</tr>
<% end if %>
<tr>
	<td bgcolor="<%= adminColor("sky") %>">브랜드<br>코멘트</td>
	<td bgcolor="#FFFFFF" colspan="3">
	<textarea class="textarea" name="dgncomment" cols="80" rows="6"><%= opartner.FOneItem.Fdgncomment %></textarea>
	</td>
</tr>
<tr>
	<td colspan="4" align=center bgcolor="#FFFFFF"><input type="button" class="button" value="브랜드 기타정보 저장" onclick="SaveBrandEtcInfo(frmetc);"></td>
</tr>
</form>
</table>
<br>
<% end if %>

<table width="600" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmpass" method=post action="doupcheedit.asp">
<input type="hidden" name="mode" value="editpass">
<input type="hidden" name="groupid" value="<%= ogroup.FOneItem.FGroupId %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="2">
		<input type="button" class="icon" value="#">
		<font color="red"><b>비밀번호 변경</b></font>
		&nbsp;
		(브랜드 비밀번호를 변경하시려면 아래 란을 채워 주시기바랍니다.)
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">기존비밀번호</td>
	<td width="480" bgcolor="#FFFFFF">
		1차 :<input type="password" class="text" name="txoldpassword" size="12" value="" maxlength="32">
		2차 :<input type="password" class="text" name="txoldpasswordS" size="12" value="" maxlength="32">
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">변경비밀번호 1차 </td>
	<td width="480" bgcolor="#FFFFFF">
		입력: <input type="password" class="text" name="txnewpassword1" size="12" value="" maxlength="32"><br>
		확인: <input type="password" class="text" name="txnewpassword2" size="12" value="" maxlength="32">
	</td>
</tr> 
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">변경비밀번호 2차 </td>
	<td width="480" bgcolor="#FFFFFF">
		입력: <input type="password" class="text" name="txnewpasswordS1" size="12" value="" maxlength="32"><br>
		확인: <input type="password" class="text" name="txnewpasswordS2" size="12" value="" maxlength="32">
	</td>
</tr>
</form>

<tr align="center" height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<input type="button" class="button" value="브랜드 비밀번호변경" onclick="EditPass(frmpass);">
	</td>
</tr>
</table>

<%
set ogroup = Nothing
set opartner = Nothing
set ooffontract = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
