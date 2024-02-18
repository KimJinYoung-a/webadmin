<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 브랜드 정보
' History : 서동석 생성
'           2021.06.18 한용민 수정(담당자 휴대폰,이메일 인증정보 데이터쪽에도 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->

<%
dim ogroup,i, groupid
dim oJungsanDiff, j

	groupid = requestCheckvar(request("groupid"),32)

set ogroup = new CPartnerGroup
	ogroup.FRectGroupid = groupid
	ogroup.GetOneGroupInfo

set oJungsanDiff = new CPartnerGroup
    oJungsanDiff.FRectGroupid = groupid
	oJungsanDiff.GetGroupPartnerJungsanDiffList
	
Dim IsNewReg : IsNewReg = (ogroup.FResultCount<1)

dim ogroupuser
set ogroupuser = new CPartnerGroup
	ogroupuser.FPageSize = 100
	ogroupuser.FCurrPage = 1
    ogroupuser.frectgroupid = groupid
    ogroupuser.Get_partner_user_list

dim existsetccount
	existsetccount=0

if ogroupuser.FResultCount > 0 then
for i = 0 to ogroupuser.FResultCount-1
	' 10 추가담당자
	if ogroupuser.FItemList(i).fgubun="10" then
		existsetccount = existsetccount + 1
	end if
next
end if
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

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

function SaveUpcheInfo(frm){
    var psocno =frm.psocno.value;
    <% if (IsNewReg) then %>
    if (frm.validchk.value!="Y"){
        alert('사업자 번호 중복확인 후 등록가능합니다.');
        return;
    }
    <% else %>
    if ((psocno!='')&&(psocno!=frm.company_no.value)){
        <% if (C_ADMIN_AUTH=true or C_OFF_AUTH or C_MD_AUTH) then %>
        if (!confirm('사업자번호 변경시 업체를 새로 등록 해야 합니다. 계속하시겠습니까?')) return;
        <% else %>
        alert('사업자번호 변경시 업체를 새로 등록 해야 합니다. ');
        return;
        <% end if %>
    }
    <% end if %>

	var errMsg = chkIsValidJungsanGubun(frm.company_no.value, frm.jungsan_gubun.value);
	if (errMsg != "OK") {
		alert(errMsg);
		retutn;
	}

    if (frm.company_name.value.length<1){
		alert('사업자 등록상의 회사명을 입력하세요.');
		frm.company_name.focus();
		return;
	}

	if (frm.ceoname.value.length<1){
		alert('사업자 등록상의 대표자명을 입력하세요.');
		frm.ceoname.focus();
		return;
	}

	if (frm.company_no.value.length<1){
		alert('사업자 등록 번호를 입력하세요.');
		frm.company_no.focus();
		return;
	}

	if (frm.jungsan_gubun.value.length<1){
		alert('과세구분을 선택하세요.');
		frm.jungsan_gubun.focus();
		return;
	}

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

	<% ' 정산담당자 데이터 필수값으로.. 정산시 곤란 %>
	if (frm.jungsan_name.value.length<1){
		alert('정산담당자 성함을 입력하세요.');
		frm.jungsan_name.focus();
		return;
	}
	if (frm.jungsan_phone.value.length<1){
		alert('정산담당자 전화번호를 입력하세요.');
		frm.jungsan_phone.focus();
		return;
	}
	if (frm.jungsan_email.value.length<1){
		alert('정산담당자 이메일을 입력하세요.');
		frm.jungsan_email.focus();
		return;
	}
	if (frm.jungsan_hp.value.length<1){
		alert('정산담당자 핸드폰을 입력하세요.');
		frm.jungsan_hp.focus();
		return;
	}

    <% if C_ADMIN_AUTH or C_OFF_AUTH or C_MD_AUTH then %>
	if (frm.jungsan_date.value.length<1){
		alert('정산일을 선택하세요.');
		frm.jungsan_date.focus();
		return;
	}

    if (frm.jungsan_date_off.value.length<1){
		alert('오프 정산일을 선택하세요. - 기본은 온라인과 동일합니다.');
		frm.jungsan_date_off.focus();
		return;
	}
	<% end if %>

	<% if existsetccount>1 then %>
		for(var i=0; i<frm.etc_idx.length; i++){
			if(frm.etc_name[i].value==''){
				alert('추가담당자명을 입력해 주세요.');
				frm.etc_name[i].focus()
				return;
			}
			if(frm.etc_hp[i].value==''){
				alert('추가담당자의 핸드폰 번호를 입력해 주세요.');
				frm.etc_hp[i].focus()
				return;
			}
			if(frm.etc_email[i].value==''){
				alert('추가담당자의 이메일을 입력해 주세요.');
				frm.etc_email[i].focus()
				return;
			}
		}
	<% elseif existsetccount=1 then %>
		if(frm.etc_name.value==''){
			alert('추가담당자명을 입력해 주세요.');
			frm.etc_name.focus()
			return;
		}
		if(frm.etc_hp.value==''){
			alert('추가담당자의 핸드폰 번호를 입력해 주세요.');
			frm.etc_hp.focus()
			return;
		}
		if(frm.etc_email.value==''){
			alert('추가담당자의 이메일을 입력해 주세요.');
			frm.etc_email.focus()
			return;
		}
	<% end if %>
	
	if(frm.addetc_name.value!=''){
		if(frm.addetc_hp.value==''){
			alert('추가담당자의 핸드폰 번호를 입력해 주세요.');
			frm.addetc_hp.focus()
			return;
		}
		if(frm.addetc_email.value==''){
			alert('추가담당자의 이메일을 입력해 주세요.');
			frm.addetc_email.focus()
			return;
		}
		frm.etcaddyn.value="Y"
	}

	if (frm.groupid.value.length<1){
		var ret = confirm('업체 정보를 저장 하시겠습니까?');
	}else{
		var ret = confirm('같은 그룹코드에 속한 브랜드정보도 일괄 수정됩니다.(반품주소,배송담당자정보는 제외) \n\n저장 하시겠습니까?');
	}

	if (ret){
		frm.submit();
	}
}

var bValidSoc = false;
function SearchSocno(frm){
    <% if (IsNewReg) then %>
    if (bValidSoc){
        clearSocField(false);
        return;
    }
    <% end if %>

	if (frm.company_no.value.length<1){
		alert('사업자 등록 번호를 입력하세요.');
		frm.company_no.focus();
		return;
	}

	if (frm.company_no.value.length != 12){
		alert('사업자 등록 번호는 000-00-00000 형식으로 입력해야 합니다.');
		frm.company_no.focus();
		return;
	}

	if (frm.groupid.value.length<1){
		icheckframe.location.href="icheckframe.asp?mode=CheckSocno&socno=" + frm.company_no.value;
	}else{
	    //기존 존재하는 사업자로 변경 불가
	    var psocno =frm.psocno.value;
        if ((psocno!='')&&(psocno!=frm.company_no.value)){
	        icheckframe.location.href="icheckframe.asp?mode=CheckSocno&socno=" + frm.company_no.value;
	    }
		// alert('사업자번호를 변경할 경우 기존 정보가 변경됩니다.');
	}

}

function AddProc(mode){
	alert('등록가능한 사업자번호입니다.');
	clearSocField(true);
}

function clearSocField(bool){
    bValidSoc = bool;
    <% if (IsNewReg) then %>
	var frm = document.frmupche;

	if (!bValidSoc){
	    frm.company_no.value="";
	    frm.validchk.value="";
	    frm.company_no.style.backgroundColor ="#FFFFFF";
	}else{
	    frm.validchk.value="Y";
	    frm.company_no.style.backgroundColor ="#EEEEEE";
	}
	frm.company_no.readOnly=(bValidSoc);
	if ( document.getElementById("coSearchBtn") ) {
		document.getElementById("coSearchBtn").value=(bValidSoc)?"재입력":"중복확인";
	}


	if (!bool){frm.company_no.focus();}
	<% end if %>
}

function PopUpcheList(frmname){
	var popwin = window.open("/admin/lib/popupchelist.asp?frmname=" + frmname,"popupchelist","width=700 height=540 scrollbars=yes resizable=yes");
	popwin.focus();
}

function PopUpcheReturnAddrOnly(groupid){
	if (groupid == "") {
		alert("그룹코드가 없습니다.");
		return;
	}


	var popwin = window.open("/admin/member/popupchereturnaddronly.asp?groupid=" + groupid,"popupchereturnaddronly","width=900 height=580 scrollbars=yes resizable=yes");
	popwin.focus();
}

function EditErpCustCD(){
    var frm = document.frmupche;

    var param = "?groupid="+frm.groupid.value
    var popwin = window.open('popErpGroupLinkEdit.asp'+param,'popErpGroupLinkEdit','width=600,height=300,scrollbars=yes,resizable=yes');
    popwin.focus();
}

var orgjungsan_gubun = "<%= ogroup.FOneItem.Fjungsan_gubun %>";
if (orgjungsan_gubun == "") {
	orgjungsan_gubun = "일반과세";
}
function fnJungsanGubunChanged() {
	var frm = document.frmupche;
	var company_no = document.getElementById("company_no");

	<% 
	'이문재 이사님 요청(과세구분을 영세로 바꾸어야 합니다.등록할때 입력이 안되고, 수정도 막혀이 있네요)으로 주석처리.	' 2022.02.23 한용민
	'if (Not IsNewReg) then 
	%>
	//	if ((orgjungsan_gubun != frm.jungsan_gubun.value) && ((orgjungsan_gubun == "영세(해외)") || (frm.jungsan_gubun.value == "영세(해외)"))) {
	//		alert("영세(해외) - 일반사업자 간 과세구분을 변경할 수 없습니다.\n\n업체를 새로 등록하세요.");
	//		frm.jungsan_gubun.value = orgjungsan_gubun;
	//		return;
	//	}
	<% 'end if %>

	<% if (Not IsNewReg) then %>
		return;
	<% end if %>

	if ((orgjungsan_gubun != "영세(해외)") && (frm.jungsan_gubun.value != "영세(해외)")) {
		orgjungsan_gubun = frm.jungsan_gubun.value;
		return;
	}
	orgjungsan_gubun = frm.jungsan_gubun.value;

	if (frm.jungsan_gubun.value == "영세(해외)") {
		// 해외는 사업자번호 자동설정된다(888-00-00000)

		company_no.className = "text_ro";
		frm.company_no.readOnly = true;
		frm.company_no.value = "888-00-00000";

		if (frm.coSearchBtn) {
			frm.coSearchBtn.disabled = true;
		}

		frm.validchk.value = "Y";
	} else {
		company_no.className = "text";
		frm.company_no.readOnly = false;
		frm.company_no.value = "";

		if (frm.coSearchBtn) {
			frm.coSearchBtn.disabled = false;
		}

		frm.validchk.value = "N";
	}
}

function chkIsValidJungsanGubun(company_no, jungsan_gubun) {
	// 000-00-00000
	// 가운데 두글자 : 구분코드
	// =========================================================================
	// 01-79 : 개인사업자+과세사업자
	// 90-99 : 개인사업자+면세사업자
	// 기타 : 과세 면세 모두 가능
	// 앞자리 888 = 영세(해외)
	// =========================================================================

	if (company_no.length != 12) {
		// 주문등록번호(원천징수) 등 체크 않함.
		return "OK";
	}

	var soc_gubun = company_no.substring(4, 6)*1;
	var IsForeign = (company_no.substring(0, 3) == "888");

	if (IsForeign) {
		if (jungsan_gubun != "영세(해외)") {
			return "영세(해외) 사업자만 가능한 사업자번호입니다.";
		}

		return "OK";
	} else {
		if (jungsan_gubun == "영세(해외)") {
			return "영세(해외) 사업자로 변경 불가능한 사업자번호입니다.";
		}

		/*
		if ((soc_gubun >= 1) && (soc_gubun <= 79)) {
			if (jungsan_gubun == "면세") {
				return "면세로 등록할 수 없는 사업자번호입니다.";
			}

			return "OK";
		}
		*/

		if ((soc_gubun >= 90) && (soc_gubun <= 99)) {
			if (jungsan_gubun != "면세") {
				return "면세로만 등록가능한 사업자번호입니다.";
			}

			return "OK";
		}

		return "OK";
	}
}

function divadd(){
	document.all.divadd1.style.display ='';
	document.all.divadd2.style.display ='';
	document.all.divadd3.style.display ='';
}

function etc_del(etc_idx){
	if (etc_idx==''){
		alert('정상적인 경로가 아닙니다.\n지정된 번호가 없습니다.');
		return;
	}
	frmetcedit.etc_idx.value=etc_idx;
	frmetcedit.submit()
}

$(function(){
	<% if C_ADMIN_AUTH or C_MngPart or C_MD_AUTH or C_PSMngPart then %>
		// 브랜드별 추가정산 계좌 추가정산		// 2020.02.11 한용민 생성
		$("#buttonregdiffjungsan").click(function(){
			$("#regdiffacct1").show();
			$("#regdiffacct2").show();
			$("#regdiffacct3").show();
			$("#regdiffacct4").show();
			$("#regdiffacct5").show();
		})
	<% end if %>
})

</script>

<form name="frmupche" method="post" action="/admin/member/doupcheedit.asp">
<input type="hidden" name="mode" value="groupedit">
<input type="hidden" name="uid" value="">
<input type="hidden" name="psocno" value="<%= ogroup.FOneItem.Fcompany_no %>">
<input type="hidden" name="validchk" value="">
<input type="hidden" name="etcaddyn" value="" />
<table width="600" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		업체코드 : <%= ogroup.FOneItem.FGroupId %>&nbsp;&nbsp;
        업체명 : <%= ogroup.FOneItem.FCompany_name %>
	</td>
</tr>
<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan=4><b>1.업체관련정보</b></td>
</tr>
<tr height="25">
	<td width="120" bgcolor="<%= adminColor("tabletop") %>">업체코드</td>
	<td bgcolor="#FFFFFF" width="200">
		<input type="text" class="text" name="groupid" value="<%= ogroup.FOneItem.FGroupId %>" size="7" maxlength="5" style="background-color:#EEEEEE;" readonly>
	</td>
	<td width="90" bgcolor="<%= adminColor("tabletop") %>">ERP연계코드</td>
	<td bgcolor="#FFFFFF" width="200">
	<% if (ogroup.FOneItem.FGroupId<>"") then %>
	    <% if (ogroup.FOneItem.FerpUsing<>1) then %>
	        연동안함
	    <% else %>
		    <% if IsNULL(ogroup.FOneItem.FerpCust_CD) then %>
		    <%= ogroup.FOneItem.FGroupId %> <!--기본-->
		    <% else %>
    		    <% if ogroup.FOneItem.FerpCUST_USE_CD<>ogroup.FOneItem.FerpCust_CD then %>
    			<strong><%= ogroup.FOneItem.FerpCUST_USE_CD%></strong>(<%= ogroup.FOneItem.FerpCust_CD %>)
    			<% else %>
    			<%= ogroup.FOneItem.FerpCust_CD %>
    		    <% end if %>
			<% end if %>
		<% end if %>

		<% if (C_ADMIN_AUTH) or (session("ssAdminPsn") = "7") or (C_MngPart) or C_OFF_AUTH or C_MD_AUTH or C_PSMngPart then %>
			<input type="button" class="button" value="수정" onClick="EditErpCustCD()">
		<% end if %>
	<% end if %>
	</td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("tabletop") %>">입점브랜드ID</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<% if ogroup.FOneItem.getBrandList = "" then %>
			<font color="red">현재 진행중인 브랜드가 없습니다.</font>
		<% else %>
			<%= ogroup.FOneItem.getBrandListHTML %>
		<% end if %>
	</td>
</tr>

<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">**사업자등록정보**</td>
</tr>

<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">회사명(상호)</td>
	<td bgcolor="#FFFFFF">
		<% if (C_ADMIN_AUTH=true or C_OFF_AUTH or C_MD_AUTH) then %>
		<input type="text" class="text" name="company_name" value="<%= ogroup.FOneItem.FCompany_name %>" size="30" maxlength="50">
		<% else %>
		<input type="text" class="text" name="company_name" value="<%= ogroup.FOneItem.FCompany_name %>" size="30" maxlength="50" style="background-color:#EEEEEE;" readonly>
		<% end if %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">대표자</td>
	<td bgcolor="#FFFFFF">
		<% if (C_ADMIN_AUTH=true or C_OFF_AUTH or C_MD_AUTH) then %>
		<input type="text" class="text" name="ceoname" value="<%= ogroup.FOneItem.Fceoname %>" size="23" maxlength="17">
		<% else %>
		<input type="text" class="text" name="ceoname" value="<%= ogroup.FOneItem.Fceoname %>" size="23" maxlength="17" style="background-color:#EEEEEE;" readonly>
		<% end if %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">사업자번호</td>
	<td bgcolor="#FFFFFF">
		<% if ((C_ADMIN_AUTH=true) or (C_OFF_AUTH) or C_MD_AUTH or (session("ssAdminPsn") = "7") or (IsNewReg)) and (LEN(TRIM(replace(ogroup.FOneItem.Fcompany_no,"-","")))=10) then %>
		<input type="text" class="text" id="company_no" name="company_no" value="<%= ogroup.FOneItem.Fcompany_no %>" size="16" maxlength="20" >
		<input id="coSearchBtn" name="coSearchBtn" type="button" class="button" value="중복확인" onClick="SearchSocno(frmupche)">
		<% else %>
		<input type="text" class="text" id="company_no" name="company_no" value="<%= ogroup.FOneItem.Fcompany_no %>" size="16" maxlength="20" style="background-color:#EEEEEE;" readonly>
		<% end if %>

	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">과세구분</td>
	<td bgcolor="#FFFFFF">
			<% if  (C_ADMIN_AUTH=true or C_OFF_AUTH or C_MD_AUTH) then %>
			<% drawSelectBoxjungsan_gubun "jungsan_gubun",ogroup.FOneItem.Fjungsan_gubun,"fnJungsanGubunChanged()" %>
		<%else%>
		<%=ogroup.FOneItem.Fjungsan_gubun%><input type="hidden" name="jungsan_gubun" value="<%=ogroup.FOneItem.Fjungsan_gubun%>"> 
		<%end if%>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">사업장소재지</td>
	<td colspan="3" bgcolor="#FFFFFF" >
		<input type="text" class="text" name="company_zipcode" value="<%= ogroup.FOneItem.Fcompany_zipcode %>" size="7" maxlength="7">
        <input type="button" class="button" value="검색" onClick="FnFindZipNew('frmupche','C')">
		<input type="button" class="button" value="검색(구)" onClick="TnFindZipNew('frmupche','C')">
        <% '<input type="button" class="button" value="검색(구)" onClick="popZip('s');"> %>
		<br>
		<input type="text" class="text" name="company_address" value="<%= ogroup.FOneItem.Fcompany_address %>" size="30" maxlength="64">&nbsp;
		<input type="text" class="text" name="company_address2" value="<%= ogroup.FOneItem.Fcompany_address2 %>" size="46" maxlength="64">
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">업태</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_uptae" value="<%= ogroup.FOneItem.Fcompany_uptae %>" size="30" maxlength="32"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">업종</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_upjong" value="<%= ogroup.FOneItem.Fcompany_upjong %>" size="30" maxlength="32"></td>
</tr>

<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">**업체기본정보**</td>
</tr>

<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">대표전화</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_tel" value="<%= ogroup.FOneItem.Fcompany_tel %>" size="16" maxlength="16" onFocusOut="phone_format(frmupche.company_tel)"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">팩스</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="company_fax" value="<%= ogroup.FOneItem.Fcompany_fax %>" size="16" maxlength="16" onFocusOut="phone_format(frmupche.company_fax)"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">사무실 주소</td>
	<td colspan="3" bgcolor="#FFFFFF" >
	<input type="text" class="text" name="return_zipcode" value="<%= ogroup.FOneItem.Freturn_zipcode %>" size="7" maxlength="7">
    <input type="button" class="button" value="검색" onClick="FnFindZipNew('frmupche','D')">
	<input type="button" class="button" value="검색(구)" onClick="TnFindZipNew('frmupche','D')">
    <% '<input type="button" class="button" value="검색(구)" onClick="popZip('m');"> %>
	<input type=checkbox name=samezip onclick="SameReturnAddr(this.checked)">상동
	<br>
		<input type="text" class="text" name="return_address" value="<%= ogroup.FOneItem.Freturn_address %>" size="30" maxlength="64">&nbsp;
		<input type="text" class="text" name="return_address2" value="<%= ogroup.FOneItem.Freturn_address2 %>" size="46" maxlength="64">
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">반품주소</td>
	<td colspan="3" bgcolor="#FFFFFF" >
	<input type="button" class="button" value="브랜드별 배송담당자 및 반품주소 설정" onClick="PopUpcheReturnAddrOnly('<%= ogroup.FOneItem.FGroupId %>')">
	</td>
</tr>
<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">
		**결제계좌정보**
		<% if C_ADMIN_AUTH or C_MngPart or C_MD_AUTH or C_PSMngPart then %>
			<input type="button" id="buttonregdiffjungsan" value="브랜드별 결제계좌 추가(대표 정산계좌와 다를경우)" class="button" >
		<% end if %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">거래은행</td>
	<td colspan="3" bgcolor="#FFFFFF" >
		<% if C_ADMIN_AUTH or C_OFF_AUTH or C_MD_AUTH then %>
			<% DrawBankCombo "jungsan_bank", ogroup.FOneItem.Fjungsan_bank %>
		<% else %>
			<%= ogroup.FOneItem.Fjungsan_bank %><input type="hidden" name="jungsan_bank" value="<%= ogroup.FOneItem.Fjungsan_bank %>"> 
		<% end if %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">계좌번호</td>
	<td colspan="3" bgcolor="#FFFFFF" >
	<% if (C_ADMIN_AUTH) or (C_MngPart) then %>
	<input type="text" class="text" name="jungsan_acctno" value="<%= ogroup.FOneItem.Fjungsan_acctno %>" size="24" maxlength="32">
	<% else %>
	<input type="text" class="text_RO" name="jungsan_acctno" value="<%= ogroup.FOneItem.Fjungsan_acctno %>" size="20" maxlength="32" readOnly >
	<font color="red">(경영지원팀 변경가능)</font>
	<% end if %>
	&nbsp;&nbsp; '-'은 빼고 번호만 입력해주시기 바랍니다.
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">예금주명</td>
	<td colspan="3" bgcolor="#FFFFFF" >
	<% if (C_ADMIN_AUTH) or (C_MngPart) then %>
	<input type="text" class="text" name="jungsan_acctname" value="<%= ogroup.FOneItem.Fjungsan_acctname %>" size="24" maxlength="16">
	<% else %>
	<input type="text" class="text_ro" name="jungsan_acctname" value="<%= ogroup.FOneItem.Fjungsan_acctname %>" size="20" maxlength="16" readOnly >
	<font color="red">(경영지원팀 변경가능)</font>
	<% end if %>

	&nbsp;&nbsp; 띄어쓰기 하지 마시기 바랍니다.
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">정산일</td>
	<td colspan="3" bgcolor="#FFFFFF" >
		<% if C_ADMIN_AUTH or C_OFF_AUTH or C_MD_AUTH then %>
			온라인 : <% DrawJungsanDateCombo "jungsan_date", ogroup.FOneItem.Fjungsan_date %>
			&nbsp;
			오프라인 : <% DrawJungsanDateCombo "jungsan_date_off", ogroup.FOneItem.Fjungsan_date_off %>
		<% else %>
			온라인 : <%= ogroup.FOneItem.Fjungsan_date %><input type="hidden" name="jungsan_date" value="<%= ogroup.FOneItem.Fjungsan_date %>">
			&nbsp;
			오프라인 : <%= ogroup.FOneItem.Fjungsan_date_off %><input type="hidden" name="jungsan_date_off" value="<%= ogroup.FOneItem.Fjungsan_date_off %>">
		<% end if %>
	</td>
</tr>
<% if (oJungsanDiff.FresultCount>0) then %>
<% for j=0 to oJungsanDiff.FresultCount-1 %>
<tr>
	<td colspan="4" bgcolor="CCCC22" height="25">**브랜드별 결제계좌정보 : 대표 정산계좌와 다를경우 사용 ** - <strong><%=oJungsanDiff.FItemList(j).Fpartnerid %></strong></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">거래은행</td>
	<td colspan="3" bgcolor="#FFFFFF" >
	<input type="hidden" name="jungsan_add_brand" value="<%=oJungsanDiff.FItemList(j).Fpartnerid%>">
	<% DrawBankCombo "jungsan_bank_add", oJungsanDiff.FItemList(j).Fjungsan_bank %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">계좌번호</td>
	<td colspan="3" bgcolor="#FFFFFF" >
	<% if (C_ADMIN_AUTH) or (C_MngPart) then %>
	<input type="text" class="text" name="jungsan_acctno_add" value="<%= oJungsanDiff.FItemList(j).Fjungsan_acctno %>" size="24" maxlength="32">
	<% else %>
	<input type="text" class="text_RO" name="jungsan_acctno_add" value="<%= oJungsanDiff.FItemList(j).Fjungsan_acctno %>" size="20" maxlength="32" readOnly >
	<font color="red">(경영지원팀 변경가능)</font>
	<% end if %>
	&nbsp;&nbsp; '-'은 빼고 번호만 입력해주시기 바랍니다.
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">예금주명</td>
	<td colspan="3" bgcolor="#FFFFFF" >
	<% if (C_ADMIN_AUTH) or (C_MngPart) then %>
	<input type="text" class="text" name="jungsan_acctname_add" value="<%= oJungsanDiff.FItemList(j).Fjungsan_acctname %>" size="24" maxlength="16">
	<% else %>
	<input type="text" class="text_ro" name="jungsan_acctname_add" value="<%= oJungsanDiff.FItemList(j).Fjungsan_acctname %>" size="20" maxlength="16" readOnly >
	<font color="red">(경영지원팀 변경가능)</font>
	<% end if %>

	&nbsp;&nbsp; 띄어쓰기 하지 마시기 바랍니다.
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">정산일</td>
	<td colspan="3" bgcolor="#FFFFFF" >
	온라인 : <% DrawJungsanDateCombo "jungsan_date_add", oJungsanDiff.FItemList(j).Fjungsan_date %>
	&nbsp;
	오프라인 : <% DrawJungsanDateCombo "jungsan_date_off_add", oJungsanDiff.FItemList(j).Fjungsan_date_off %>
	</td>
</tr>
<% next %>
<% end if %>

<% if C_ADMIN_AUTH or C_MngPart or C_MD_AUTH or C_PSMngPart then %>
	<tr id="regdiffacct1" style="display:none;">
		<td colspan="4" bgcolor="CCCC99" height="25">
			**브랜드별 결제계좌정보 입력 : 대표 정산계좌와 다를경우 입력 ** - <% DrawAcctDiffBrand "jungsan_add_brand","",groupid,"" %>
		</td>
	</tr>
	<tr id="regdiffacct2" style="display:none;">
		<td bgcolor="<%= adminColor("tabletop") %>">거래은행</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<% DrawBankCombo "jungsan_bank_add", "" %>
		</td>
	</tr>
	<tr id="regdiffacct3" style="display:none;">
		<td bgcolor="<%= adminColor("tabletop") %>">계좌번호</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<input type="text" class="text" name="jungsan_acctno_add" value="" size="24" maxlength="32">
			&nbsp;&nbsp; '-'은 빼고 번호만 입력해주시기 바랍니다.
		</td>
	</tr>
	<tr id="regdiffacct4" style="display:none;">
		<td bgcolor="<%= adminColor("tabletop") %>">예금주명</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<input type="text" class="text" name="jungsan_acctname_add" value="" size="24" maxlength="16">
			&nbsp;&nbsp; 띄어쓰기 하지 마시기 바랍니다.
		</td>
	</tr>
	<tr id="regdiffacct5" style="display:none;">
		<td bgcolor="<%= adminColor("tabletop") %>">정산일</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			온라인 : <% DrawJungsanDateCombo "jungsan_date_add", "" %>
			&nbsp;
			오프라인 : <% DrawJungsanDateCombo "jungsan_date_off_add", "" %>
		</td>
	</tr>
<% end if %>

<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">**담당자정보**</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">담당자명</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_name" value="<%= ogroup.FOneItem.Fmanager_name %>" size="30" maxlength="32"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">일반전화</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_phone" value="<%= ogroup.FOneItem.Fmanager_phone %>" size="16" maxlength="16" onFocusOut="phone_format(frmupche.manager_phone)"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_email" value="<%= ogroup.FOneItem.Fmanager_email %>" size="30" maxlength="64"></td>
	<td bgcolor="<%= adminColor("tabletop") %>">핸드폰</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="manager_hp" value="<%= ogroup.FOneItem.Fmanager_hp %>" size="16" maxlength="16" onFocusOut="phone_format(frmupche.manager_hp)"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">배송담당자명</td>
	<td bgcolor="#FFFFFF" colspan="3">
		<input type="button" class="button" value="브랜드별 반품주소 수정" onClick="PopUpcheReturnAddrOnly('<%= ogroup.FOneItem.FGroupId %>')">
		(브랜드별로 수정 가능합니다.)
	</td>
</tr>
<tr>
	<td width="80" bgcolor="<%= adminColor("tabletop") %>">정산담당자명</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_name" value="<%= ogroup.FOneItem.Fjungsan_name %>" size="30" maxlength="32"></td>
	<td width="80" bgcolor="<%= adminColor("tabletop") %>">일반전화</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_phone" value="<%= ogroup.FOneItem.Fjungsan_phone %>" size="16" maxlength="16" onFocusOut="phone_format(frmupche.jungsan_phone)"></td>
</tr>
<tr>
	<td width="60" bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_email" value="<%= ogroup.FOneItem.Fjungsan_email %>" size="30" maxlength="64"></td>
	<td width="60" bgcolor="<%= adminColor("tabletop") %>">핸드폰</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="jungsan_hp" value="<%= ogroup.FOneItem.Fjungsan_hp %>" size="16" maxlength="16" onFocusOut="phone_format(frmupche.jungsan_hp)"></td>
</tr>
<% if ogroupuser.FResultCount > 0 then %>
<% for i = 0 to ogroupuser.FResultCount-1 %>
<%
' 10 추가담당자
if ogroupuser.FItemList(i).fgubun="10" then
%>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">추가담당자명<%= i + 1 %><input type="hidden" name="etc_idx" value="<%= ogroupuser.FItemList(i).fidx %>"></td>
	<td bgcolor="#FFFFFF" colspan=3>
		<input type="text" class="text" name="etc_name" value="<%= ogroupuser.FItemList(i).fname %>" maxlength="32" />
		<button type="button" class="button" onClick="etc_del('<%= ogroupuser.FItemList(i).fidx %>');">삭제</button>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">핸드폰</td>
	<td bgcolor="#FFFFFF" colspan=3><input type="text" class="text" name="etc_hp" value="<%= ogroupuser.FItemList(i).fhp %>" maxlength="16" /></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
	<td bgcolor="#FFFFFF" colspan=3><input type="text" class="text" name="etc_email" value="<%= ogroupuser.FItemList(i).femail %>" maxlength="64" /></td>
</tr>
<% end if %>
<% next %>
<% end if %>
<tr id="divadd1" style="display:none;">
	<td bgcolor="<%= adminColor("tabletop") %>">추가담당자명</td>
	<td bgcolor="#FFFFFF" colspan=3><input type="text" class="text" name="addetc_name" value="" maxlength="32" /></td>
</tr>
<tr id="divadd2" style="display:none;">
	<td bgcolor="<%= adminColor("tabletop") %>">핸드폰</td>
	<td bgcolor="#FFFFFF" colspan=3><input type="text" class="text" name="addetc_hp" value="" maxlength="16" /></td>
</tr>
<tr id="divadd3" style="display:none;">
	<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
	<td bgcolor="#FFFFFF" colspan=3><input type="text" class="text" name="addetc_email" value="" maxlength="64" /></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>"><button type="button" class="button" onClick="divadd();">신규추가담당자</button></td>
	<td bgcolor="#FFFFFF" colspan=3></td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<input type="button" class="button" value="업체정보 저장" onclick="SaveUpcheInfo(frmupche);">
	</td>
</tr>
</table>
</form>
<form name="frmetcedit" method="post" action="/admin/member/doupcheedit.asp">
<input type="hidden" name="mode" value="etc_del" />
<input type="hidden" name="etc_idx" value="" />
<input type="hidden" name="groupid" value="<%= ogroup.FOneItem.FGroupId %>" />
</form>
<iframe src="" name="icheckframe" width="0" height="0" frameborder="1" marginwidth="0" marginheight="0" topmargin="0" scrolling="no"></iframe>

<%
set oJungsanDiff = Nothing
set ogroup = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
