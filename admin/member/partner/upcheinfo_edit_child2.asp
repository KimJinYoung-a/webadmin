<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 업체정보등록/변경
' History : 2015.05.27 강준구 생성
'			2021.12.06 한용민 수정(권한수정)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/admin/member/partner/partnerCls.asp"-->

<%
dim ogroup,i, vTIdx, groupid, vGubun, vCompNOchgOX, vSocNo, groupid_old, arrFileList, intLoop
	vTIdx 			= request("tidx")
	groupid 		= request("groupid")
	groupid_old		= request("groupid_old")
	vGubun 			= Request("gb")
	vCompNOchgOX 	= Request("compnochgox")
	vSocNo			= Request("socno")

If groupid_old = "" Then
	groupid_old = groupid
End IF

If vTIdx = "" Then
	set ogroup = new CPartnerGroup
	ogroup.FRectGroupid = groupid
	ogroup.GetOneGroupInfo
Else
	set ogroup = new cPartnerInfoReq
	ogroup.Ftidx = vTIdx
	ogroup.Fgroupid = groupid
	ogroup.fRequestDetail
	
	arrFileList = ogroup.fnGetFileList
End If

%>

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

function SaveUpcheInfo(frm,gubun){
	<% If vGubun = "" Then %>
	alert("변경할 내용에 맞는\n왼쪽편 내용페이지에 변경 버튼을 클릭하세요.");
	return;
	<% End If %>

    var psocno =frm.psocno.value;
    
	    if (frm.uid.value.length<1){
			alert('입점브랜드ID를 왼쪽편 내용페이지에서 선택하세요.');
			return;
		}
    <% If vGubun = "companyreginfo" Then %>
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
	<% End If %>
	<% If vGubun = "companyreginfo" OR vGubun = "jungsandate" Then %>
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
	<% End If %>
 
	var ret = confirm('이 정보로 변경하겠습니까?');

	if (ret){
	
		if(gubun == "temp")
		{
			if(psocno != frm.company_no.value)
			{
			
				frm.groupid.value = "";
			}
			frm.action = "upcheinfo_edit_proc.asp";
		}
		else if(gubun == "real")
		{
			//alert("실제적용 프로세스 작업중");
			//return;
			frm.action = "upcheinfo_edit_real_proc.asp";
		}

		frm.submit();
	}
}

function statusChange(a){
	var message = "";
	if(a == "0"){
		message = "신청서를 삭제하시겠습니까?";
	}else if(a == "1"){
		message = "신청전환으로 변경하시겠습니까?";
	}else if(a == "2"){
		message = "작업중전환으로 변경하시겠습니까?";
	}

	if(confirm(message) == true) {
		frmupche.status.value = a;
		frmupche.action = "upcheinfo_edit_proc.asp";
		frmupche.submit();
	} else {
		return false;
	}
}

function AddProc(mode){
	alert('등록가능한 사업자번호입니다.');
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


	var popwin = window.open("/admin/member/partner/popupchereturnaddronly.asp?groupid=" + groupid,"popupchereturnaddronly","width=900 height=580 scrollbars=yes resizable=yes");
	popwin.focus();
}

function fileupload(){
	<% If vGubun = "" Then %>
	alert("변경할 내용에 맞는\n왼쪽편 내용페이지에 변경 버튼을 클릭하세요.");
	return;
	<% End If %>
	window.open('popUpload.asp','worker','width=420,height=200,scrollbars=yes');
}

function filedownload(idx){
	filefrm.file_idx.value = idx;
	filefrm.submit();
}

function clearRow(tdObj) {
	if(confirm("선택하신 파일을 삭제하시겠습니까?") == true) {
		var tblObj = tdObj.parentNode.parentNode.parentNode;
		var trIdx = tdObj.parentNode.parentNode.rowIndex;
	
		tblObj.deleteRow(trIdx);
	} else {
		return false;
	}
}
</script>

<form name="frmupche" method="post" style="margin:0px;">
<input type="hidden" name="tidx" value="<%=vTIdx%>">
<input type="hidden" name="gubun" value="<%=vGubun%>">
<input type="hidden" name="mode" value="groupedit">
<input type="hidden" name="psocno" value="<%= ogroup.FOneItem.getDecCompNo %>">
<table width="600" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">&nbsp;* <b><font size="2">변경사항 신청서</font></b><%=CHKIIF(vTIdx="",""," (신청일 : " & ogroup.FOneItem.Fregdate & ")")%></td>
</tr>
<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan=4><b>1.업체관련정보</b></td>
</tr>
<tr height="25">
	<td width="120" bgcolor="<%= adminColor("tabletop") %>">업체코드</td>
	<td bgcolor="#FFFFFF" width="200">
		<input type="text" class="text" name="groupid" value="<%= ogroup.FOneItem.FGroupId %>" size="7" maxlength="5" style="background-color:#EEEEEE;" readonly>
		<input type="hidden" name="groupid_old" value="<%= CHKIIF(vTIdx="",groupid_old,ogroup.FOneItem.Fgroupid_old) %>">
	</td>
	<td width="90" bgcolor="<%= adminColor("tabletop") %>">업체명</td>
	<td bgcolor="#FFFFFF" width="200">
		<%= ogroup.FOneItem.FCompany_name %>
	</td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("tabletop") %>">입점브랜드ID</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input type="text" name="uid" value="<%=CHKIIF(vTIdx="","",ogroup.FOneItem.getBrandList)%>" size="60" readonly>
		<input type="hidden" name="old_uid" value="<%=ogroup.FOneItem.getBrandList%>">
	</td>
</tr>

<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">**사업자등록정보**</td>
</tr>

<tr height="25">
	<td bgcolor="<%= adminColor("tabletop") %>">회사명(상호)</td>
	<td bgcolor="#FFFFFF">
		<% IF vGubun = "companyreginfo" Then %>
			<input type="text" class="text" name="company_name" value="<%= ogroup.FOneItem.FCompany_name %>" size="27" maxlength="32">
		<% Else  %>
			<%= ogroup.FOneItem.FCompany_name %><input type="hidden" name="company_name" value="<%= ogroup.FOneItem.FCompany_name %>">
		<% End If %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">대표자</td>
	<td bgcolor="#FFFFFF">
		<% IF vGubun = "companyreginfo" Then %>
			<input type="text" class="text" name="ceoname" value="<%= ogroup.FOneItem.Fceoname %>" size="16" maxlength="16">
		<% Else  %>
			<%= ogroup.FOneItem.Fceoname %><input type="hidden" name="ceoname" value="<%= ogroup.FOneItem.Fceoname %>">
		<% End If %>
	</td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("tabletop") %>">사업자번호</td>
	<td bgcolor="#FFFFFF">
		<% IF vGubun = "companyreginfo" Then %>
			<input type="text" class="text" name="company_no" value="<%= CHKIIF(ogroup.FOneItem.getDecCompNo="",vSocNo,ogroup.FOneItem.getDecCompNo) %>" size="16" maxlength="20" style="background-color:#EEEEEE;" readonly>
		<% Else  %>
			<%= ogroup.FOneItem.getDecCompNo %><input type="hidden" name="company_no" value="<%= ogroup.FOneItem.getDecCompNo %>">
		<% End If %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">과세구분</td>
	<td bgcolor="#FFFFFF">
		<% IF vGubun = "companyreginfo" Then %>
		<select name="jungsan_gubun" class="select">
			<option value="일반과세" <% if ogroup.FOneItem.Fjungsan_gubun="일반과세" then response.write "selected" %> >일반과세</option>
			<option value="간이과세" <% if ogroup.FOneItem.Fjungsan_gubun="간이과세" then response.write "selected" %> >간이과세</option>
			<option value="원천징수" <% if ogroup.FOneItem.Fjungsan_gubun="원천징수" then response.write "selected" %> >원천징수</option>
			<option value="면세" <% if ogroup.FOneItem.Fjungsan_gubun="면세" then response.write "selected" %> >면세</option>
			<option value="영세(해외)" <% if ogroup.FOneItem.Fjungsan_gubun="영세(해외)" then response.write "selected" %> >영세(해외)</option>
		</select>
		<% Else  %>
			<%= ogroup.FOneItem.Fjungsan_gubun %><input type="hidden" name="jungsan_gubun" value="<%= ogroup.FOneItem.Fjungsan_gubun %>">
		<% End If %>
	</td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("tabletop") %>">사업장소재지</td>
	<td colspan="3" bgcolor="#FFFFFF" >
		<% IF vGubun = "companyreginfo" Then %>
			<input type="text" class="text" name="company_zipcode" value="<%= ogroup.FOneItem.Fcompany_zipcode %>" size="7" maxlength="7">
			<input type="button" class="button" value="검색" onClick="FnFindZipNew('frmupche','C')">
			<input type="button" class="button" value="검색(구)" onClick="TnFindZipNew('frmupche','C')">
			<% '<input type="button" class="button" value="검색(구)" onClick="popZip('s');"> %>
			<br>
			<input type="text" class="text" name="company_address" value="<%= ogroup.FOneItem.Fcompany_address %>" size="27" maxlength="64">&nbsp;
			<input type="text" class="text" name="company_address2" value="<%= ogroup.FOneItem.Fcompany_address2 %>" size="46" maxlength="64">
		<% Else  %>
			[<%= ogroup.FOneItem.Fcompany_zipcode %>]<%= ogroup.FOneItem.Fcompany_address %> <%= ogroup.FOneItem.Fcompany_address2 %>
			<input type="hidden" name="company_zipcode" value="<%= ogroup.FOneItem.Fcompany_zipcode %>">
			<input type="hidden" name="company_address" value="<%= ogroup.FOneItem.Fcompany_address %>">
			<input type="hidden" name="company_address2" value="<%= ogroup.FOneItem.Fcompany_address2 %>">
		<% End If %>
	</td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("tabletop") %>">업태</td>
	<td bgcolor="#FFFFFF">
		<% IF vGubun = "companyreginfo" Then %>
			<input type="text" class="text" name="company_uptae" value="<%= ogroup.FOneItem.Fcompany_uptae %>" size="27" maxlength="32">
		<% Else  %>
			<%= ogroup.FOneItem.Fcompany_uptae %><input type="hidden" name="company_uptae" value="<%= ogroup.FOneItem.Fcompany_uptae %>">
		<% End If %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">업종</td>
	<td bgcolor="#FFFFFF">
		<% IF vGubun = "companyreginfo" Then %>
			<input type="text" class="text" name="company_upjong" value="<%= ogroup.FOneItem.Fcompany_upjong %>" size="27" maxlength="32">
		<% Else  %>
			<%= ogroup.FOneItem.Fcompany_upjong %><input type="hidden" name="company_upjong" value="<%= ogroup.FOneItem.Fcompany_upjong %>">
		<% End If %>
	</td>
</tr>

<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">**업체기본정보**</td>
</tr>

<tr height="25">
	<td bgcolor="<%= adminColor("tabletop") %>">대표전화</td>
	<td bgcolor="#FFFFFF">
		<% IF vGubun = "companyreginfo" Then %>
			<input type="text" class="text" name="company_tel" value="<%= ogroup.FOneItem.Fcompany_tel %>" size="16" maxlength="16" onFocusOut="phone_format(frmupche.company_tel)">
		<% Else  %>
			<%= ogroup.FOneItem.Fcompany_tel %><input type="hidden" name="company_tel" value="<%= ogroup.FOneItem.Fcompany_tel %>">
		<% End If %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">팩스</td>
	<td bgcolor="#FFFFFF">
		<% IF vGubun = "companyreginfo" Then %>
			<input type="text" class="text" name="company_fax" value="<%= ogroup.FOneItem.Fcompany_fax %>" size="16" maxlength="16" onFocusOut="phone_format(frmupche.company_fax)">
		<% Else  %>
			<%= ogroup.FOneItem.Fcompany_fax %><input type="hidden" name="company_fax" value="<%= ogroup.FOneItem.Fcompany_fax %>">
		<% End If %>
	</td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("tabletop") %>">사무실 주소</td>
	<td colspan="3" bgcolor="#FFFFFF" >
		<% IF vGubun = "companyreginfo" Then %>
			<input type="text" class="text" name="return_zipcode" value="<%= ogroup.FOneItem.Freturn_zipcode %>" size="7" maxlength="7">
			<input type="button" class="button" value="검색" onClick="FnFindZipNew('frmupche','D')">
			<input type="button" class="button" value="검색(구)" onClick="TnFindZipNew('frmupche','D')">
			<% '<input type="button" class="button" value="검색(구)" onClick="popZip('m');"> %>
			<input type=checkbox name=samezip onclick="SameReturnAddr(this.checked)">상동
			<br>
			<input type="text" class="text" name="return_address" value="<%= ogroup.FOneItem.Freturn_address %>" size="30" maxlength="64">&nbsp;
			<input type="text" class="text" name="return_address2" value="<%= ogroup.FOneItem.Freturn_address2 %>" size="46" maxlength="64">
		<% Else  %>
			[<%= ogroup.FOneItem.Freturn_zipcode %>]<%= ogroup.FOneItem.Freturn_address %> <%= ogroup.FOneItem.Freturn_address2 %>
			<input type="hidden" name="return_zipcode" value="<%= ogroup.FOneItem.Freturn_zipcode %>">
			<input type="hidden" name="return_address" value="<%= ogroup.FOneItem.Freturn_address %>">
			<input type="hidden" name="return_address2" value="<%= ogroup.FOneItem.Freturn_address2 %>">
		<% End If %>
	</td>
</tr>
<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">**결제계좌정보**</td>
</tr>

<tr height="25">
	<td bgcolor="<%= adminColor("tabletop") %>">거래은행</td>
	<td colspan="3" bgcolor="#FFFFFF" >
		<% IF vGubun = "bankinfo" OR vGubun = "companyreginfo" Then %>
			<% DrawBankCombo "jungsan_bank", ogroup.FOneItem.Fjungsan_bank %>
		<% Else  %>
			<%=ogroup.FOneItem.Fjungsan_bank%><input type="hidden" name="jungsan_bank" value="<%= ogroup.FOneItem.Fjungsan_bank %>">
		<% End If %>
	</td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("tabletop") %>">계좌번호</td>
	<td colspan="3" bgcolor="#FFFFFF" >
		<% IF vGubun = "bankinfo" OR vGubun = "companyreginfo" Then %>
			<input type="text" class="text" name="jungsan_acctno" value="<%= ogroup.FOneItem.Fjungsan_acctno %>" size="24" maxlength="32">
			&nbsp;&nbsp; '-'은 빼고 번호만 입력해주시기 바랍니다.
		<% Else  %>
			<%=ogroup.FOneItem.Fjungsan_acctno%><input type="hidden" name="jungsan_acctno" value="<%= ogroup.FOneItem.Fjungsan_acctno %>">
		<% End If %>
	</td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("tabletop") %>">예금주명</td>
	<td colspan="3" bgcolor="#FFFFFF" >
		<% IF vGubun = "bankinfo" OR vGubun = "companyreginfo" Then %>
			<input type="text" class="text" name="jungsan_acctname" value="<%= ogroup.FOneItem.Fjungsan_acctname %>" size="24" maxlength="16">
			&nbsp;&nbsp; 띄어쓰기 하지 마시기 바랍니다.
		<% Else  %>
			<%=ogroup.FOneItem.Fjungsan_acctname%><input type="hidden" name="jungsan_acctname" value="<%= ogroup.FOneItem.Fjungsan_acctname %>">
		<% End If %>
	</td>
</tr>
<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">
		<table width="100%" cellspacing="0" cellpadding="0" border="0" class="a">
		<tr>
			<td>**정산일정보**</td>
		</tr>
		</table>
	</td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("tabletop") %>">정산일</td>
	<td colspan="3" bgcolor="#FFFFFF" >
		<% IF vGubun = "companyreginfo" OR vGubun = "jungsandate" Then %>
			온라인 : <% DrawJungsanDateCombo "jungsan_date", ogroup.FOneItem.Fjungsan_date %>
			&nbsp;
			오프라인 : <% DrawJungsanDateCombo "jungsan_date_off", ogroup.FOneItem.Fjungsan_date_off %>
		<% Else  %>
			온라인 : <%=ogroup.FOneItem.Fjungsan_date%>&nbsp;오프라인 : <%=ogroup.FOneItem.Fjungsan_date_off%>
			<input type="hidden" name="jungsan_date" value="<%= ogroup.FOneItem.Fjungsan_date %>">
			<input type="hidden" name="jungsan_date_off" value="<%= ogroup.FOneItem.Fjungsan_date_off %>">
		<% End If %>
	</td>
</tr>
<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">**담당자정보**</td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("tabletop") %>">담당자명</td>
	<td bgcolor="#FFFFFF">
		<% IF vGubun = "companyreginfo" Then %>
			<input type="text" class="text" name="manager_name" value="<%= ogroup.FOneItem.Fmanager_name %>" size="30" maxlength="32">
		<% Else  %>
			<%= ogroup.FOneItem.Fmanager_name %><input type="hidden" name="manager_name" value="<%= ogroup.FOneItem.Fmanager_name %>">
		<% End If %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">일반전화</td>
	<td bgcolor="#FFFFFF">
		<% IF vGubun = "companyreginfo" Then %>
			<input type="text" class="text" name="manager_phone" value="<%= ogroup.FOneItem.Fmanager_phone %>" size="16" maxlength="16" onFocusOut="phone_format(frmupche.manager_phone)">
		<% Else  %>
			<%= ogroup.FOneItem.Fmanager_phone %><input type="hidden" name="manager_phone" value="<%= ogroup.FOneItem.Fmanager_phone %>">
		<% End If %>
	</td>
</tr>
<tr height="25">
	<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
	<td bgcolor="#FFFFFF">
		<% IF vGubun = "companyreginfo" Then %>
			<input type="text" class="text" name="manager_email" value="<%= ogroup.FOneItem.Fmanager_email %>" size="30" maxlength="64">
		<% Else  %>
			<%= ogroup.FOneItem.Fmanager_email %><input type="hidden" name="manager_email" value="<%= ogroup.FOneItem.Fmanager_email %>">
		<% End If %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>">핸드폰</td>
	<td bgcolor="#FFFFFF">
		<% IF vGubun = "companyreginfo" Then %>
			<input type="text" class="text" name="manager_hp" value="<%= ogroup.FOneItem.Fmanager_hp %>" size="16" maxlength="16" onFocusOut="phone_format(frmupche.manager_hp)">
		<% Else  %>
			<%= ogroup.FOneItem.Fmanager_hp %><input type="hidden" name="manager_hp" value="<%= ogroup.FOneItem.Fmanager_hp %>">
		<% End If %>
	</td>
</tr>

<tr height="25">
	<td width="80" bgcolor="<%= adminColor("tabletop") %>">정산담당자명</td>
	<td bgcolor="#FFFFFF">
		<% IF vGubun = "companyreginfo" Then %>
			<input type="text" class="text" name="jungsan_name" value="<%= ogroup.FOneItem.Fjungsan_name %>" size="30" maxlength="32">
		<% Else  %>
			<%= ogroup.FOneItem.Fjungsan_name %><input type="hidden" name="jungsan_name" value="<%= ogroup.FOneItem.Fjungsan_name %>">
		<% End If %>
	</td>
	<td width="80" bgcolor="<%= adminColor("tabletop") %>">일반전화</td>
	<td bgcolor="#FFFFFF">
		<% IF vGubun = "companyreginfo" Then %>
			<input type="text" class="text" name="jungsan_phone" value="<%= ogroup.FOneItem.Fjungsan_phone %>" size="16" maxlength="16" onFocusOut="phone_format(frmupche.jungsan_phone)">
		<% Else  %>
			<%= ogroup.FOneItem.Fjungsan_phone %><input type="hidden" name="jungsan_phone" value="<%= ogroup.FOneItem.Fjungsan_phone %>">
		<% End If %>
	</td>
</tr>
<tr height="25">
	<td width="60" bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
	<td bgcolor="#FFFFFF">
		<% IF vGubun = "companyreginfo" Then %>
			<input type="text" class="text" name="jungsan_email" value="<%= ogroup.FOneItem.Fjungsan_email %>" size="30" maxlength="64">
		<% Else  %>
			<%= ogroup.FOneItem.Fjungsan_email %><input type="hidden" name="jungsan_email" value="<%= ogroup.FOneItem.Fjungsan_email %>">
		<% End If %>
	</td>
	<td width="60" bgcolor="<%= adminColor("tabletop") %>">핸드폰</td>
	<td bgcolor="#FFFFFF">
		<% IF vGubun = "companyreginfo" Then %>
			<input type="text" class="text" name="jungsan_hp" value="<%= ogroup.FOneItem.Fjungsan_hp %>" size="16" maxlength="16" onFocusOut="phone_format(frmupche.jungsan_hp)">
		<% Else  %>
			<%= ogroup.FOneItem.Fjungsan_hp %><input type="hidden" name="jungsan_hp" value="<%= ogroup.FOneItem.Fjungsan_hp %>">
		<% End If %>
	</td>
</tr>
<tr>
	<td colspan="4" bgcolor="#FFFFFF" height="25">
		**첨부파일**&nbsp;&nbsp;&nbsp;<input type="button" value="파일업로드" onClick="fileupload()" class="button">
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td width="100%" style="padding:3 0 3 10">
				<table cellpadding="0" cellspacing="0" vorder="0" id="fileup">
				<%
				IF isArray(arrFileList) THEN
					For intLoop =0 To UBound(arrFileList,2)
				%>
					<tr>
						<td>
							<input type='hidden' name='info_file' value='<%=arrFileList(1,intLoop)%>'>
							<input type='hidden' name='info_realfile' value='<%=arrFileList(2,intLoop)%>'>
							<img src='https://fiximage.10x10.co.kr/web2009/common/cmt_del.gif' border='0' style='cursor:pointer' onClick='clearRow(this)'>
							<span id="<%=intLoop%>" class="a" onClick="filedownload(<%=arrFileList(0,intLoop)%>)" style="cursor:pointer"><%=Split(Replace(arrFileList(1,intLoop),"http://",""),"/")(4)%></span>
						</td>
					</tr>
				<%
					Next
					Response.Write "<input type='hidden' name='isfile' value='o'>"
				Else
				%>
					<tr>
						<td>
						</td>
					</tr>
				<% End If %>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td colspan="4" bgcolor="#FFFFFF" style="padding:10 0 10 7;">
		* COMMENT<br><textarea name="comment" rows="4" cols="<%=CHKIIF(InStr(UCase(cstr(request.ServerVariables("HTTP_USER_AGENT"))),"MSIE"),"80","68")%>"><%=ogroup.FOneItem.FComment%></textarea>
		<% If vTIdx = "" Then %>
		<br><br><img src="/images/icon_save.gif" style="cursor:pointer;" onclick="SaveUpcheInfo(frmupche,'temp');" title="업체정보저장">
		<% Else %>
		<br><br>
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td>최종수정일 : <%=ogroup.FOneItem.Flastupdate%>, 최종확인인:<%=ogroup.FOneItem.Fusername%></td>
			<td style="padding-right:10px;" align="right">
				현재상태 : <b><font color="blue"><%=RequestStateName(ogroup.FOneItem.Fstatus)%></font></b>
				<input type="hidden" name="status" value="<%=ogroup.FOneItem.Fstatus%>">
			</td>
		</tr>
		<tr>
			<td style="padding-top:3px;">
				<%
					If ogroup.FOneItem.Fstatus = "1" OR ogroup.FOneItem.Fstatus = "2" Then
						If ogroup.FOneItem.Fstatus = "3" Then
							if C_MngPart or C_ADMIN_AUTH then
				%>
							<img src="/images/coop_modify.gif" style="cursor:pointer;" onclick="SaveUpcheInfo(frmupche,'temp');" title="신청서내용저장">
				<%
							End If
						Else
				%>
							<img src="/images/coop_modify.gif" style="cursor:pointer;" onclick="SaveUpcheInfo(frmupche,'temp');" title="신청서내용저장">
				<%
							If C_MngPart or C_ADMIN_AUTH OR ogroup.FOneItem.Freguserid = session("ssBctId") Then
				%>
							&nbsp;<img src="/images/icon_delete.gif" style="cursor:pointer;" onclick="statusChange('0');" title="삭제">
				<%
							End IF
						End If
					Else
						If ogroup.FOneItem.Fstatus = "3" AND (C_MngPart or C_ADMIN_AUTH) Then
				%>
							<img src="/images/coop_modify.gif" style="cursor:pointer;" onclick="SaveUpcheInfo(frmupche,'temp');" title="신청서내용저장">&nbsp;&nbsp;&nbsp;
				<%
						End If
						If ogroup.FOneItem.Fstatus = "0" Then
							Response.Write "※ <b>이미 처리가 삭제처리된 신청서 입니다.</b>"
						ElseIf ogroup.FOneItem.Fstatus = "3" Then
							Response.Write "※ <b>이미 처리가 완료 및 실제 적용된 신청서 입니다.</b>"
						End If
					End If
				%>
			</td>
			<td style="padding:3px 10px 0 0;" align="right">
			<% if C_MngPart or C_ADMIN_AUTH then %>
				<% If ogroup.FOneItem.Fstatus = "1" OR ogroup.FOneItem.Fstatus = "2" Then %>
					&nbsp;<img src="/images/coop_req.gif" style="cursor:pointer;" onclick="statusChange('1');" title="신청전환">
					&nbsp;<img src="/images/coop_jak.gif" style="cursor:pointer;" onclick="statusChange('2');" title="작업중전환">
					&nbsp;<img src="/images/coop_won.gif" style="cursor:pointer;" onclick="SaveUpcheInfo(frmupche,'real');" title="완료전환">
				<% ElseIf ogroup.FOneItem.Fstatus = "0" Then %>
					&nbsp;<img src="/images/coop_req.gif" style="cursor:pointer;" onclick="statusChange('1');" title="신청전환">
				<% End If %>
			<% ElseIf ogroup.FOneItem.Fstatus = "2" Then %>
				&nbsp;<img src="/images/coop_req.gif" style="cursor:pointer;" onclick="statusChange('1');" title="신청전환">
			<% End If %>
			</td>
		</tr>
		</table>
		<% End If %>
	</td>
</tr>
</table>
</form>
<iframe src="" name="icheckframe" width="0" height="0" frameborder="0" marginwidth="0" marginheight="0" topmargin="0" scrolling="no"></iframe>

<form name="filefrm" method="post" action="<%=uploadImgUrl%>/linkweb/partner_info/partner_info_download.asp" target="fileiframe">
<input type="hidden" name="tidx" value="<%=vTIdx%>">
<input type="hidden" name="file_idx" value="">
</form>
<iframe src="" width="0" height="0" name="fileiframe" frameborder="0" width="0" height="0"></iframe>

<%
set ogroup = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->