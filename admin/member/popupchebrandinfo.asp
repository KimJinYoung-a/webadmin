<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->

<%

dim ogroup,opartner,i
dim designer
dim groupid

designer = request("designer")

set opartner = new CPartnerUser
opartner.FRectDesignerID = designer
opartner.GetOnePartnerNUser


set ogroup = new CPartnerGroup
ogroup.FRectGroupid = opartner.FOneItem.FGroupid
ogroup.GetOneGroupInfo


dim ooffontract
set ooffontract = new COffContractInfo
ooffontract.FRectDesignerID = designer
ooffontract.GetPartnerOffContractInfo


dim returnsongjangStr

returnsongjangStr = returnsongjangStr + "10x10" & chr(9)
returnsongjangStr = returnsongjangStr + "(주)텐바이텐" & chr(9)
returnsongjangStr = returnsongjangStr + ogroup.FOneItem.FCompany_name  & chr(9)
returnsongjangStr = returnsongjangStr + ogroup.FOneItem.Fdeliver_phone  & chr(9)
returnsongjangStr = returnsongjangStr + ogroup.FOneItem.Fdeliver_hp  & chr(9)
returnsongjangStr = returnsongjangStr + replace(ogroup.FOneItem.Freturn_zipcode,"-","") & chr(9)
returnsongjangStr = returnsongjangStr + ogroup.FOneItem.Freturn_address  & chr(9)
returnsongjangStr = returnsongjangStr + ogroup.FOneItem.Freturn_address2  & chr(9)
returnsongjangStr = returnsongjangStr + "10x10 반품" & chr(9)
returnsongjangStr = returnsongjangStr + "반품상품" & chr(9)
returnsongjangStr = returnsongjangStr + opartner.FOneItem.FID
%>

<!-- returnsongjangStr = FormatDate(now(),"0000.00.00 00:00:00")
returnsongjangStr = Replace(returnsongjangStr,".","")
returnsongjangStr = Replace(returnsongjangStr,":","")
returnsongjangStr = Replace(returnsongjangStr," ","")
returnsongjangStr = returnsongjangStr & chr(9)
-->

<script language='javascript'>
function copyComp(comp) {
	comp.focus()
	comp.select()
	therange=comp.createTextRange()
	therange.execCommand("Copy")
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

function SaveBrandInfo(frm){
	if (frm.prtidx.value.length<1){
		alert('랙 번호를 입력하세요. - [기본값 9999]');
		frm.prtidx.focus();
		return;
	}

	if (frm.password.value.length<1){
		alert('브랜드 패스워드를 입력하세요.');
		frm.password.focus();
		return;
	}

	if (frm.socname_kor.value.length<1){
		alert('스트리트명(한글)을 입력하세요.');
		frm.socname_kor.focus();
		return;
	}

	if (frm.socname.value.length<1){
		alert('스트리트명(영문)을 입력하세요.');
		frm.socname.focus();
		return;
	}

	if ((!frm.isusing[0].checked)&&(!frm.isusing[1].checked)){
		alert('사용여부를 선택하세요.');
		frm.isusing[0].focus();
		return;
	}

	if ((!frm.isextusing[0].checked)&&(!frm.isextusing[1].checked)){
		alert('제휴몰 사용여부를 선택하세요.');
		frm.isextusing[0].focus();
		return;
	}

	if ((!frm.streetusing[0].checked)&&(!frm.streetusing[1].checked)){
		alert('스트리트 사용여부를 선택하세요.');
		frm.streetusing[0].focus();
		return;
	}

	if ((!frm.extstreetusing[0].checked)&&(!frm.extstreetusing[1].checked)){
		alert('제휴몰 스트리트 사용여부를 선택하세요.');
		frm.extstreetusing[0].focus();
		return;
	}

	if ((!frm.specialbrand[0].checked)&&(!frm.specialbrand[1].checked)){
		alert('커뮤니티 사용여부를 선택하세요.');
		frm.specialbrand[0].focus();
		return;
	}

	if (frm.userdiv.value.length<1){
		alert('브랜드 구분을 선택하세요.');
		frm.userdiv.focus();
		return;
	}

	if (frm.maeipdiv.value.length<1){
		alert('매입 구분을 선택하세요.');
		frm.maeipdiv.focus();
		return;
	}

	if (!IsDouble(frm.defaultmargine.value)){
		alert('기본마진을 입력하세요. - 실수만 가능합니다.');
		frm.defaultmargine.focus();
		return;
	}


	var ret = confirm('브랜드 정보를 저장 하시겠습니까?');

	if (ret){
		frm.submit();
	}
}

function SaveUpcheInfo(frm){
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


	if (frm.groupid.value.length<1){
		var ret = confirm('업체 정보를 저장 하시겠습니까?');
	}else{
		var ret = confirm('같은 그룹코드에 있는 기존 업체 정보도 수정됩니다. 저장 하시겠습니까?');
	}

	if (ret){
		frm.submit();
	}
}

function ModiInfo(frm){
	var ret = confirm('저장 하시겠습니까?');

	if (ret){
		//frm.submit();
	}

}

function PopUpcheList(frmname){
	var popwin = window.open("/admin/lib/popupchelist.asp?frmname=" + frmname,"popupchelist","width=700 height=540 scrollbars=yes resizable=yes");
	popwin.focus();
}
</script>
<form name=frmbuf>
<input type=hidden name=company_name value="<%= opartner.FOneItem.FCompany_name %>">
<input type=hidden name=ceoname value="<%= opartner.FOneItem.Fceoname %>">
<input type=hidden name=company_no value="<%= opartner.FOneItem.Fcompany_no %>">
<input type=hidden name=jungsan_gubun value="<%= opartner.FOneItem.Fjungsan_gubun %>">
<input type=hidden name=company_zipcode value="<%= opartner.FOneItem.Fzipcode %>">
<input type=hidden name=company_address value="<%= opartner.FOneItem.Faddress %>">
<input type=hidden name=company_address2 value="<%= opartner.FOneItem.Fmanager_address %>">
<input type=hidden name=company_uptae value="<%= opartner.FOneItem.Fcompany_uptae %>">
<input type=hidden name=company_upjong value="<%= opartner.FOneItem.Fcompany_upjong %>">
<input type=hidden name=company_tel value="<%= opartner.FOneItem.Ftel %>">
<input type=hidden name=company_fax value="<%= opartner.FOneItem.Ffax %>">

<input type=hidden name=jungsan_bank value="<%= opartner.FOneItem.Fjungsan_bank %>">
<input type=hidden name=jungsan_acctno value="<%= opartner.FOneItem.Fjungsan_acctno %>">
<input type=hidden name=jungsan_acctname value="<%= opartner.FOneItem.Fjungsan_acctname %>">
<input type=hidden name=manager_name value="<%= opartner.FOneItem.Fmanager_name %>">
<input type=hidden name=manager_phone value="<%= opartner.FOneItem.Fmanager_phone %>">
<input type=hidden name=manager_email value="<%= opartner.FOneItem.Femail %>">
<input type=hidden name=manager_hp value="<%= opartner.FOneItem.Fmanager_hp %>">

<input type=hidden name=deliver_name value="<%= opartner.FOneItem.Fdeliver_name %>">
<input type=hidden name=deliver_phone value="<%= opartner.FOneItem.Fdeliver_phone %>">
<input type=hidden name=deliver_email value="<%= opartner.FOneItem.Fdeliver_email %>">
<input type=hidden name=deliver_hp value="<%= opartner.FOneItem.Fdeliver_hp %>">

<input type=hidden name=jungsan_name value="<%= opartner.FOneItem.Fjungsan_name %>">
<input type=hidden name=jungsan_phone value="<%= opartner.FOneItem.Fjungsan_phone %>">
<input type=hidden name=jungsan_email value="<%= opartner.FOneItem.Fjungsan_email %>">
<input type=hidden name=jungsan_hp value="<%= opartner.FOneItem.Fjungsan_hp %>">


</form>

<table width="600" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
	        	브랜드 ID : <input type="text" name="designer" value="<%= designer %>" Maxlength="32" size="16">
	        </td>
	        <td valign="top" align="right" bgcolor="F4F4F4">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>


<table width="600" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#000000>
<form name="frmupche" method="post" action="doupcheedit.asp">
<input type="hidden" name="mode" value="groupedit">
<input type="hidden" name="uid" value="<%= designer %>">
	<tr bgcolor="#DDDDFF">
		<td colspan=4><b>1.업체관련정보</b></td>
	</tr>
	<tr >
		<td width="100" bgcolor="#DDDDFF">업체코드</td>
		<td bgcolor="#FFFFFF" width="200">
		<input type="text" name="groupid" value="<%= ogroup.FOneItem.FGroupId %>" size="7" maxlength="5" style="background-color:#EEEEEE;" readonly>
		<% if (C_ADMIN_AUTH=true) then %>
		<input type="button" value="업체선택" onClick="PopUpcheSelect('frmupche');">
		<% end if %>
		</td>
		<td width="100" bgcolor="#DDDDFF">업체명</td>
		<td bgcolor="#FFFFFF" width="200">
		<%= ogroup.FOneItem.FCompany_name %>
		</td>
	</tr>
	<tr >
		<td width="100" bgcolor="#DDDDFF">입점브랜드ID</td>
		<td colspan="3" bgcolor="#FFFFFF"><%= ogroup.FOneItem.getBrandList %></td>
	</tr>

	<tr>
		<td height="30" colspan="4" bgcolor="#FFFFFF" align="center"><input class="button" type="button" value="업체정보 보기" onclick="PopUpcheInfoEdit('<%= ogroup.FOneItem.FGroupId %>');"></td>
	</tr>
</form>
</table>

<br>
<br>
<table width="600" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#000000>
<form name="frmbrand" method="post" action="doupcheedit.asp">
<input type="hidden" name="uid" value="<%= opartner.FOneItem.FID %>">
<input type="hidden" name="mode" value="brandedit">
<tr>
	<td bgcolor="#FFDDDD" colspan=4><b>2.브랜드관련정보</b></td>
</tr>
<tr>
	<td width="100" bgcolor="#FFDDDD">회사명</td>
	<td bgcolor="#FFFFFF"><%= opartner.FOneItem.FCompany_name %></td>
	<td width="100" bgcolor="#FFDDDD" >브랜드ID</td>
	<td bgcolor="#FFFFFF"><%= opartner.FOneItem.FID %></td>
</tr>

<tr>
	<td bgcolor="#FFDDDD">스트리트명<br>(한글)</td>
	<td bgcolor="#FFFFFF">
		<%= opartner.FOneItem.Fsocname_kor %>
	</td>
	<td bgcolor="#FFDDDD">스트리트명<br>(영문)</td>
	<td bgcolor="#FFFFFF">
		<%= opartner.FOneItem.Fsocname %>
	</td>
</tr>

<tr>
	<td bgcolor="#FFDDDD" >어드민 오픈여부</td>
	<td bgcolor="#FFFFFF" colspan=2>
		<% if opartner.FOneItem.Fpartnerusing="Y" then %>
		사용함
		<% else %>
		<font color=red>사용안함</font>
		<% end if %>
	</td>
	<td bgcolor="#FFFFFF" align="right"><a href="javascript:PopBrandAdminUsingChange('<%= opartner.FOneItem.FID %>');"><img src="/images/icon_modify.gif" border="0"></td>
</tr>
<tr>
	<td height="30" colspan="4" bgcolor="#FFFFFF" align="center"><input class="button" type="button" value="브랜드정보 보기" onclick="PopBrandInfoEdit('<%= opartner.FOneItem.FID %>');"></td>
</tr>

<!--
<% if ogroup.FOneItem.FGroupId<>"" then %>
<tr>
	<td colspan="6" bgcolor="#FFFFFF" align="center"><input type="button" value="브랜드정보 저장" onclick="SaveBrandInfo(frmbrand);"></td>
</tr>
<% else %>
<tr>
	<td colspan="6" bgcolor="#FFFFFF" align="center"><input type="button" value="브랜드정보 저장" onclick="alert('업체정보를 먼저 저장 하신후 브랜드정보를 저장 할 수 있습니다.');"></td>
</tr>
<% end if %>
-->
</form>
</table>

<br>
<br>

<%
set opartner = Nothing
set ogroup = Nothing
set ooffontract = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->