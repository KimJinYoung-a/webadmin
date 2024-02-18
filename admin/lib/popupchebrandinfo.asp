<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->

사용 안하는 메뉴 입니다.<br>
<br>
본 페이지가 보일경우 관리자에게 신고 하세요. <br>
어떤 메뉴에서 클릭시 보여졌는지 또는 어떤 액션중에 나온 페이지인지.
<br>

<br><br>
<a href="/admin/member/popupchebrandinfo.asp?designer=<%= request("designer") %>"><font color="blue">새로운 메뉴로 이동 &gt;&gt;</font></a>

<br><br>
<font color="#999999">신규적용 Fnc :  javascript:PopUpcheBrandInfoEdit('makerid')</font> <br>
<%
'' 사용 안하는 메뉴
dbget.close()	:	response.End

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

	if (frm.jungsan_date.value.length<1){
		alert('정산일을 선택하세요.');
		frm.jungsan_date.focus();
		return;
	}

	var ret = confirm('브랜드 정보를 저장 하시겠습니까?');

	if (ret){
		frm.submit();
	}
}

function SaveUpcheInfo(frm){
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
		<input type="button" value="업체선택" onClick="PopUpcheList('frmupche');">
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
		<td colspan="4" bgcolor="#FFFFFF" height="25">**사업자등록정보**</td>
	</tr>

	<tr>
		<td width="100" bgcolor="#DDDDFF">회사명(상호)</td>
		<td bgcolor="#FFFFFF">
			<% if (C_ADMIN_AUTH=true) then %>
			<input type="text" name="company_name" value="<%= ogroup.FOneItem.FCompany_name %>" size="28" maxlength="32">
			<% else %>
			<input type="text" name="company_name" value="<%= ogroup.FOneItem.FCompany_name %>" size="28" maxlength="32" style="background-color:#EEEEEE;" readonly>
			<% end if %>
		</td>
		<td width="100" bgcolor="#DDDDFF">대표자</td>
		<td bgcolor="#FFFFFF">
			<% if (C_ADMIN_AUTH=true) then %>
			<input type="text" name="ceoname" value="<%= ogroup.FOneItem.Fceoname %>" size="16" maxlength="16">
			<% else %>
			<input type="text" name="ceoname" value="<%= ogroup.FOneItem.Fceoname %>" size="16" maxlength="16" style="background-color:#EEEEEE;" readonly>
			<% end if %>
		</td>
	</tr>
	<tr>
		<td width="100" bgcolor="#DDDDFF">사업자번호</td>
		<td bgcolor="#FFFFFF">
			<% if (C_ADMIN_AUTH=true) then %>
			<input type="text" name="company_no" value="<%= ogroup.FOneItem.Fcompany_no %>" size="16" maxlength="20">
			<% else %>
			<input type="text" name="company_no" value="<%= ogroup.FOneItem.Fcompany_no %>" size="16" maxlength="20" style="background-color:#EEEEEE;" readonly>
			<% end if %>
		</td>
		<td width="100" bgcolor="#DDDDFF">과세구분</td>
		<td bgcolor="#FFFFFF">
			<select name="jungsan_gubun">
			<option value="일반과세" <% if ogroup.FOneItem.Fjungsan_gubun="일반과세" then response.write "selected" %> >일반과세</option>
			<option value="간이과세" <% if ogroup.FOneItem.Fjungsan_gubun="간이과세" then response.write "selected" %> >간이과세</option>
			<option value="원천징수" <% if ogroup.FOneItem.Fjungsan_gubun="원천징수" then response.write "selected" %> >원천징수</option>
			<option value="면세" <% if ogroup.FOneItem.Fjungsan_gubun="면세" then response.write "selected" %> >면세</option>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">사업장소재지</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<input type="text" name="company_zipcode" value="<%= ogroup.FOneItem.Fcompany_zipcode %>" size="7" maxlength="7"><a href="javascript:popZip('s');"><img src="http://www.10x10.co.kr/images/zip_search.gif" border=0 align="absmiddle"></a><br>
			<input type="text" name="company_address" value="<%= ogroup.FOneItem.Fcompany_address %>" size="16" maxlength="64">&nbsp;
			<input type="text" name="company_address2" value="<%= ogroup.FOneItem.Fcompany_address2 %>" size="42" maxlength="64">
		</td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">업태</td>
		<td bgcolor="#FFFFFF"><input type="text" name="company_uptae" value="<%= ogroup.FOneItem.Fcompany_uptae %>" size="24" maxlength="32"></td>
		<td bgcolor="#DDDDFF">업종</td>
		<td bgcolor="#FFFFFF"><input type="text" name="company_upjong" value="<%= ogroup.FOneItem.Fcompany_upjong %>" size="24" maxlength="32"></td>
	</tr>

	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**업체기본정보**</td>
	</tr>

	<tr>
		<td bgcolor="#DDDDFF">대표전화</td>
		<td bgcolor="#FFFFFF"><input type="text" name="company_tel" value="<%= ogroup.FOneItem.Fcompany_tel %>" size="16" maxlength="16"></td>
		<td bgcolor="#DDDDFF">팩스</td>
		<td bgcolor="#FFFFFF"><input type="text" name="company_fax" value="<%= ogroup.FOneItem.Fcompany_fax %>" size="16" maxlength="16"></td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">반품 주소</td>
		<td colspan="3" bgcolor="#FFFFFF" >
		<input type="text" name="return_zipcode" value="<%= ogroup.FOneItem.Freturn_zipcode %>" size="7" maxlength="7"><a href="javascript:popZip('m');"><img src="http://www.10x10.co.kr/images/zip_search.gif" border=0 align="absmiddle"></a>
		<input type=checkbox name=samezip onclick="SameReturnAddr(this.checked)">상동
		<br>
			<input type="text" name="return_address" value="<%= ogroup.FOneItem.Freturn_address %>" size="16" maxlength="64">&nbsp;
			<input type="text" name="return_address2" value="<%= ogroup.FOneItem.Freturn_address2 %>" size="42" maxlength="64">
		</td>
	</tr>
	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**결제계좌정보**</td>
	</tr>

	<tr>
		<td bgcolor="#DDDDFF">거래은행</td>
		<td colspan="3" bgcolor="#FFFFFF" >
		<% DrawBankCombo "jungsan_bank", ogroup.FOneItem.Fjungsan_bank %>
		</td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">계좌번호</td>
		<td colspan="3" bgcolor="#FFFFFF" ><input type="text" name="jungsan_acctno" value="<%= ogroup.FOneItem.Fjungsan_acctno %>" size="16" maxlength="32">
		&nbsp;&nbsp; '-'은 빼고 번호만 입력해주시기 바랍니다.
		</td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">예금주명</td>
		<td colspan="3" bgcolor="#FFFFFF" ><input type="text" name="jungsan_acctname" value="<%= ogroup.FOneItem.Fjungsan_acctname %>" size="24" maxlength="16">
		&nbsp;&nbsp; 띄어쓰기 하지 마시기 바랍니다.
		</td>
	</tr>

	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">**담당자정보**</td>
	</tr>

	<tr>
		<td bgcolor="#DDDDFF">담당자명</td>
		<td bgcolor="#FFFFFF" width="220"><input type="text" name="manager_name" value="<%= ogroup.FOneItem.Fmanager_name %>" size="24" maxlength="32"></td>
		<td bgcolor="#DDDDFF">일반전화</td>
		<td bgcolor="#FFFFFF" width="220"><input type="text" name="manager_phone" value="<%= ogroup.FOneItem.Fmanager_phone %>" size="16" maxlength="16"></td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" name="manager_email" value="<%= ogroup.FOneItem.Fmanager_email %>" size="24" maxlength="64"></td>
		<td bgcolor="#DDDDFF">핸드폰</td>
		<td bgcolor="#FFFFFF"><input type="text" name="manager_hp" value="<%= ogroup.FOneItem.Fmanager_hp %>" size="16" maxlength="16"></td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">배송담당자명</td>
		<td bgcolor="#FFFFFF" width="220"><input type="text" name="deliver_name" value="<%= ogroup.FOneItem.Fdeliver_name %>" size="24" maxlength="32"></td>
		<td bgcolor="#DDDDFF">일반전화</td>
		<td bgcolor="#FFFFFF" width="220"><input type="text" name="deliver_phone" value="<%= ogroup.FOneItem.Fdeliver_phone %>" size="16" maxlength="16"></td>
	</tr>
	<tr>
		<td bgcolor="#DDDDFF">E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" name="deliver_email" value="<%= ogroup.FOneItem.Fdeliver_email %>" size="24" maxlength="64"></td>
		<td bgcolor="#DDDDFF">핸드폰</td>
		<td bgcolor="#FFFFFF"><input type="text" name="deliver_hp" value="<%= ogroup.FOneItem.Fdeliver_hp %>" size="16" maxlength="16"></td>
	</tr>

	<tr>
		<td width="80" bgcolor="#DDDDFF">정산담당자명</td>
		<td bgcolor="#FFFFFF" width="220"><input type="text" name="jungsan_name" value="<%= ogroup.FOneItem.Fjungsan_name %>" size="24" maxlength="32"></td>
		<td width="80" bgcolor="#DDDDFF">일반전화</td>
		<td bgcolor="#FFFFFF" width="220"><input type="text" name="jungsan_phone" value="<%= ogroup.FOneItem.Fjungsan_phone %>" size="16" maxlength="16"></td>
	</tr>
	<tr>
		<td width="60" bgcolor="#DDDDFF">E-Mail</td>
		<td bgcolor="#FFFFFF"><input type="text" name="jungsan_email" value="<%= ogroup.FOneItem.Fjungsan_email %>" size="24" maxlength="64"></td>
		<td width="60" bgcolor="#DDDDFF">핸드폰</td>
		<td bgcolor="#FFFFFF"><input type="text" name="jungsan_hp" value="<%= ogroup.FOneItem.Fjungsan_hp %>" size="16" maxlength="16"></td>
	</tr>
	<tr>
		<td colspan="4" bgcolor="#FFFFFF" align="center"><input type="button" value="업체정보 저장" onclick="SaveUpcheInfo(frmupche);"></td>
	</tr>
</form>
</table>

<br>

<table width="600" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#000000>
<form name="frmbrand" method="post" action="doupcheedit.asp">
<input type="hidden" name="uid" value="<%= opartner.FOneItem.FID %>">
<input type="hidden" name="mode" value="brandedit">
<tr>
	<td bgcolor="#FFDDDD" colspan=6><b>2.브랜드관련정보</b></td>
</tr>
<tr>
	<td width="100" bgcolor="#FFDDDD">회사명</td>
	<td bgcolor="#FFFFFF" colspan=2><%= opartner.FOneItem.FCompany_name %></td>
	<td width="100" bgcolor="#FFDDDD" >랙번호</td>
	<td bgcolor="#FFFFFF" colspan=2>
	<input type="text" name="prtidx" value="<%= opartner.FOneItem.getRackCode %>" size="4" maxlength="4">
	(기본값 : 9999)</td>
	</td>
</tr>
<tr>
	<td bgcolor="#FFDDDD">브랜드ID</td>
	<td bgcolor="#FFFFFF" colspan=2><%= opartner.FOneItem.FID %></td>
	<td bgcolor="#FFDDDD">패스워드</td>
	<td bgcolor="#FFFFFF" colspan=2>
	<input type="text" name="password" value="<%= opartner.FOneItem.Fppass %>">
	</td>
</tr>
<tr>
	<td bgcolor="#FFDDDD">스트리트명<br>(한글)</td>
	<td bgcolor="#FFFFFF" colspan=2>
	<input type="text" name="socname_kor" value="<%= opartner.FOneItem.Fsocname_kor %>">
	</td>
	<td bgcolor="#FFDDDD">스트리트명<br>(영문)</td>
	<td bgcolor="#FFFFFF" colspan=2>
	<input type="text" name="socname" value="<%= opartner.FOneItem.Fsocname %>">
	</td>
</tr>
<tr>
	<td rowspan="3" bgcolor="#FFDDDD">브랜드<br>사용여부<br>(카테고리노출)</td>
	<td bgcolor="#FFFFFF">텐바이텐</td>
	<td bgcolor="#FFFFFF"><input type=radio name="isusing" value="Y" <% if opartner.FOneItem.Fisusing="Y" then response.write "checked" %> >사용 <input type=radio name="isusing" value="N" <% if opartner.FOneItem.Fisusing="N" then response.write "checked" %> >사용안함</td>
	<td rowspan="3" bgcolor="#FFDDDD">스트리트<br>표시여부<br>(브랜드운영관련)</td>
	<td bgcolor="#FFFFFF">텐바이텐</td>
	<td bgcolor="#FFFFFF"><input type=radio name="streetusing" value="Y" <% if opartner.FOneItem.Fstreetusing="Y" then response.write "checked" %> >사용 <input type=radio name="streetusing" value="N" <% if opartner.FOneItem.Fstreetusing="N" then response.write "checked" %> >사용안함</td>
</tr>
<tr >
	<td bgcolor="#FFFFFF">제휴몰</td>
	<td bgcolor="#FFFFFF"><input type=radio name="isextusing" value="Y" <% if opartner.FOneItem.Fisextusing="Y" then response.write "checked" %> >사용 <input type=radio name="isextusing" value="N" <% if opartner.FOneItem.Fisextusing="N" then response.write "checked" %> >사용안함	</td>
	<td bgcolor="#FFFFFF">제휴몰</td>
	<td bgcolor="#FFFFFF"><input type=radio name="extstreetusing" value="Y" <% if opartner.FOneItem.Fextstreetusing="Y" then response.write "checked" %> >사용 <input type=radio name="extstreetusing" value="N" <% if opartner.FOneItem.Fextstreetusing="N" then response.write "checked" %> >사용안함	</td>
</tr>
<tr >
	<td bgcolor="#FFFFFF" colspan=2></td>
	<td bgcolor="#FFFFFF">커뮤니티</td>
	<td bgcolor="#FFFFFF"><input type=radio name="specialbrand" value="Y" <% if opartner.FOneItem.Fspecialbrand="Y" then response.write "checked" %>>사용 <input type=radio name="specialbrand" value="N" <% if opartner.FOneItem.Fspecialbrand="N" then response.write "checked" %>>사용안함</td>
</tr>
<tr >
	<td bgcolor="#FFDDDD">업체구분</td>
	<td bgcolor="#FFFFFF" colspan=2><% DrawBrandGubunCombo "userdiv", opartner.FOneItem.Fuserdiv %></td>
	<td bgcolor="#FFDDDD">등록일</td>
	<td bgcolor="#FFFFFF" colspan=2><%= opartner.FOneItem.Fregdate %></td>
</tr>
<tr >
	<td bgcolor="#FFDDDD">카테고리</td>
	<td bgcolor="#FFFFFF" colspan=2><% SelectBoxBrandCategory "catecode", opartner.FOneItem.Fcatecode %></td>
	<td bgcolor="#FFDDDD" >담당MD</td>
	<td bgcolor="#FFFFFF" colspan=2><% drawSelectBoxCoWorker "mduserid", opartner.FOneItem.Fmduserid %></td>
</tr>
<tr>
	<td bgcolor="#FFDDDD" >반품송장</td>
	<td bgcolor="#FFFFFF" colspan=5>

	<input type=text name=brandsongjang value="<%= returnsongjangStr %>" size=50 > <a href="javascript:copyComp(frmbrand.brandsongjang);">복사</a>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan=6>**계약관련사항**</td>
</td>
<tr >
	<td bgcolor="#FFDDDD" >기본마진</td>
	<td bgcolor="#FFFFFF" colspan=2>
	<% DrawBrandMWUCombo "maeipdiv",opartner.FOneItem.Fmaeipdiv %>
	<input type="text" name="defaultmargine" value="<%= opartner.FOneItem.Fdefaultmargine %>" size="4" style="text-align:right"> %
	</td>
	<td bgcolor="#FFDDDD" >정산일 :</td>
	<td bgcolor="#FFFFFF" colspan=2>
	<% DrawJungsanDateCombo "jungsan_date", opartner.FOneItem.Fjungsan_date %>
	</td>
</tr>
<tr >
	<td bgcolor="#FFDDDD" >오프라인(직영점)</td>
	<td bgcolor="#FFFFFF" colspan=2>
		<table border=0 cellspacing=0 cellpadding=0 class=a width=100%>
		<tr>
			<td width="100"><b>직영점대표</b></td>
			<td><%= ooffontract.GetSpecialChargeDivName("streetshop000") %></td>
			<td width="40"><%= ooffontract.GetSpecialDefaultMargin("streetshop000") %> %</td>
		</tr>
		<% for i=0 to ooffontract.FResultCount-1 %>
		<% if (ooffontract.FItemList(i).Fshopdiv<>"3") and (ooffontract.FItemList(i).Fshopid<>"streetshop000") then %>
		<tr>
			<td><%= ooffontract.FItemList(i).Fshopname %></td>
			<td><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
			<td><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
		</tr>
		<% end if %>
		<% next %>
		</table>
	</td>
	<td bgcolor="#FFDDDD" >정산일 </td>
	<td bgcolor="#FFFFFF" colspan=2>
	<% DrawJungsanDateCombo "jungsan_date_off", opartner.FOneItem.Fjungsan_date_off %>
	</td>
</tr>
<tr >
	<td bgcolor="#FFDDDD" >오프라인(가맹점)</td>
	<td bgcolor="#FFFFFF" colspan=2>
		<table border=0 cellspacing=0 cellpadding=0 class=a width=100%>
		<tr>
			<td width="100"><b>가맹점점대표</b></td>
			<td><%= ooffontract.GetSpecialChargeDivName("streetshop800") %></td>
			<td width="40"><%= ooffontract.GetSpecialDefaultMargin("streetshop800") %> %</td>
		</tr>
		<% for i=0 to ooffontract.FResultCount-1 %>
		<% if (ooffontract.FItemList(i).Fshopdiv="3") and (ooffontract.FItemList(i).Fshopid<>"streetshop800") then %>
		<tr>
			<td ><%= ooffontract.FItemList(i).Fshopname %></td>
			<td ><%= ooffontract.FItemList(i).GetChargeDivName() %></td>
			<td><%= ooffontract.FItemList(i).Fdefaultmargin %> %</td>
		</tr>
		<% end if %>
		<% next %>
		</table>
	</td>
	<td bgcolor="#FFDDDD" >정산일 </td>
	<td bgcolor="#FFFFFF" colspan=2>
	<% DrawJungsanDateCombo "jungsan_date_frn", opartner.FOneItem.Fjungsan_date_frn %>
	</td>
</tr>


<!--
<tr>
	<td bgcolor="#FFDDDD" >어드민 오픈여부</td>
	<td bgcolor="#FFFFFF" colspan=5>
		<% if opartner.FOneItem.Fpartnerusing="Y" then %>
		<input type="radio" name="partnerusing" value="Y" checked >사용함
		<input type="radio" name="partnerusing" value="N" >사용안함
		<% else %>
		<input type="radio" name="partnerusing" value="Y"  >사용함
		<input type="radio" name="partnerusing" value="N" checked ><font color=red>사용안함</font>
		<% end if %>
	</td>
</tr>
-->

<% if ogroup.FOneItem.FGroupId<>"" then %>
<tr>
	<td colspan="6" bgcolor="#FFFFFF" align="center"><input type="button" value="브랜드정보 저장" onclick="SaveBrandInfo(frmbrand);"></td>
</tr>
<% else %>
<tr>
	<td colspan="6" bgcolor="#FFFFFF" align="center"><input type="button" value="브랜드정보 저장" onclick="alert('업체정보를 먼저 저장 하신후 브랜드정보를 저장 할 수 있습니다.');"></td>
</tr>
<% end if %>
</form>
</table>

<br>

<table  width="600" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#000000>
<form name="frmetc" method=post action="http://partner.10x10.co.kr/linkweb/doprofileimageadmin.asp" enctype="multipart/form-data">
<input type=hidden name=designerid value="<%= opartner.FOneItem.FID %>">
	<tr>
		<td bgcolor="#DDDDFF" colspan=4><b>3.브랜드 기타정보</b></td>
	</tr>
	<tr>
		<td width="110" bgcolor="#DDDDFF" >로고</td>
		<td bgcolor="#FFFFFF">
		<img name=logoimg src="<%= opartner.FOneItem.getSocLogoUrl %>" width=150 height=100><br>
		(브랜드 로고는 150x100 픽셀로 지정되 있습니다.)<br>
		<input type=file name=file1 size=40 onchange="ChangeLogo(this,frmetc.logoimg);">
		</td>
	</tr>
	<tr>
		<td width="110" bgcolor="#DDDDFF" >타이틀</td>
		<td bgcolor="#FFFFFF">
		<img name=titleimg src="<%= opartner.FOneItem.getTitleImgUrl %>" width=300 height=75><br>
		(타이틀이미지는 600x150 픽셀로 지정되 있습니다.)(600x150)<br>
		<input type=file name=file2 size=40 onchange="ChangeTitle(this,frmetc.titleimg);">
		</td>
	</tr>
	<tr>
		<td width="110" bgcolor="#DDDDFF" >디자이너<br>코멘트</td>
		<td bgcolor="#FFFFFF">
		<textarea name="dgncomment" cols=64 rows=6><%= opartner.FOneItem.Fdgncomment %></textarea>
		</td>
	</tr>
	<tr>
		<td colspan="2" align=center bgcolor="#FFFFFF"><input type="button" value="브랜드 기타정보 저장" onclick="SaveBrandEtcInfo(frmetc);"></td>
	</tr>
</form>
</table>
<%
set opartner = Nothing
set ogroup = Nothing
set ooffontract = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->