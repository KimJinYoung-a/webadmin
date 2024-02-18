<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  사원권한등록
' History : 정윤정 생성
'			2018.04.17 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%
IF application("Svr_Info")<>"Dev" THEN
	if Not(C_privacyadminuser) or Not(isVPNConnect) then
			response.write "승인된 페이지가 아닙니다. 관리자 문의요망 [접근권한:" & C_privacyadminuser & "/VPN:" & isVPNConnect & "]"
			response.end
	end if
end if

Dim sEmpNo,sUsername
Dim suserid,sfrontid, ipart_sn,iposit_sn,ijob_sn ,ilevel_sn,iuserdiv, lv1customerYN, lv2partnerYN, lv3InternalYN, icriticinfouser
Dim cMember, i, mydpID, isdispmember

sEmpNo = requestCheckVar(Request("sEPN"),14)
IF 	sEmpNo = "" THEN
    Alert_return("잘못된 유입경로입니다.")
    response.end ''추가/2014/07/14
END IF

IF application("Svr_Info")="Dev" THEN
	isdispmember = true
else
	' ISMS 심사로 인해 개인정보 접근권한 생성/수정/변경 특정사람만 보이게(한용민,허진원,이문재)	' 2020.10.12 한용민
	if C_privacyadminuser or C_PSMngPart then
		isdispmember = true
	else
		isdispmember = false
	end if
end if

set cMember = new CTenByTenMember
	cMember.Fempno = sEmpNo
	cMember.fnGetMemberData
	sempno   		= cMember.Fempno
	suserid     	= cMember.Fuserid
	sfrontid    	= cMember.Ffrontid
	sUsername 		= cMember.Fusername
	ipart_sn        = cMember.Fpart_sn
	iposit_sn       = cMember.Fposit_sn
	ijob_sn         = cMember.Fjob_sn
	ilevel_sn       = cMember.Flevel_sn
	iuserdiv        = cMember.Fuserdiv
	icriticinfouser = cMember.Fcriticinfouser
	lv1customerYN = cMember.Flv1customerYN
	lv2partnerYN = cMember.Flv2partnerYN
	lv3InternalYN = cMember.Flv3InternalYN
	mydpID = myDepartmentId(session("ssBctID"))
set 	cMember = nothing

Dim oAddLevel
set oAddLevel = new CPartnerAddLevel
oAddLevel.FRectUserid=suserid
oAddLevel.FRectOnlyAdd = "on"

if (oAddLevel.FRectUserID<>"") then
    oAddLevel.getUserAddLevelList
end if

dim olog
Set olog = new CTenByTenMember
	olog.FPagesize = 50
	olog.FCurrPage = 1
	olog.frectempno=sEmpNo
	olog.getUserTenbytenAdminAuthLog()
%>

<html>
<head>
<title>직원정보 등록</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language="javascript" src="/js/common.js"></script>
<script type="text/javascript">

function chRetireUser(){
	if(!document.frmRetireUser.sUI.value) {
		alert("어드민 아이디가 없습니다");
		return;
	}
	if(!document.frmRetireUser.sFUI.value) {
		alert("텐바이텐사이트 아이디가 없습니다");
		return;
	}
	if(!document.frmRetireUser.sUN.value) {
		alert("이름이 입력되지 않았습니다.");
		return;
	}
	if (frm_member.selPN.value!='35' || frm_member.selLN.value!='' || frm_member.lv1customerYN.checked!=false || frm_member.lv2partnerYN.checked!=false || frm_member.lv3InternalYN.checked!=false){
		alert("어드민 권한이 사용중에 있습니다.\n권한을 제외시키시고 이용해 주세요.");
		return;
	}
	if (confirm('해당 프론트 아이디를 퇴사처리 하시겠습니까?\n->직원쿠폰 회수\n->회원등급 변경')==true){
		frmRetireUser.submit();
	}
}

function jsChkSubmit(){
	<% if not(C_ADMIN_AUTH or C_PSMngPartPower) then %>
		if(!document.frm_member.sUI.value) {
			alert("WEBADMIN 아이디를 입력해주십시요.");
			document.frm_member.sUI.focus();
				return ;
		}
	<% else %>
		if(document.frm_member.sUI.value == "") {
			if (confirm("[관리자권한] WEBADMIN 아이디가 없습니다.\n\n계속 진행하시겠습니까?") != true) {
				alert("WEBADMIN 아이디를 입력해주십시요.");
				document.frm_member.sUI.focus();
				return ;
			}
		}
	<% end if %>

	// WEBADMIN 아이디가 입력되어 있을경우에만 체크함
	if(frm_member.sUI.value != "") {
		if(typeof(document.frm_member.sP) != "undefined"){
			if(!document.frm_member.sP.value) {
				alert("비밀번호를 입력해주십시요.");
				document.frm_member.sP.focus();
				return ;
			}

			if (document.frm_member.sP.value.replace(/\s/g, "").length < 6 || document.frm_member.sP.value.replace(/\s/g, "").length > 16){
				alert("비밀번호는 공백없이 6~16자입니다.");
				document.frm_member.sP.focus();
				return ;
			}

			if ((document.frm_member.sP.value)!=(document.frm_member.sP1.value)){
				alert("비밀번호가 일치하지 않습니다.");
				document.frm_member.sP1.focus();
				return;
			}

			if (!fnChkComplexPassword(frm_member.sP.value)) {
				alert('새로운 패스워드는 영문/숫자/특수문자 등 두가지 이상의 조합으로 입력하세요. 최소길이 10자(2조합) , 8자(3조합)');
				frm_member.sP.focus();
				return;
			}
		}
	}

	if(document.frm_member.hidID.value =="0"){
		alert("아이디 중복체크를 해주세요");
		return ;
	}

	if(document.frm_member.selPN.value == "") {
		alert("어드민권한(부서)을 입력해주십시요.");
		return ;
	}

	if(document.frm_member.selLN.value == "") {
		if (confirm("!!! 어드민 권한등급 삭제 !!!\n\n진행하시겠습니까?") == true) {
			//
		} else {
			alert("어드민권한(등급)을 입력해주십시요.");
			return;
		}
	}

	if ((document.frm_member.selUD.value=="")||(document.frm_member.selUD.value*1>111)){
		alert("기존권한설정오류 ");
		return ;
	}

	document.frm_member.submit();
}
			//아이디 중복체크
	function jsChkID(){
		var winID;
		var frm = document.frmID;
		if(!document.frm_member.sUI.value){
			alert("아이디를 입력해주세요");
			document.frm_member.sUI.focus();
			return;
		}
		frmID.sUI.value = document.frm_member.sUI.value;
		frmID.sUN.value = document.frm_member.sUN.value;
		winID = window.open("","popid","width=0, height=0");
		document.frmID.target = "popid";
		document.frmID.submit();
	}
	function jsChkfrontname(){
		var winID;
		var frm = document.frmchk;
		if(!document.frm_member.sUI.value){
			alert("아이디를 입력해주세요");
			document.frm_member.sUI.focus();
			return;
		}
		frmchk.sUI.value = document.frm_member.sUI.value;
		frmchk.sUN.value = document.frm_member.sUN.value;
		frmchk.sFUI.value = document.frm_member.sFUI.value;
		frmchk.sEN.value = "<%= sEmpNo %>";
		winID = window.open("","popid","width=0, height=0");
		document.frmchk.target = "popid";
		frmchk.mode.value="frontnamewebadmincheck";
		document.frmchk.submit();
		frmchk.mode.value="";
	}
	function jschangefrontname(){
		var winID;
		var frm = document.frmchk;
		if(!document.frm_member.sUI.value){
			alert("아이디를 입력해주세요");
			document.frm_member.sUI.focus();
			return;
		}
		frmchk.sUI.value = document.frm_member.sUI.value;
		frmchk.sUN.value = document.frm_member.sUN.value;
		frmchk.sFUI.value = document.frm_member.sFUI.value;
		frmchk.sEN.value = "<%= sEmpNo %>";

		var ret = confirm('동일인이 확실한 경우에만 변경하셔야 합니다.\n직원명강제변경[어드민->프론트사이트] 하시겠습니까?');
		if (ret){
			winID = window.open("","popid","width=0, height=0");
			document.frmchk.target = "popid";
			frmchk.mode.value="frontnamewebadminchange";
			document.frmchk.submit();
			frmchk.mode.value="";
		}
	}

	// 권한 선택 팝업
	function popAuthSelect()
	{
		<% if application("Svr_Info")<>"Dev" THEN %>
			<% if Not(C_ADMIN_AUTH or C_PSMngPartPower) then %>
				alert('권한이 없습니다. 관리자 문의 요망');
				return;
			<% end if %>
		<% end if %>

		var popwin = window.open("/admin/menu/pop_Menu_auth.asp?userid=<%= suserid %>", "popMenuAuth","width=360,height=200,scrollbars=no");
		popwin.focus();
	}

	// 팝업에서 선택권한 추가
	function addAuthItem(psn,pnm,lsn,lnm)
	{
		var lenRow = tbl_auth.rows.length;

		// 기존에 값에 중복 파트 여부 검사
		if(lenRow>1)	{
			for(l=0;l<document.all.part_sn.length;l++)	{
				if(document.all.part_sn[l].value==psn) {
					alert("이미 권한이 지정된 부서입니다.\n기존 부서를 삭제하고 권한을 다시 지정해주세요.");
					return;
				}
			}
		}
		else {
			if(lenRow>0) {
				if(document.all.part_sn.value==psn) {
					alert("이미 권한이 지정된 부서입니다.\n기존 부서를 삭제하고 권한을 다시 지정해주세요.");
					return;
				}
			}
		}

		// 행추가
		var oRow = tbl_auth.insertRow(lenRow);
		oRow.onmouseover=function(){tbl_auth.clickedRowIndex=this.rowIndex};

		// 셀추가 (부서,등급,삭제버튼)
		var oCell1 = oRow.insertCell(0);
		var oCell2 = oRow.insertCell(1);
		var oCell3 = oRow.insertCell(2);

		oCell1.innerHTML = pnm + "<input type='hidden' name='part_sn' value='" + psn + "'>";
		oCell2.innerHTML = lnm + "<input type='hidden' name='level_sn' value='" + lsn + "'>";
		oCell3.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delAuthItem()' align=absmiddle>";
	}

	// 선택권한 삭제
	function delAuthItem()
	{
		<% if application("Svr_Info")<>"Dev" THEN %>
			<% if not(C_ADMIN_AUTH or C_PSMngPartPower) then %>
				alert('권한이 없습니다. 관리자 문의 요망');
				return;
			<% end if %>
		<% end if %>
		if(confirm("선택한 권한을 삭제하시겠습니까?"))
			tbl_auth.deleteRow(tbl_auth.clickedRowIndex);
	}
	//비밀번호변경
	function jsChangePW(uid){
	    var popwinPass = window.open("/admin/member/tenbyten/pop_ChangPassword.asp?userid="+uid,"popwinPass","width=1024,height=400,scrollbars=yes,resizable=yes");
		popwinPass.focus();
		
	}
</script>
</head>
<body leftmargin="10" topmargin="10">
<form name="frm_member" method="post" action="/admin/member/tenbyten/procAdminAuth.asp" style="margin:0px;">
<input type="hidden" name="hidID" value="1">
<input type="hidden" name="sEN" value="<%=sEmpNo%>">
<input type="hidden" name="selPoN" value="<%=iposit_sn%>">
<input type="hidden" name="selJN" value="<%=ijob_sn%>">
<input type="hidden" name="sUN" value="<%=sUsername%>">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><strong>사원 권한정보 등록</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="3" cellspacing="1" align="center" class="a" bgcolor=#BABABA>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>" width="130">사번</td>
			<td bgcolor="#FFFFFF">
				<%= sempno %>
			</td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>" width="130">WEBADMIN 아이디</td>
			<td bgcolor="#FFFFFF">
					<input type="text" name="sUI" class="text" size="20" value="<%=suserid%>" onClick="document.frm_member.hidID.value=0;" onKeypress="document.frm_member.hidID.value=0;"> <input type="button" name="btnChkID" value="아이디 중복체크" onClick="jsChkID();" class="input">
			</td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">텐바이텐사이트 아이디</td>
			<td bgcolor="#FFFFFF">
				<input type="text" name="sFUI" class="text" size="20" value="<%=sfrontid%>">
				<input type="button" name="btnChkID" value="직원명동일체크[어드민,프론트사이트]" onClick="jsChkfrontname();" class="input">

				<% if isdispmember then %>
					<% if C_ADMIN_AUTH or C_PSMngPart then %>
						<input type="button" name="btnChkID" value="직원명강제변경[어드민->프론트사이트]" onClick="jschangefrontname();" class="input">
					<% end if %>
				<% end if %>
			</td>
		</tr>
		<% IF isNull(suserid) or suserid = "" THEN %>
			<tr align="left" height="25">
				<td bgcolor="<%= adminColor("tabletop") %>">비밀번호</td>
				<td bgcolor="#FFFFFF">
					<input type="password" name="sP" class="text" size="20" maxlength="60" value="">
				</td>
			</tr>
			<tr align="left" height="25">
				<td bgcolor="<%= adminColor("tabletop") %>">비밀번호확인</td>
				<td bgcolor="#FFFFFF">
					<input type="password" name="sP1" class="text" size="20" maxlength="60" value="">
				</td>
			</tr>
		<% ELSE %>
			<tr align="left" height="25">
				<td bgcolor="<%= adminColor("tabletop") %>">패스워드</td>
				<td bgcolor="#FFFFFF">
					<% if isdispmember then %>
						<% If (C_ADMIN_AUTH) or C_PSMngPart OR (mydpID = "88") Then %>
							<input type="button" class="button" value="변경(아이디 로그인용 패스워드)" onClick="jsChangePW('<%=suserid%>');">
						<% End If %>
					<% End If %>
					<br>※ 패스워드 변경시 초기화 되는 기능
					<br>1. 계정이 사용안함 상태 인경우, 사용함으로 변경됨
					<br>2. 장기간 미사용으로 인해 계정이 잠긴경우, 잠김이 해제됨.
					<br>3. 패스워드를 틀려서 잠긴경우, 잠김이 해제 됩니다.
				</td>
			</tr>
		<% END IF %>
	    <tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">어드민 권한(등급)</td>
			<td bgcolor="#FFFFFF">
				<%=printPartOption("selPN", ipart_sn)%>
				&nbsp;&nbsp;
				<%=printLevelOption("selLN", ilevel_sn)%>
			</td>
		</tr>
	    <tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">개인정보취급권한</td>
			<td bgcolor="#FFFFFF">
				<% 'Call DrawSelectBoxCriticInfoUser("criticinfouser", icriticinfouser) %>
				<input type="hidden" name="criticinfouser" value="<%= icriticinfouser %>">
				<input type="checkbox" name="lv1customerYN" value="Y" <% if lv1customerYN = "Y" then %>checked<% end if %> >LV1(고객정보)
				<input type="checkbox" name="lv2partnerYN" value="Y" <% if lv2partnerYN = "Y" then %>checked<% end if %> >LV2(파트너정보)
				<input type="checkbox" name="lv3InternalYN" value="Y" <% if lv3InternalYN = "Y" then %>checked<% end if %> >LV3(내부정보)
			</td>
		</tr>
	    <tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">추가 권한</td>
			<td bgcolor="#FFFFFF">
			    <table border="0" cellspacing="0" class="a">
			    <tr>
			        <td >
        			    <table name='tbl_auth' id='tbl_auth' class=a>
        			    <% if (oAddLevel.FResultCount>0) then %>
        			        <% for i=0 to oAddLevel.FResultCount-1 %>
        			        <tr onMouseOver='tbl_auth.clickedRowIndex=this.rowIndex'>
        						<td><%= oAddLevel.FItemList(i).Fpart_name %><input type='hidden' name='part_sn' value='<%= oAddLevel.FItemList(i).Fpart_sn %>'></td>
        						<td><%= oAddLevel.FItemList(i).Flevel_name %><input type='hidden' name='level_sn' value='<%= oAddLevel.FItemList(i).Flevel_sn %>'></td>
        						<td><img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delAuthItem()' align=absmiddle></td>
        					</tr>
        				    <% next %>
        				<% else %>
        				    <tr onMouseOver='tbl_auth.clickedRowIndex=this.rowIndex'>
						    <td><input type='hidden' name='part_sn' value=''></td>
						    <td><input type='hidden' name='level_sn' value=''></td>
						    <td></td>
					        </tr>
        				<% end if %>
        			    </table>
			        </td>
        			<td valign="bottom"><input type="button" class='button' value="추가" onClick="popAuthSelect()"></td>
        		</tr>
        		</table>


			</td>
		</tr>
		<tr align="left" height="25">
			<td bgcolor="<%= adminColor("tabletop") %>">기존권한</td>
			<td bgcolor="#FFFFFF">
				<% DrawAuthBoxTenUser "selUD",iUserdiv %>
			</td>
		</tr>

		</table>

	</td>
</tr>
<Tr>
	<td align="center">
		<% if isdispmember then %>
			<input type="button" value="저장" class="input" onclick="jsChkSubmit();">
			<% if C_ADMIN_AUTH or C_PSMngPart then %>
				<% if sfrontid<>"" and not(isnull(sfrontid)) and sUsername<>"" and not(isnull(sUsername)) then %>
					&nbsp;&nbsp;
					<input type="button" value="프론트퇴사처리(쿠폰회수,등급변경)" class="input" onclick="chRetireUser();">
				<% end if %>
			<% end if %>
		<% end if %>
	</td>
</tr>
</table>
</form>

<% if olog.FResultCount>0 then %>
	<br>
	<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%=olog.FtotalCount%></b>
			&nbsp;&nbsp;※ 최근 50건만 표기 됩니다.
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width=60>로그번호</td>
		<td>변경내용</td>
		<td width=100>변경일</td>
	</tr>
	<% for i=0 to olog.FResultCount - 1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= olog.FitemList(i).flogidx %></td>
		<td align="left"><%= olog.FitemList(i).flogmsg %></td>
		<td>
			<%= olog.FitemList(i).fadminid %>
			<Br><%= left(olog.FitemList(i).fregdate,10) %>
			<Br><%= mid(olog.FitemList(i).fregdate,12,22) %>
		</td>
	</tr>
	<% next %>
	</table>
<% end if %>

<!-- 아이디 중복체크-->
<form name="frmID" method="post" action="/admin/member/tenbyten/member_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="R">
<input type="hidden" name="sEN" value="">
<input type="hidden" name="sUI" value="">
<input type="hidden" name="sUN" value="">
</form>
<form name="frmRetireUser" method="post" action="/admin/member/tenbyten/member_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="RetireUser">
<input type="hidden" name="sUI" value="<%=suserid%>">
<input type="hidden" name="sFUI" value="<%=sfrontid%>">
<input type="hidden" name="sUN" value="<%=sUsername%>">
</form>
<form name="frmchk" method="post" action="/admin/member/tenbyten/member_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="sEN" value="">
<input type="hidden" name="sUI" value="">
<input type="hidden" name="sUN" value="">
<input type="hidden" name="sFUI" value="">
</form>
