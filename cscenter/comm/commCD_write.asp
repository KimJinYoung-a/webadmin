<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [cs]공통코드관리
' Hieditor : 이상구 생성
'			 2023.08.28 한용민 수정(고객노출여부 추가, 소스표준코드로 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/CsCommCdcls.asp"-->
<%
dim oComm, i, lp, groupCd, menupos
	groupCd     = requestCheckVar(request("groupCd"),32)
	menupos = requestCheckVar(getNumeric(request("menupos")),10)

set oComm = new CCommCd
%>
<script language="javascript" src="/admin/menu/colorbox.js"></script>
<script type='text/javascript'>

// 입력폼 검사
function chk_form(frm){
	if(!frm.comm_group.value){
		alert("그룹을 선택해주십시오.");
		frm.comm_group.focus();
		return false;
	}

	if(frm.comm_cd.value.length<4){
		alert("공통코드를 입력해주십시오.\n\n※코드는 4자리입니다.");
		frm.comm_cd.focus();
		return false;
	}

	if(!frm.comm_name.value){
		alert("코드명을 입력해주십시오.");
		frm.comm_name.focus();
		return false;
	}
	if(!frm.dispyn.value){
		alert("노출여부를 선택해주세요.");
		frm.dispyn.focus();
		return false;
	}

	return true;
}

// 코드 중복 검사
function chkDuple(ccd){
	if(ccd.length<4){
		alert("공통코드를 입력해주십시오.\n\n※코드는 4자리입니다.");
		return;
	}else{
		FrameCHK.location = "inc_chk_commCd.asp?comm_cd=" + ccd;
	}
}

</script>

<form name="frm" method="POST" onSubmit="return chk_form(this)" action="/cscenter/comm/CommCd_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="write">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="groupCd" value="<%=groupCd%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BABABA">
<tr align="center" bgcolor="#FFFFFF">
	<td height="26" align="left" colspan="2"><b>공통코드 신규등록</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">그룹</td>
	<td width="630" bgcolor="#FFFFFF">
		<select class="select" name="comm_group">
			<option value="">전체</option>
			<%= oComm.optGroupCd(groupCd)%>
		</select>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">공통코드</td>
	<td width="630" bgcolor="#FFFFFF">
		<input type="text" class="text" name="comm_cd" size="4" maxlength="4">
		<img src="/images/icon_1.gif" width="55" height="21" border="0" onClick="chkDuple(frm.comm_cd.value)" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">코드명</td>
	<td bgcolor="#FDFDFD"><input type="text" class="text" name="comm_name" size="20" maxlength="30"></td>
</tr>
<tr>
	<td bgcolor="#E8E8F1" align="center">표시색상</td>
	<td bgcolor="#FFFFFF" >
		<input type="text" name="prvColor" readonly style="background-color:'';width:21px;height:21px;border:1px solid #606060;">
		<input type="text" class="text" name="menucolor" size="7" maxlength="7" value="" readonly onClick="ShowColorBox(event.clientX, event.clientY+document.body.scrollTop)" style="cursor:pointer">
	</td>
</tr>
<tr>
	<td bgcolor="#E8E8F1" align="center">표시순서</td>
	<td bgcolor="#FFFFFF" ><input type="text" class="text" name="sortno" size="3" maxlength="8" style="text-align:right;"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">프론트노출여부</td>
	<td bgcolor="#FFFFFF" >
		<% drawSelectBoxisusingYN "dispyn", "N","" %>
	</td>
</tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="submit" value="저장" class="button">
	</td>
</tr>
</table>
</form>
<iframe name="FrameCHK" src="" frameborder="0" width="0" height="0"></iframe>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
