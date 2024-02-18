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
dim comm_cd, oComm, i, lp, menupos
	comm_cd     = requestCheckVar(request("comm_cd"),32)
	menupos = requestCheckVar(getNumeric(request("menupos")),10)

set oComm = new CCommCd
	oComm.FRectcommCd = comm_cd
	oComm.GetCommRead

if (oComm.FResultCount = 0) then
	response.write "<script type='text/javascript'>alert('존재하지 않는 코드입니다.');history.back();</script>"
	dbget.close()	:	response.End
end if
%>
<script language="javascript" src="/admin/menu/colorbox.js"></script>
<script type='text/javascript'>

// 입력폼 검사
function chk_form(frm){
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

	// 폼 전송
	return true;
}

</script>

<form name="frm" method="POST" onSubmit="return chk_form(this)" action="/cscenter/comm/CommCd_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="modify">
<input type="hidden" name="comm_cd" value="<%=comm_cd%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BABABA">
<tr align="center" bgcolor="#FFFFFF">
	<td height="26" align="left" colspan="2"><b>공통코드 상세 내용 / 수정</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">그룹</td>
	<td width="630" bgcolor="#FFFFFF"><%=oComm.FItemList(0).Fgroup_name%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">공통코드</td>
	<td width="630" bgcolor="#FFFFFF"><b><%=oComm.FItemList(0).Fcomm_cd%></b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">코드명</td>
	<td bgcolor="#FDFDFD"><input type="text" name="comm_name" value="<%=db2html(oComm.FItemList(0).Fcomm_name)%>" size="20" maxlength="30"></td>
</tr>
<tr>
	<td bgcolor="#E8E8F1" align="center">표시색상</td>
	<td bgcolor="#FFFFFF" >
		<input type="text" name="prvColor" readonly style="background-color:'<%= oComm.FItemList(0).Fcomm_color %>';width:21px;height:21px;border:1px solid #606060;">
		<input type="text" name="menucolor" size="7" maxlength="7" value="<%= oComm.FItemList(0).Fcomm_color %>" readonly onClick="ShowColorBox(event.clientX, event.clientY+document.body.scrollTop)" style="cursor:pointer">
	</td>
</tr>
<tr>
	<td bgcolor="#E8E8F1" align="center">표시순서</td>
	<td bgcolor="#FFFFFF" ><input type="text" class="text" name="sortno" value="<%=db2html(oComm.FItemList(0).Fsortno)%>" size="3" maxlength="8" style="text-align:right;"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">사용여부</td>
	<td bgcolor="#FFFFFF" >
		<input type="radio" name="comm_isDel" value="N" <% if oComm.FItemList(0).Fcomm_isDel="사용" then Response.Write "checked"%>> 사용 &nbsp; &nbsp;
		<input type="radio" name="comm_isDel" value="Y" <% if oComm.FItemList(0).Fcomm_isDel="삭제" then Response.Write "checked"%>> 삭제
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">프론트노출여부</td>
	<td bgcolor="#FFFFFF" >
		<% drawSelectBoxisusingYN "dispyn", oComm.FItemList(0).fdispyn,"" %>
	</td>
</tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="submit" value="저장" class="button">
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
