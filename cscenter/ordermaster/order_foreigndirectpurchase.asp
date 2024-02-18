<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 해외 직구 상품 배송정보
' History : 2018.04.25 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->

<%
dim orderserial, oUniPassNumber
	orderserial = requestCheckVar(request("orderserial"),16)

If orderserial <> "" And Not isnull(orderserial) Then
	oUniPassNumber = fnUniPassNumber(orderserial)
end if
%>
<script type="text/javascript">

document.title = "해외 직구 정보";

function fnCustomNumberSubmit(){
	var frm =  document.frm;
	if(!frm.customNumber.value || frm.customNumber.value.length < 13){
		alert('13자리의 개인통관고유부호 를 입력 해주세요.');
		frm.customNumber.focus();
		return;
	}

	var str1 = frm.customNumber.value.substring(0,1);
	var str2 = frm.customNumber.value.substring(1,13);

	if((str1.indexOf("P") < 0) == true){
		alert('P로 시작하는 13자리 번호를 입력 해주세요.');
		frm.customNumber.focus();
		return;
	}

	var regNumber = /^[0-9]*$/;
	if (!regNumber.test(str2)){
		alert('번호를 숫자만 입력해주세요.');
		frm.customNumber.focus();
		return;
	}

	frm.mode.value = "editforeigndirectpurchase";
	frm.action = "/cscenter/ordermaster/order_info_edit_process.asp";
	frm.submit();
}

</script>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" onsubmit="return false;" action="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="orderserial" value="<%=orderserial%>" />
<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td colspan="2">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    		<tr>
    			<td width="50%">
    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>해외 직구 정보</b>
			    </td>
			    <td align="right">
			    	<input type="button" value="저장하기" class="csbutton" onclick="fnCustomNumberSubmit();">
			    </td>
			</tr>
		</table>
    </td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">개인통관 고유부호</td>
    <td><input type="text" id="individualNum" name="customNumber" value="<%=oUniPassNumber%>" maxlength="14" size=14 /></td>
</tr>
</table>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
