<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/mobile/todaythemeCls.asp" -->
<%
'###############################################
' PageName : popSubItemEdit.asp
' Discription : 서브컨텐츠 상품코드 일괄 등록
' History : 2013.12.17 이종화 : 신규 생성
'###############################################

'// 변수 선언
Dim listidx

'// 파라메터 접수
listidx = request("listidx")

%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link href="/js/jqueryui/css/evol.colorpicker.css" rel="stylesheet">
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/evol.colorpicker.min.js"></script>
<script type="text/javascript">
$(function(){
	//라디오버튼
	$("#rdoUsing").buttonset().children().next().attr("style","font-size:11px;");
});

// 폼검사
function SaveForm(frm) {
	var selChk=true;
	if(frm.subItemidArray.value=="") {
		alert("일괄 등록하실 상품코드를 입력해주세요");
		frm.subItemidArray.focus();
		return;
	}

	if(selChk) {
		frm.submit();
	} else {
		return;
	}
}
</script>
<center>
<form name="frmSub" method="post" action="doSubRegItemCdArray.asp" style="margin:0px;">
<input type="hidden" name="listidx" value="<%=listidx%>" />
<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d" style="table-layout: fixed;">
<tr bgcolor="#FFFFFF">
    <td height="25" colspan="4" bgcolor="#F8F8F8"><b>소재 정보 - 상품코드 일괄 등록</b></td>
</tr>
<colgroup>
	<col width="100" />
	<col width="*" />
	<col width="100" />
	<col width="*" />
</colgroup>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">상품코드</td>
    <td colspan="3">
        <textarea name="subItemidArray" class="textarea" title="상품코드" style="width:100%; height:80px;"></textarea>
        <p>※ 상품코드를 쉼표(,) 또는 엔터로 구분하여 입력</p>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">정렬순서</td>
    <td>
        <input type="text" name="subSortNo" class="text" size="4" value="0" />
    </td>
    <td bgcolor="#DDDDFF">사용여부</td>
    <td>
		<span id="rdoUsing">
		<input type="radio" name="subIsUsing" id="rdoUsing1" value="Y" checked /><label for="rdoUsing1">사용</label>
		<input type="radio" name="subIsUsing" id="rdoUsing2" value="N" /><label for="rdoUsing2">삭제</label>
		</span>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="4" align="center"><input type="button" value=" 저 장 " onClick="SaveForm(this.form);"></td>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->