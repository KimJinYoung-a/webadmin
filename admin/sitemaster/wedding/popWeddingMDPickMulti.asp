<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' Discription : 웨딩 기획전 등록페이지(모바일)
' History : 2018.04.16 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<script language='javascript'>

	function SaveMainContents(frm){
	    if (frm.itemarr.value==""){
	        alert('아이템 번호를 입력 하세요.');
	        frm.itemarr.focus();
	        return;
	    }

	    if (confirm('저장 하시겠습니까?')){
	        frm.submit();
	    }
	}

</script>

<table width="100%" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmcontents" method="post" action="doWeddingMDPickUpdate.asp" onsubmit="return false;">
<input type="hidden" name="mode" value="multi">
<tr bgcolor="#FFFFFF">
    <td width="100" bgcolor="#DDDDFF">아이템번호</td>
    <td>
		<textarea name="itemarr" rows="5" cols="50"></textarea>(콤마로 구분해주세요.)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" value=" 저 장 " onClick="SaveMainContents(frmcontents);"></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
