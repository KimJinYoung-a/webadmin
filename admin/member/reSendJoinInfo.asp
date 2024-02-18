<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim brandid, qstring
brandid = requestCheckVar(Request("brandid"),32)
qstring = Request("qs")
%>
<script language="JavaScript" src="/js/jquery-1.7.2.min.js"></script>
<script language='javascript'>
function fnSendJoinPage(){
	if($("#email").val()==""){
        alert("업체 담당자 E-Mail 주소를 입력해주세요.");
	}
    else if($("#hp").val()==""){
        alert("업체 담당자 핸드폰번호를 입력해주세요.");
	}
    else{
		$.ajax({
			type: "POST",
			url: "ajaxSendJoinPage.asp",
			data: "brandid=<%=brandid%>&qs=<%=qstring%>&email=" + $("#email").val() + "&hp=" + $("#hp").val() ,
			cache: false,
			success: function(message) {
				if(message=="OK") {
					alert("발송완료했습니다.");
				} else {
					alert("발송 실패. 다시 시도해주세요.");
				}
			},
			error: function(err) {
				alert(err.responseText);
			}
		});
    }
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<tr bgcolor="#FFFFFF">
    <td>업체 담당자 E-Mail</td><td><input type="text" class="text" name="email" id="email" size="30"></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td>업체 담당자 핸드폰</td><td><input type="text" class="text" name="hp" id="hp" size="15"></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" class="button" value=" 전송 " onclick="fnSendJoinPage();"></td>
</tr>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->