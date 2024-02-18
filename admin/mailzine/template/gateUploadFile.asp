<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 메일진
' History : 2018.04.27 이상구 생성(메일러 연동 생성 메일러로 발송 내역 전송. 메일 가져오기 생성.)
'			2019.06.24 정태훈 수정(템플릿 기능 신규 추가)
'			2020.05.28 한용민 수정(TMS 메일러 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/util/htmllib.asp"--> 
<%
	Dim sName, strImgUrl
	sName	= requestCheckVar(Request("sName"),50) 
	strImgUrl	= requestCheckVar(Request("sImgUrl"),100)
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script language="javascript">
<!--
	alert("이미지가 등록되었습니다.\n\n이미지 등록후 저장버튼을 눌러야 처리완료됩니다.");
	$("input[name='<%=sName%>']",opener.document).val("<%=strImgUrl%>");
	$("#<%=sName%>",opener.document).html("<%=strImgUrl%>");
	window.close();
//-->
</script>