<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/academy/lib/classes/QnA_cls.asp"-->
<%
	'// 변수 선언 //
	dim qnaId, qcd, ccd

	dim oQnA, i, lp

	'// 파라메터 접수 //
	qnaId = RequestCheckvar(request("qnaId"),10)
	qcd = RequestCheckvar(request("qcd"),10)
	ccd = RequestCheckvar(request("ccd"),10)

	'// 내용 접수
	set oQnA = new CQnA
	oQnA.FRectqnaId = qnaId
%>
<html>
<head>
<script language="javascript">
<!--
	function inputCont()
	{
		parent.frm_write.ansContents.value = document.frmCont.tempCont.value;
	}
//-->
</script>
</head>
<body onload="inputCont()">
<form name="frmCont">
<textarea name="tempCont"><%=oQnA.inputAnswerCont(qnaId,qcd,ccd)%></textarea>
</form>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->