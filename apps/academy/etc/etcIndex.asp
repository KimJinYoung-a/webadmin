<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<%
response.Charset="UTF-8"
Response.ContentType="text/html;charset=UTF-8"
%>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 기타"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<!-- #include virtual="/apps/academy/lib/chkLogin.asp"-->
<%
'//뱃지 카운트 설정 및 확인
Dim ordercount, qnacount, sql, MakerID
MakerID = requestCheckVar(request.cookies("partner")("userid"),32)

sql = "exec [db_academy].[dbo].[sp_Academy_App_IconBadgeCountQnASet] '" + Cstr(MakerID) + "'"
rsACADEMYget.CursorLocation = adUseClient
rsACADEMYget.Open sql,dbACADEMYget,adOpenForwardOnly, adLockReadOnly
if not rsACADEMYget.EOF Then
	qnacount=rsACADEMYget("qnacnt")
end if
rsACADEMYget.Close
'qnacount=9
%>
</head>
<body>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content bgGry">
			<h1 class="hidden">기타</h1>
			<div class="etcMain">
				<h2 class="">게시물 관리</h2>
				<ul class="btnNav">
					<li class="boardTalk" onClick="fnAPPpopupCheerUpTalk('<%=g_AdminURL%>/apps/academy/etc/talk.asp')"><div>응원톡</div></li>
					<li class="boardReview" onClick="fnAPPpopupItemReview('<%=g_AdminURL%>/apps/academy/etc/reviewList.asp')"><div>구매후기</div></li>
					<li class="boardQna" onClick="fnAPPpopupQna('<%=g_AdminURL%>/apps/academy/etc/qnaList.asp')"><div>Q&ampA<% If qnacount>0 Then %><span class="badge"><% If qnacount>=10 Then %><% If qnacount>=100 Then %>99<i></i><% Else %><%=qnacount%><% End If %><% Else %><%=qnacount%><% End If %></span><% End If %></div></li>
				</ul>
			</div>
		</div>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<script type="text/javascript">
<!--
jQuery(document).ready(function(){
<% if qnacount > 0 then %>
fnAPPChangeBadgeCount("qnacount",<%=qnacount%>);
<% end if %>
});
//-->
</script>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->