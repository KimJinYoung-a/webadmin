<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<% response.Charset="UTF-8" %>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 공지사항"
%>
<!-- #include virtual="/apps/academy/lib/chkLogin.asp"-->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/commlib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/apps/academy/notice/lecturerNoticecls.asp"-->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<%
Dim vidx, UserID
Dim olect, MakerID

vidx = requestCheckVar(request("idx"),10)
MakerID	= requestCheckVar(request.cookies("partner")("userid"),32)

Set olect = New ClecturerList
olect.FrectDoc_Idx = vidx
olect.fnGetlecturerView

Function eregi_replace(pattern, replace, text)
	Dim eregObj
	' Create regular expression
	Set eregObj= New RegExp 
	eregObj.Pattern= pattern ' Set Pattern(패턴 설정)
	eregObj.IgnoreCase = True ' Set Case Insensitivity(대소문자 구분 여부)
	eregObj.Global = True ' Set All Replace(전체 문서에서 검색)
	eregi_replace = eregObj.Replace(text, replace) ' Replace String 
End Function
 
Function auto_link(Contents)
	Dim regex_file, regex_http, regex_mail
	regex_http = "(http|https|ftp|telnet|news):\/\/(([\xA1-\xFEa-z0-9_\-]+\.[][\xA1-\xFEa-z0-9:;&#@=_~%\?\/\.\,\+\-]+)(\/|[\.]*[a-z0-9]))"
	' 특수문자와 링크시 target 삭제
	Contents = eregi_replace("&(quot|gt|lt)","!$1", Contents)
	' html 사용시 Link 보호
	Contents = eregi_replace("href=""(" & regex_http & ")""[^>]*>"," href=""javascript:fnAPPpopupOuterBrowser('$2_orig://$3')"">", Contents)
	Contents = eregi_replace("(background|codebase|src)[ \n]*=[\n""' ]*(" & regex_http & ")[""']*","$1=""$3_orig://$4""",Contents)
	'링크가 안된 Url및 Email Address 자동 링크
	Contents = eregi_replace("(" & regex_http & ")" ,"$1", Contents)
	' 보호를 위해 치환된것 복구
	Contents = eregi_replace("!(quot|gt|lt)","&$1", Contents)
	Contents = eregi_replace("(http|https|ftp|telnet|news|mms)_orig","$1", Contents)
	Contents = eregi_replace("#-#","@",Contents) 
	' File Link시 Target을 삭제
	Contents = eregi_replace("(\.(" & regex_file & ")"") target=""_blank""","$1", Contents)
	auto_link = Contents
End Function
%>
<script>
<!--
function fnDelAns(ans_idx){
	if(confirm("등록된 글을 삭제하시겠습니까?")){
		document.ansfrm.aidx.value=ans_idx;
		document.ansfrm.target="FrameCKP";
		document.ansfrm.submit();
	}
}

function fnFreeBoardRelyEnd(){
	location.reload();
}

function fnAskEditEnd(){
	fnAPPParentsWinReLoad();
}

function fnFreeBoardAskDel(){
	if(confirm("등록된 글을 삭제하시겠습니까?")){
		document.askfrm.target="FrameCKP";
		document.askfrm.submit();
	}
}

function fnFreeBoardAskEnd(){
	fnAPPParentsWinReLoad();
	setTimeout(function(){
		fnAPPclosePopup();
		}, 300);
}

//-->
</script>
</head>
<body>
<div class="wrap">
	<div class="container">
		<!-- content -->
<% If olect.FOneItem.FDoc_Type="G010" Then %>
		<div class="content">
			<h1 class="hidden">공지사항</h1>
			<div class="noticeView">
				<div class="noticeTit">
					<h2><%=olect.FOneItem.FDoc_Subj%></h2>
					<span><%=FormatDate(olect.FOneItem.FDoc_Regdate,"0000.00.00")%></span>
				</div>
				<div class="noticeCont">
					<%=auto_link(olect.FOneItem.FDoc_Content)%>
				</div>
			</div>
		</div>
<% Else %>
<%
Dim oans, i
set oans = new ClecturerList
oans.FPageSize = 40
oans.FCurrPage = 1
oans.FrectDoc_Idx = vidx
oans.fnGetolectList
%>
		<div class="content">
			<h1 class="hidden">문의하기</h1>
			<div class="askView">
				<div class="askTit <% If olect.FOneItem.fdoc_ans_ox="o" Then %>flagNoti<% Else %>flagAprv<% End If %>">
					<% If olect.FOneItem.fdoc_ans_ox="o" Then %>
					<dfn>처리완료</dfn>
					<% Else %>
					<dfn>대기</dfn>
					<% End If %>
					<h2><%=olect.FOneItem.FDoc_Subj%></h2>
					<span><%=FormatDate(olect.FOneItem.FDoc_Regdate,"0000.00.00")%></span>
				</div>
				<div class="askCont">
					<%=auto_link(olect.FOneItem.FDoc_Content)%>
				</div>
				<div class="btnGroup">
					<button type="button" class="btn btnGrn" onclick="fnAPPpopupFreeBoardReply('<%=g_AdminURL%>/apps/academy/notice/popReplyWrite.asp?Doc_Id=<%=olect.FOneItem.fdoc_idx%>&mode=write');">답글</button>
					<% If olect.FOneItem.FDoc_Id=MakerID Then %>
					<button type="button" class="btn btnWht" onclick="fnAPPpopupFreeBoardAsk('<%=g_AdminURL%>/apps/academy/notice/freeboardWrite.asp?didx=<%=vidx%>&mode=edit');">수정</button>
					<button type="button" class="btn btnWht" onclick="fnFreeBoardAskDel()">삭제</button>
					<% End If %>
				</div>
				<!-- 답글 -->
				<% If oans.fresultcount > 0 Then %>
				<% For i =0 To oans.fresultcount -1 %>
				<div class="reply">
					<p class="writer"><%= getthefingers_staff("", oans.FItemList(i).fpart_sn, oans.FItemList(i).fcompany_name) %></p>
					<p class="txt"><%=replace(oans.FItemList(i).fans_content,vbCrLf,"<br>")%></p>
					<p class="txtInfo"><%=FormatDate(oans.FItemList(i).fans_regdate,"0000.00.00")%></p>
					<div class="btnGroup">
						<% If oans.FItemList(i).fid = MakerID Then %>
						<button type="button" class="btn btnWht" onclick="fnAPPpopupFreeBoardReply('<%=g_AdminURL%>/apps/academy/notice/popReplyWrite.asp?Doc_Id=<%=olect.FOneItem.FDoc_Id%>&Ans_Idx=<%=oans.FItemList(i).fans_idx%>&mode=edit');">수정</button>
						<% End If %>
						<% If oans.FItemList(i).fid = MakerID Or (fingmaster(MakerID)) Then %>
						<button type="button" class="btn btnWht" onclick="fnDelAns(<%=oans.FItemList(i).fans_idx%>);">삭제</button>
						<% End If %>
					</div>
				</div>
				<% Next %>
				<% End If %>
			</div>
		</div>
<%
Set oans = Nothing
End If
%>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
</body>
</html>
<form method="post" name="ansfrm" action="/apps/academy/notice/dofreeboardWrite.asp">
	<input type="hidden" name="aidx">
	<input type="hidden" name="del" value="o">
</form>
<form method="post" name="askfrm" action="/apps/academy/notice/lecturer_proc.asp">
	<input type="hidden" name="didx" value="<%=vidx%>">
	<input type="hidden" name="mode" value="del">
</form>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<% Set olect = Nothing %>
<!-- #include virtual="/apps/academy/lib/pms_badge_check.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->