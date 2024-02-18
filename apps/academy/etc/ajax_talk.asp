<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<% 
	response.Charset="UTF-8"
	Response.ContentType="text/html;charset=UTF-8"
%>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - 응원톡 더보기"
'####################################################
' Description : 응원톡 리스트
' History : 2017-01-11 이종화 생성
'####################################################
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/chkLogin.asp"-->
<!-- #include virtual="/apps/academy/etc/talk_cls.asp" -->
<!-- #include virtual="/apps/academy/lib/pageformlib.asp"-->
<%
Dim oCheerTalk, myJob , ItsMeStr
Dim page
Dim j, lecturer_id

lecturer_id	= requestCheckVar(request.cookies("partner")("userid"),32)
'lecturer_id		= "kkomegii99"
page	= requestCheckVar(request("page"),9)
myJob			= requestCheckVar(request("gubun"),9)

IF lecturer_id = "" THEN
	Response.Write "<script language='javascript'>alert('잘못된 경로입니다.3');</script>"
	Response.Write "<script language='javascript'>location.href = '/corner/lectureDetail.asp?lecturer_id="&lecturer_id&"';</script>"
	response.end
END IF

If page = "" or page = 0 Then page = 1


SET oCheerTalk = new cCorner
	oCheerTalk.FCurrPage	= page
	oCheerTalk.FPageSize	= 12
	oCheerTalk.FRectGubun	= myJob		'L : 강사, D : 작가
	oCheerTalk.FRectParamid	= lecturer_id
	oCheerTalk.getCheerTalkCommentList()
%>
			<% If oCheerTalk.FTotalCount > 0 Then %>
				<ul class="artList">
				<% For j = 0 To oCheerTalk.FResultCount - 1 %>
					<%
						If Trim(oCheerTalk.FItemList(j).FUserid) = Trim(lecturer_id) Then
							ItsMeStr = " teacher"
						Else
							ItsMeStr = ""
						End If
					%>
					<% If lecturer_id = oCheerTalk.FItemList(j).FUserid Then  %>
					<li class='<%= Chkiif(oCheerTalk.FItemList(j).FReply_num <> 0, "reply master", "") %>'>
						<p class="writer<%= ItsMeStr %>"><%= oCheerTalk.FItemList(j).FUserid %></p>
					<% Else %>
					<li <%= Chkiif(oCheerTalk.FItemList(j).FReply_num <> 0, "class='reply'", "") %>>
						<p class="writer<%= ItsMeStr %>"><%= printUserId(oCheerTalk.FItemList(j).FUserid, 3, "*") %></p>
					<% End If %>
						<p class="txt"><%= nl2br(oCheerTalk.FItemList(j).FComment) %></p>
						<p class="txtInfo"><%= FormatDate(oCheerTalk.FItemList(j).FRegdate,"0000.00.00") %></p>
						<div class="btnGroup">
							<% If oCheerTalk.FItemList(j).FReply_num = 0 Then %>
							<button type="button" onclick="cheerTalkWrite('reply', '<%= oCheerTalk.FItemList(j).FReply_group_idx %>', '');" class="btn btnGrn">답글</button>
							<% End If %>
						<% If lecturer_id = oCheerTalk.FItemList(j).FUserid Then %>
							<button type="button" class="btn btnWht" onclick="cheerTalkWrite('edit', '<%= oCheerTalk.FItemList(j).FReply_group_idx %>','<%= oCheerTalk.FItemList(j).Fidx %>');">수정</button>
							<button type="button" class="btn btnWht" onclick="cheerTalkDel('<%= oCheerTalk.FItemList(j).Freply_group_idx %>','<%= oCheerTalk.FItemList(j).Fidx %>','<%= oCheerTalk.FItemList(j).Freply_depth %>');">삭제</button>
						<% End If %>
						</div>
					</li>
				<% Next %>
				</ul>
				<form method="post" name="delfrm">
				<input type="hidden" name="mode" value="del">
				<input type="hidden" name="idx" value="">
				<input type="hidden" name="ridx" value="">
				<input type="hidden" name="depth" value="">
				<input type="hidden" name="gubun" value="L">
				<input type="hidden" name="paramid" value="<%= lecturer_id %>">
				</form>
				<div class="paging">
					<a href="" onclick="fngopage('<%= oCheerTalk.FCurrPage %>','b','<%= lecturer_id %>'); return false;" class="btnPrev"><span>이전 페이지</span></a>
					<span><input type="number" id="textpage2" class="pageNum" maxlength = "4" onkeydown="return onlyNumber2(event, '<%= lecturer_id %>')" style='ime-mode:disabled;' value="<%= CInt(oCheerTalk.FCurrPage) %>" />		 / <%= CInt(oCheerTalk.FTotalPage) %></span>
					<a href="" onclick="fngopage('<%= oCheerTalk.FCurrPage + 1 %>','f','<%= lecturer_id %>'); return false;" class="btnNext"><span>다음 페이지</span></a>
				</div>
			<% Else %>
				<div class="noData"><span>등록된 응원톡이 없습니다.</span></div>
			<% End If %>
<% set oCheerTalk = nothing %>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->