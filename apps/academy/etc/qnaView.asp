<%@ codepage="65001" language="VBScript" %>
<% Option Explicit %>
<%
response.Charset="UTF-8"
Response.ContentType="text/html;charset=UTF-8"
%>
<%
Dim pageTitle
pageTitle="2016 The Fingers Artist Admin App - Q&A VIEW"
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/apps/academy/lib/htmllib.asp" -->
<!-- #include virtual="/apps/academy/lib/chkLogin.asp"-->
<!-- #include virtual="/apps/academy/etc/qnaCls.asp" -->
<!-- #include virtual="/apps/academy/lib/pageformlib.asp"-->
<!-- #include virtual="/apps/academy/lib/head.asp" -->
<%
Dim MakerID, oDiyItemQnAList, i, gridx

gridx = getNumeric(requestCheckVar(request("gridx"),10))
MakerID	= requestCheckVar(request.cookies("partner")("userid"),32)
'Response.write gridx &"<br>"
'Response.write MakerID &"<br>"
'Response.end
If (MakerID="") Then
	Response.Write "<script>alert('잘못된 접속 입니다.');fnAPPclosePopup();</script>"
	Response.End
End If

'상품 QnA 리스트
set oDiyItemQnAList = new DiyItemCls
oDiyItemQnAList.FPageSize = 30
oDiyItemQnAList.FCurrPage = 1
oDiyItemQnAList.FRectgroupidx = gridx
oDiyItemQnAList.FRectMakerid = MakerID
oDiyItemQnAList.FRectStateDIV = 0
oDiyItemQnAList.FRectmode = "reply"
oDiyItemQnAList.GetDiyQnaList()

'Response.write oDiyItemQnAList.FresultCount &"<br>"
'Response.write MakerID &"<br>"
'Response.end
%>
<script>
<!--
var _gridx;

function fnQnaViewRelold(ridx){
	setTimeout(function(){
		fnAPPParentsWinReLoad();
	}, 300);
	setTimeout(function(){
		fnAPPParentsWinJsCall("fnEtcMainRelold(\"\")");
	}, 600);
}

function fnRelyDel(ridx,idx){
	_gridx=ridx;
	document.rfrm.ridx.value=ridx;
	document.rfrm.idx.value=idx;
	document.rfrm.mode.value="adel";
	document.rfrm.action="/apps/academy/etc/doqnareply.asp";
	document.rfrm.target="FrameCKP";
	document.rfrm.submit();
}

function fnQnARelyEnd(msg,mode,qnacount){
	//fnAPPParentsWinReLoad();
	alert(msg);
	window.location.reload(true);
}
//-->
</script>
<div class="wrap bgGry">
	<div class="container">
		<!-- content -->
		<div class="content bgGry" id="contents">
			<h1 class="hidden">Q&amp;A</h1>
			<% if oDiyItemQnAList.FresultCount > 0 then %>
			<div class="qnaViewWrap boardList">
				<% if oDiyItemQnAList.FItemList(0).FanswerYN = "Y" then %>
				<div class="artInfo flagFinish">
				<% else %>
				<div class="artInfo flagIng">
				<% end if %>
					<a href="javascript:fnAPPpopupOuterBrowser('<%=wwwUrl%>/diyshop/shop_prd.asp?itemid=<%= oDiyItemQnAList.FItemList(0).Fitemid %>');" class="pushList">
						<div class="artThumb"><img src="<%= oDiyItemQnAList.FItemList(0).FListimage %>" alt="" onerror="this.src='http://image.thefingers.co.kr/apps/2016/thumb_default.png'" /></div>
						<p class="fs1-1r"><%= oDiyItemQnAList.FItemList(0).Fitemid %></p>
						<strong><%= oDiyItemQnAList.FItemList(0).Fitemname %></strong>
						<% if oDiyItemQnAList.FItemList(0).FanswerYN = "Y" then %>
						<dfn>답변완료</dfn>
						<% else %>
						<dfn>답변중</dfn>
						<% end if %>
					</a>
				</div>
				<!-- 타이틀 -->
				<div class="titleCont">
					<p class="title"><%= oDiyItemQnAList.FItemList(0).Ftitle %></p>
					<p class="txtInfo"><%= printUserId(oDiyItemQnAList.FItemList(0).Fuserid,2,"*") %><span>l</span><%= FormatDate(oDiyItemQnAList.FItemList(0).FRegdate,"0000.00.00") %></p>
				</div>
				<div class="viewCont">
					<% for i=0 to oDiyItemQnAList.FresultCount-1 %>
					<!-- 질문 -->
					<div <%=chkIIF(oDiyItemQnAList.FItemList(i).Fqna="Q","class='question'","class='answer'")%>>
						<span class="ico"><%= oDiyItemQnAList.FItemList(i).Fqna %></span>
						<p class="txt"><%= nl2br(oDiyItemQnAList.FItemList(i).Fcomment) %></p>
						<div class="btnGroup">
							<% if MakerID = oDiyItemQnAList.FItemList(i).Fmakerid and oDiyItemQnAList.FItemList(i).FQna = "Q" and oDiyItemQnAList.FItemList(i).Freply_num+1 >= oDiyItemQnAList.FTotalCount then %>
							<button type="button" class="btn btnGrn" onclick="fnAPPpopupQnaWrite('<%=g_AdminURL%>/apps/academy/etc/popqnaReplyWrite.asp?ridx=<%= oDiyItemQnAList.FItemList(i).Freply_group_idx %>&rowidx=<%= oDiyItemQnAList.FItemList(i).Fidx %>');">답글</button>
							<% end if %>
							<% if MakerID = oDiyItemQnAList.FItemList(i).Fmakerid and oDiyItemQnAList.FItemList(i).FQna<>"Q" and oDiyItemQnAList.FItemList(i).Freply_num+1 >= oDiyItemQnAList.FTotalCount then %>
							<button type="button" class="btn btnWht" onclick="fnAPPpopupQnaWrite('<%=g_AdminURL%>/apps/academy/etc/popqnaReplyWrite.asp?ridx=<%= oDiyItemQnAList.FItemList(i).Freply_group_idx %>&rowidx=<%= oDiyItemQnAList.FItemList(i).Fidx %>&mode=edit');">수정</button>
							<% end if %>
							<% if MakerID = oDiyItemQnAList.FItemList(i).Fmakerid and oDiyItemQnAList.FItemList(i).FQna = "A" then %>
							<button type="button" class="btn btnWht" onclick="fnRelyDel(<%= oDiyItemQnAList.FItemList(i).Freply_group_idx %>,<%= oDiyItemQnAList.FItemList(i).Fidx %>)">삭제</button>
							<% end if %>
						</div>
					</div>
					<% next %>
				</div>
			</div>
			<% end if %>
		<!--// content -->
		<div id="layerMask" class="layerMask"></div>
	</div>
</div>
<form name="rfrm" id="rfrm" method="post" style="margin:0px;">
<input type="hidden" name="mode">
<input type="hidden" name="ridx">
<input type="hidden" name="idx">
</form>
<iframe name="FrameCKP" src="about:blank" frameborder="0" width="0" height="0"></iframe>
<%
SEt oDiyItemQnAList = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->