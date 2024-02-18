<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/myqnaCls.asp"-->
<%
'####################################################
' Description :  1:1 상담관리 상세
' History : 2016.07.28 김진영 생성
' History : 2016-11-08 이종화 추가
'####################################################
%>
<%
Dim oMyqna, i, idx, gridx, reidx
Dim masterQState, masterQGubun, masterQRegID, masterQRegName, masterQEmail, masterQRegdate, masterQLastRegdate, masterQPhoneChk, masterQPhoneNumber, masterQTitle, masterQSmsOK
Dim masterQorderserial , masterQq_itemid , masterQq_itemoption , masterQItemNames , masterQitemcount ,masterQitemoptionname , masterQtotalsum, masterQitemcost

Dim regIDnName
idx		= getNumeric(requestCheckVar(request("idx"),9))
gridx	= getNumeric(requestCheckVar(request("gridx"),9))

SET oMyqna = new CQna
	oMyqna.FRectIdx = idx
	oMyqna.FRectGroupIdx = gridx
	oMyqna.getOnemyqna

	If oMyqna.FResultCount < 1 Then
		response.write "<script>alert('오류가 발생했습니다.');location.replace('/academy/board/myqnaList.asp?menupos="&menupos&"');</script>"
		response.end
	End If

	masterQState		= oMyqna.FOneItem.getAnswerName
	masterQGubun		= oMyqna.FOneItem.FLecture_gubun
	masterQRegID		= oMyqna.FOneItem.FUserid
	masterQRegdate		= oMyqna.FOneItem.FRegdate
	masterQLastRegdate	= oMyqna.FOneItem.FLastRegdate
	masterQPhoneNumber	= oMyqna.FOneItem.FSmsnum & " (답변수신)"
	masterQSmsOK		= oMyqna.FOneItem.FSmsok
	masterQTitle		= oMyqna.FOneItem.FTitle

	masterQorderserial		=	oMyqna.FOneItem.Forderserial
	masterQq_itemid			=	oMyqna.FOneItem.Fq_itemid	
	masterQq_itemoption		=	oMyqna.FOneItem.Fq_itemoption
	masterQItemNames		=	oMyqna.FOneItem.FItemNames
	masterQitemcount		=	oMyqna.FOneItem.Fitemcount
	masterQitemoptionname	=	oMyqna.FOneItem.Fitemoptionname
	masterQtotalsum			=	oMyqna.FOneItem.Ftotalsum
	masterQitemcost			=	oMyqna.FOneItem.Fitemcost

	Dim GetItemNames
	if (masterQitemcount>1) then
		GetItemNames = masterQItemNames + " 외 " + CStr(masterQitemcount-1) + "건"
	else
		GetItemNames = masterQItemNames
	end If
	
	Call getMyinfo(masterQRegID, masterQRegName, masterQEmail)
SET oMyqna = nothing
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script>
function chggubun(){
	var frm = document.frm;
	if(confirm("문의분야를 변경하시겠습니까?")){
		frm.mode.value = "C";
		frm.submit();
	}
}
function fnqnaDel(){
	var frm = document.frm;
	if(confirm("문의글을 삭제하시겠습니까?")){
		frm.mode.value = "D";
		frm.submit();
	}
}
function goView(vidx, vgridx){
	location.href='/academy/board/myqnaView.asp?menupos=<%=menupos%>&idx='+vidx+'&gridx='+vgridx;	
}
// 답변 머릿말 넣기
function chgCont(qcd, ccd, regid){
	var reStr;
	var rstStr = $.ajax({
		type: "POST",
		url: "ajax_myqnaTextarea.asp",
		data: "groupcd="+qcd+"&commcd="+ccd+"&regid="+regid,
		dataType: "text",
		async: false
	}).responseText;
	reStr = rstStr.split("|");
	if(reStr[0]=="OK"){
		$("#ansContents").val(reStr[1]);
	}
}
// 답변 수정시 머릿말 넣기
function chgContEdit(qcd, ccd, regid){
	var reStr;
	var rstStr = $.ajax({
		type: "POST",
		url: "ajax_myqnaTextarea.asp",
		data: "groupcd="+qcd+"&commcd="+ccd+"&regid="+regid,
		dataType: "text",
		async: false
	}).responseText;
	reStr = rstStr.split("|");
	if(reStr[0]=="OK"){
		$("#ansContentsEdit").val(reStr[1]);
	}
}
// 답변 등록
function fnQnareplyAdd(){
	var frm = document.replyfrm;
	if(confirm("답변글을 등록하시겠습니까?")){
		frm.mode.value = "addreply";
		frm.submit();
	}
}
// 답변 삭제
function fnQnareplyDel(vidx){
	var frm = document.frm;
	if(confirm("답변글을 삭제하시겠습니까?")){
		frm.mode.value = "adel";
		frm.reidx.value = vidx;
		frm.submit();
	}
}
// 답변글 수정시 폼변경
function fnQnareplyEditForm(vidx, commid){
	var editTrid = "QnAList"+vidx;
	var commVal = $("#"+commid+"").html();
	var repComm;
	repComm = commVal.replace(/<BR>/gi, "\n")
	$("#"+editTrid+"").hide();
	$("#replyEditTBL").show();
	$("#editidx").val(vidx);
	$("#ansContentsEdit").val(repComm);
}
// 답변글 수정
function fnQnareplyEdit(){
	var frm = document.replyEditfrm;
	if(confirm("답변글을 수정하시겠습니까?")){
		frm.mode.value = "edit";
		frm.submit();
	}
}
</script>
<!-- ########################################### 마스터 정보 시작 ########################################### -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="POST" action="/academy/board/doMyQnaProc.asp">
<input type="hidden" name="mode" >
<input type="hidden" name="idx" value="<%= idx %>" >
<input type="hidden" name="reidx" value="<%= reidx %>" >
<input type="hidden" name="gridx" value="<%= gridx %>" >
<input type="hidden" name="menupos" value="<%= menupos %>" >
<col width="15%" />
<col width="35%" />
<col width="15%" />
<col width="35%" />
<tr align="center" bgcolor="#FFFFFF" height="35">
	<td align="left" bgcolor="<%= adminColor("gray") %>">상태</td>
	<td align="left"><%= masterQState %></td>
	<td align="left" bgcolor="<%= adminColor("gray") %>">문의분야</td>
	<td align="left">
		<select class="select" name="gubunVal">
			<option value="1" <%= Chkiif(masterQGubun = "1", "selected", "") %> >작품(상품) 주문/결제</option>
			<option value="2" <%= Chkiif(masterQGubun = "2", "selected", "") %>>주문 취소/반품/교환</option>
			<option value="3" <%= Chkiif(masterQGubun = "3", "selected", "") %>>작품 배송 관련 문의</option>
			<option value="4" <%= Chkiif(masterQGubun = "4", "selected", "") %>>수강신청/결제 문의</option>
			<option value="5" <%= Chkiif(masterQGubun = "5", "selected", "") %>>수강 취소</option>
			<option value="6" <%= Chkiif(masterQGubun = "6", "selected", "") %>>개인정보 관련 문의</option>
			<option value="7" <%= Chkiif(masterQGubun = "7", "selected", "") %>>이벤트/쿠폰/마일리지</option>
			<option value="8" <%= Chkiif(masterQGubun = "8", "selected", "") %>>회원탈퇴/재가입</option>
			<option value="9" <%= Chkiif(masterQGubun = "9", "selected", "") %>>기타 문의</option>
		</select>
		<input type="button" class="button" value="구분변경" onclick="chggubun();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="35">
	<td align="left" bgcolor="<%= adminColor("gray") %>">작성자</td>
	<td align="left"><%= masterQRegName %>(<%= masterQRegID %>)</td>
	<td align="left" bgcolor="<%= adminColor("gray") %>">작성일(최종갱신)</td>
	<td align="left"><%= masterQRegdate %>&nbsp;(<%= masterQLastRegdate %>) </td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="35">
	<td align="left" bgcolor="<%= adminColor("gray") %>">이메일</td>
	<td align="left"><%= Chkiif(masterQEmail <> "", masterQEmail & " (답변수신)", "") %></td>
	<td align="left" bgcolor="<%= adminColor("gray") %>">휴대폰</td>
	<td align="left"><%= Chkiif(masterQSmsOK = "Y", masterQPhoneNumber, "") %></td>
</tr>
<% If masterQorderserial <> "" Then %>
<tr align="center" bgcolor="#FFFFFF" height="35">
	<td align="left" bgcolor="<%= adminColor("gray") %>">주문 및 수강번호</td>
	<td align="left" colspan="3"><%=masterQorderserial %> / <%=GetItemNames%> / <%= masterQtotalsum %>원</td>
</tr>
<% End If %>
<% If masterQq_itemid <> "" Then %>
<tr align="center" bgcolor="#FFFFFF" height="35">
	<td align="left" bgcolor="<%= adminColor("gray") %>">작품 코드</td>
	<td align="left" colspan="3"><a href="<%=wwwFingers%>/diyshop/shop_prd.asp?itemid=<%=masterQq_itemid %>" target="_blank"><%=masterQq_itemid %> / <%=masterQItemNames%>[<%=masterQitemoptionname%>] / <%= masterQitemcost %>원</a></td>
</tr>
<% End If %>
<tr align="center" bgcolor="#FFFFFF" height="35">
	<td align="left" bgcolor="<%= adminColor("gray") %>">제목</td>
	<td align="left"><%= masterQTitle %></td>
	<td align="center" colspan="2">
		<input type="button" class="button" value="삭제" onclick="fnqnaDel();" style=color:red;font-weight:bold>
	</td>
</tr>
</form>
</table>
<br>
<!-- ############################################ 마스터 정보 끝 ############################################ -->
<!-- ########################################### 디테일 정보 시작 ########################################### -->
<%
Dim lastqna, qstContents, lastRegdate, lastSMSok, lastSmsNum
Dim QnaColor
SET oMyqna = new CQna
	oMyqna.FCurrPage = 1
	oMyqna.FPageSize = 500
	oMyqna.FRectGroupIdx = gridx
	oMyqna.getqnaDetailList
%>
<% If oMyqna.FResultCount > 0 Then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<%
	For i = 0 to oMyqna.FResultCount - 1 
		If oMyqna.FItemList(i).FQna = "Q" Then
			QnaColor = "<font size='4' color='RED'><strong>"&oMyqna.FItemList(i).FQna&".</strong></font>"
		Else
			QnaColor = "<font size='4' color='BLUE'><strong>"&oMyqna.FItemList(i).FQna&".</strong></font>"
		End IF
%>
<tr align="LEFT" bgcolor="#FFFFFF" height="35" id="QnAList<%= oMyqna.FItemList(i).Fidx %>">
	<td><%= QnaColor %><br>
		<span id="QnAComm<%= oMyqna.FItemList(i).Fidx %>"><%= nl2br(oMyqna.FItemList(i).Fcomment) %></span>
	<% If oMyqna.FItemList(i).FanswerYN ="Y" and oMyqna.FItemList(i).Freply_num+1 >= oMyqna.FTotalCount AND oMyqna.FItemList(i).FQna = "A" Then %>
		<br><button type="button" onclick="fnQnareplyEditForm('<%= oMyqna.FItemList(i).Fidx %>', 'QnAComm<%= oMyqna.FItemList(i).Fidx %>');" class="button">수정</button>
		&nbsp;<button type="button" onclick="fnQnareplyDel('<%= oMyqna.FItemList(i).Fidx %>');" class="button">삭제</button>
	<% End If %>
	</td>
</tr>
<% 
		lastqna			= oMyqna.FItemList(i).FQna 
		If lastqna = "Q" Then
			qstContents		= oMyqna.FItemList(i).Fcomment
			lastRegdate		= oMyqna.FItemList(i).FRegdate
			lastSMSok		= oMyqna.FItemList(i).FSmsok
			lastSmsNum		= oMyqna.FItemList(i).FSmsnum
		End If
%>
<%	Next %>
</table>
<br>
<% End If %>
<!-- ########################################### 디테일 정보 끝 ########################################### -->
<!-- ################################### 질문글 일 때 답변 등록 폼 시작 ################################### -->
<% If lastqna = "Q" Then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="replyfrm" method="POST" action="/academy/board/doMyQnaProc.asp">
<input type="hidden" name="mode" >
<input type="hidden" name="idx" value="<%= idx %>" >
<input type="hidden" name="gridx" value="<%= gridx %>" >
<input type="hidden" name="menupos" value="<%= menupos %>" >
<!-- 메일에 필요한 내용 hidden 처리 -->
<input type="hidden" name="usermail" value="<%= masterQEmail %>" >
<input type="hidden" name="qstContents" value="<%= qstContents %>" >
<input type="hidden" name="lastRegdate" value="<%= lastRegdate %>" >
<input type="hidden" name="masterQRegName" value="<%= masterQRegName %>" >
<input type="hidden" name="masterQTitle" value="<%= masterQTitle %>" >
<!-- ################################-->
<!-- SMS전송에 필요한 내용 hidden 처리 -->
<input type="hidden" name="lastSMSok" value="<%= lastSMSok %>" >
<input type="hidden" name="lastSmsNum" value="<%= lastSmsNum %>" >
<!-- ################################-->
<tr align="LEFT" bgcolor="#FFFFFF" height="35">
	<td>
		<font size='4' color='BLUE'><strong>A.</strong></font><br />
		머릿말
		<select name="preface" id="preface" class="select" onchange="chgCont(this.value, compliment.value, '<%=masterQRegID%>')">
			<%= oMyqna.optPrfCd("'A000'", "H999")%>
		</select>
		/ 인사말
		<select name="compliment" id="compliment" class="select" onchange="chgCont(preface.value, this.value, '<%=masterQRegID%>')">
			<option value="">선택</option>
			<%= oMyqna.optCommCd("'E000'", "")%>
		</select>
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan="3"><textarea name="ansContents" class="textarea" id="ansContents" rows="20" cols="100"></textarea>
		&nbsp;<input type="button" value="답변하기" class="button" onclick="fnQnareplyAdd();">
	</td>
</tr>
</form>
</table>
<% End If %>
<!-- #################################### 질문글 일 때 답변 등록 폼 끝 ##################################### -->
<!-- ################################### 답변 수정 클릭시 나오는 폼 시작 ################################### -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" id="replyEditTBL" style="display:none;">
<form name="replyEditfrm" method="POST" action="/academy/board/doMyQnaProc.asp">
<input type="hidden" name="mode" >
<input type="hidden" name="idx" value="<%= idx %>" >
<input type="hidden" name="reidx" id="editidx" >
<input type="hidden" name="gridx" value="<%= gridx %>" >
<input type="hidden" name="menupos" value="<%= menupos %>" >
<!-- 메일에 필요한 내용 hidden 처리 -->
<input type="hidden" name="usermail" value="<%= masterQEmail %>" >
<input type="hidden" name="qstContents" value="<%= qstContents %>" >
<input type="hidden" name="lastRegdate" value="<%= lastRegdate %>" >
<input type="hidden" name="masterQRegName" value="<%= masterQRegName %>" >
<input type="hidden" name="masterQTitle" value="<%= masterQTitle %>" >
<!-- ################################-->
<!-- SMS전송에 필요한 내용 hidden 처리 -->
<input type="hidden" name="lastSMSok" value="<%= lastSMSok %>" >
<input type="hidden" name="lastSmsNum" value="<%= lastSmsNum %>" >
<!-- ################################-->
<tr align="LEFT" bgcolor="#FFFFFF" height="35">
	<td>
		<font style=font-weight:bold>A.</font><br />
		머릿말
		<select name="preface" id="preface" class="select" onchange="chgContEdit(this.value, compliment.value, '<%=masterQRegID%>')">
			<%= oMyqna.optPrfCd("'A000'", "H999")%>
		</select>
		/ 인사말
		<select name="compliment" id="compliment" class="select" onchange="chgContEdit(preface.value, this.value, '<%=masterQRegID%>')">
			<option value="">선택</option>
			<%= oMyqna.optCommCd("'E000'", "")%>
		</select>
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan="3"><textarea name="ansContentsEdit" class="textarea" id="ansContentsEdit" rows="20" cols="100"></textarea><input type="button" value="답변하기" class="button" onclick="fnQnareplyEdit();"></td>
</tr>
</form>
</table>
<% SET oMyqna = nothing %>
<!-- ################################### 답변 수정 클릭시 나오는 폼 끝 ################################### -->
<!-- ######################################## 지난 문의 목록 시작 ######################################## -->
<%
SET oMyqna = new CQna
	oMyqna.FCurrPage = 1
	oMyqna.FPageSize = 200
	oMyqna.FRectUserid = masterQRegID
	oMyqna.getUserQnAList

If oMyqna.FResultCount > 0 Then
%>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="35" align="left" bgcolor="BLACK">
	<td colspan="6"><font color="WHITE"><%=masterQRegID%> 회원의 지난 문의 목록</font></td>
</tr>
<tr height="35" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80">번호</td>
	<td width="250">문의분야</td>
	<td width="80">상태</td>
	<td>제목</td>
	<td width="140">등록자</td>
	<td width="140">등록일</td>
</tr>
<% For i=0 to oMyqna.FResultCount - 1 %>
<tr height="30" style="cursor:pointer;" align="center" bgcolor='#FFFFFF'" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background='ffffff'; onclick="goView('<%= oMyqna.FItemList(i).FIdx %>','<%= oMyqna.FItemList(i).FReply_group_idx %>')">
	<td align="center"><%= oMyqna.FItemList(i).FIdx %></td>
	<td align="center"><%= oMyqna.FItemList(i).getQnaGubunName %></td>
	<td align="center"><%= oMyqna.FItemList(i).getAnswerName %></td>
	<td align="left"><%= oMyqna.FItemList(i).FTitle %></td>
	<td align="center"><%= oMyqna.FItemList(i).FUserid %></td>
	<td align="center"><%= FormatDate(oMyqna.FItemList(i).FRegdate,"0000.00.00") %></td>
</tr>
<% Next %>
</table>
<% End If %>
<% SET oMyqna = nothing %>
<!-- ################################### 지난 문의 목록 끝 ################################### -->
<script>
$(function(){
	chgCont("H999", "<%=masterQRegID%>")
});
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->