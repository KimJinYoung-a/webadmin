<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 1:1 상담
' History : 2009.04.17 이상구 생성
'			2016.03.25 한용민 수정(문의분야 모두 DB화 시킴)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/myqnacls.asp" -->
<%
dim i, j, reffrom, orderinfo
	reffrom = request("reffrom")

'나의 1:1질문답변
dim boardqna
set boardqna = New CMyQNA

boardqna.read(request("id"))

if boardqna.results(0).userid <> "" then
	set orderinfo = New CMyQNAOrderInfo
	'orderinfo.UserOrderInfo (boardqna.results(0).userid)
	'orderinfo.UserMinusOrderInfo (boardqna.results(0).userid)
end if

if boardqna.results(0).userid <> "" or boardqna.results(0).orderserial <> "" then
	dim myqnalist
	set myqnalist = New CMyQNA
	if boardqna.results(0).userid <> "" then
	    myqnalist.SearchUserID = boardqna.results(0).userid
	end if
	if boardqna.results(0).orderserial <> "" then
	    myqnalist.SearchOrderSerial = boardqna.results(0).orderserial
	end if

    myqnalist.PageSize = 100
    myqnalist.CurrPage = 1

	'// 예전 상담 목록
    myqnalist.list
end if

dim useridForShow : useridForShow = "고객"
if boardqna.results(0).userid <> "" then
	useridForShow = boardqna.results(0).userid
end if

dim extItemURL

%>

<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script type="text/javascript">

function SubmitForm(){
	var replyuser = "<%= session("ssBctID") %>";
	var userid = "<%= boardqna.results(0).userid %>";

	// 고객 전담자 지정
	/* 해제, skyer9, 2016-04-29
	if (replyuser != "bseo") {
		if (userid == "majorblue") {
			if (confirm("안내!!\n\n이수정 팀장님이 전담하기로 되어 있는 고객분입니다.\n\n계속 진행하시겠습니까?") != true) {
				return;
			}
		}
	}
	*/

	if (document.frm.replytitle.value == "") {
		alert("제목을 입력하세요.");
		return;
	}
	if (document.frm.replycontents.value == "") {
		alert("내용을 입력하세요.");
		return;
	}

	if (confirm("입력이 정확합니까?") == true) {
		var btnSubmit = document.getElementById('btnSubmit');
		btnSubmit.disabled = true;

		document.frm.submit();
	}
}

function updateqadiv(){
	if (confirm("수정하시겠습니까?")){
		document.updateform.submit();
	}
}

function updateitemid() {
	var itemid = document.frm.itemid.value;

	if (itemid == "") {
		alert("상품코드를 입력하세요.");
		return;
	}

	if (itemid*0 != 0) {
		alert("잘못된 상품코드입니다.");
		return;
	}

	if (confirm("상품코드를 지정하시겠습니까?")) {
		document.updateform.mode.value = "CGHITEMID";
		document.updateform.itemid.value = itemid;

		document.updateform.submit();
	}
}

function updateorderserial() {
	var orderserial = document.frm.orderserial.value;

	if (orderserial == "") {
		alert("주문번호를 입력하세요.");
		return;
	}

	if (confirm("주문번호를 지정하시겠습니까?")) {
		document.updateform.mode.value = "CGHORDSERIAL";
		document.updateform.orderserial.value = orderserial;

		document.updateform.submit();
	}
}

function delqadiv(){
	if (confirm("삭제하시겠습니까?")){
		document.delform.submit();
	}
}

function popMyQNAUTF8(idx) {
    var window_width = 600;
    var window_height = 400;
	var popwin = window.open("popMyQNAUTF8.asp?idx=" + idx,"popMyQNAUTF8","width=" + window_width + " height=" + window_height + " left=50 top=50 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

function popMyQNA_IMAGE(idx, imgidx) {
    var window_width = 1200;
    var window_height = 800;
	var popwin = window.open("popMyQNA_IMAGE.asp?idx=" + idx + "&imgidx=" + imgidx,"popMyQNA_IMAGE","width=" + window_width + " height=" + window_height + " left=50 top=50 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

function jsSetMakerID(delmakerid) {
	var frm = document.frm;

	if (delmakerid == true) {
		frm.targetMakerID.value = '';
	} else {
		if (frm.targetMakerID.value == '') {
			alert('브랜드를 입력하세요.');
			frm.targetMakerID.focus();
			return;
		}
	}

	if (confirm('저장 하시겠습니까?') == true) {
		frm.mode.value = "setmakerid";
		frm.submit();
	}
}

document.title = "1:1 상담리스트";

function resizeTextArea(textarea, textareawidth) {
	var lines = textarea.value.split("\n");

	if (lines.length < 10) {
		return;
	}

	var textareaheight = 1;
	for (x = 0; x < lines.length; x++) {
		c = lines[x].length;

		if (c >= textareawidth) {
			textareaheight += (Math.ceil(c / textareawidth) - 1);
		}
	}
	textareaheight += (lines.length - 1);

	textarea.rows = textareaheight;
}

window.onload = function() {
	if (document.getElementById("replycontents")) {
		resizeTextArea(document.getElementById("replycontents"), 90);
	}

	// 첨부파일
	var idAttachFile = document.getElementById('idAttachFile');
	if (idAttachFile && idAttachFile.style) {
		if (idAttachFile.clientWidth > 800) {
			idAttachFile.style.width = 800;
		}
	}

	var idAttachFile2 = document.getElementById('idAttachFile2');
	if (idAttachFile2 && idAttachFile2.style) {
		if (idAttachFile2.clientWidth > 800) {
			idAttachFile2.style.width = 800;
		}
	}
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<img src="/images/icon_star.gif" align="absbottom">
	    <font color="red"><strong>1:1 상담 답변</strong></font>
	</td>
</tr>
</table>

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" border="0">
<form method=post name="updateform" action="cscenter_qna_board_act.asp">
<input type="hidden" name="mode" value="CHG">
<input type="hidden" name="id" value="<% = boardqna.results(0).id %>">
<input type="hidden" name="sitename" value="<% = boardqna.results(0).Fsitename %>">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="orderserial" value="">
<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan="15">
		<img src="/images/icon_arrow_down.gif" align="absbottom">
	    <font color="red"><b>문의내용</b></font>
	    &nbsp;&nbsp;
	    질문유형수정 :
	    <% drawSelectBoxqadiv "qadiv", boardqna.results(0).qadiv, "", "Y", "N", "Y" %>

		<!-- 파트장 이상 -->
	    <input type="button" class="button" value="수정" onclick="updateqadiv();" <% if Not (C_ADMIN_AUTH or C_CSPowerUser) then %>disabled<% end if %>>
	    <% if Not (C_ADMIN_AUTH or C_CSPowerUser) then %><font color=gray>(수정불가 : 파트장문의)</font><% end if %>
	</td>
</tr>
</form>

<form method="post" name="frm" action="cscenter_qna_board_act.asp" onsubmit="return false">
<!--
<%' if boardqna.results(0).replyuser<>"" then %>
<input type="hidden" name="mode" value="reply">
<%' else %>
<input type="hidden" name="mode" value="firstreply">
<%' end if %>
-->
<input type="hidden" name="mode" value="REP">
<input type="hidden" name="id" value="<%= boardqna.results(0).id %>">
<input type="hidden" name="username" value="<%= boardqna.results(0).username %>">
<input type="hidden" name="userphone" value="<%= boardqna.results(0).userphone %>">
<input type="hidden" name="regdate" value="<%= boardqna.results(0).regdate %>">
<input type="hidden" name="title" value="<%= boardqna.results(0).title %>">
<input type="hidden" name="contents" value='<%= replace(html2db(boardqna.results(0).contents),"'","") %>'> <!-- -.- -->
<input type="hidden" name="replydate" value="<%= boardqna.results(0).replydate %>">
<input type="hidden" name="email" value="<%= Replace(boardqna.results(0).usermail, " ", "") %>">
<input type="hidden" name="emailok" value="<%= boardqna.results(0).emailok %>">
<input type="hidden" name="extsitename" value="<%= boardqna.results(0).Fextsitename %>">
<input type="hidden" name="sitename" value="<%= boardqna.results(0).Fsitename %>">
<input type="hidden" name="replyuser" value="<%= session("ssBctID") %>">
<input type="hidden" name="imsitxt">
<tr>
	<td width="90" align="center" bgcolor="#FFFFFF"><b>작성자</b></td>
	<td width="570" bgcolor="#FFFFFF">
	    <font color="#464646"><%= boardqna.results(0).username %>(<%= boardqna.results(0).userid %>/<%= boardqna.results(0).orderserial %>)</font>
	    &nbsp;&nbsp;
	    [ <font color="<%= getUserLevelColorByDate(boardqna.results(i).fUserLevel, left(boardqna.results(0).regdate,10)) %>">
		<b><%= getUserLevelStrByDate(boardqna.results(i).fUserLevel, left(boardqna.results(0).regdate,10)) %></b></font> ]
	    <%
	    	if boardqna.results(0).Frealnamecheck="Y" then
	    		Response.Write " / 실명확인회원"
	    	end if
			if boardqna.results(0).Fsitename <> "10x10" then
				response.write " / <b>" & boardqna.results(0).Fsitename & "</b>"
				if (boardqna.results(0).FuserGubun = "M") then
					response.write " / 상담사"
				else
					response.write " / 고객"
				end if
			end if
	    %>
	    <% if boardqna.results(0).userid<>"" then %>
    	    <a href="javascript:PopOrderMasterWithCallRingUserid('<%= boardqna.results(0).userid %>');"> >> [ID 로 주문검색]</a>
		<% end if %>
	</td>
	<td width="90" align="center" bgcolor="#FFFFFF"><b>문의주문번호</b></td>
	<td bgcolor="#FFFFFF">
	    <% if boardqna.results(0).orderserial<>"" then %>
    	    <a href="javascript:PopOrderMasterWithCallRingOrderserial('<%= boardqna.results(0).orderserial %>');"><%= boardqna.results(0).orderserial %> >>상세보기</a>
		<% end if %>
		<% if boardqna.results(0).orderserial = "" and boardqna.results(0).Fsitename <> "10x10" then %>
			<input type="text" class="text" name="orderserial" size="20" value="">
			<input type="button" class="button" value="저장" onclick="updateorderserial()">
		<% end if %>
	</td>
</tr>
<tr height="25">
	<td align="center" bgcolor="#FFFFFF"><b>작성일시</b></td>
	<td bgcolor="#FFFFFF"><font color="#464646"><%= boardqna.results(i).regdate %></font></td>
	<td align="center" bgcolor="#FFFFFF"><b>문의상품</b></td>
	<td bgcolor="#FFFFFF">
	    <%= boardqna.results(0).itemid %>
	    &nbsp;&nbsp;
	    <% if boardqna.results(0).itemid<>"" and boardqna.results(0).itemid>0 then %>
	    	<a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= boardqna.results(0).itemid %>" target="_blank">>>상품보기</a>
			<%
			if boardqna.results(0).Fsitename <> "10x10" then
				extItemURL = GetExtItemURL(boardqna.results(0).Fsitename, boardqna.results(0).itemid)
				if extItemURL = "" then
					rw "&nbsp;&nbsp; >> <a href='javascript:alert(""작업이전"")'>제휴몰 상품보기</a>"
				else
					rw "&nbsp;&nbsp; >> <a href='" & extItemURL & "' target=_blank>제휴몰 상품보기</a>"
				end if
			end if
			%>
	    <% elseif (boardqna.results(0).orderserial<>"") then %>
	    	<input type="text" class="text" name="itemid" size="6" value="">
	    	<input type="button" class="button" value="저장" onclick="updateitemid()">
		<% else %>
	    	<input type="text" class="text" name="itemid" size="6" value="">
	    	<input type="button" class="button" value="저장" onclick="updateitemid()">
	    <% end if %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#FFFFFF"><b>답변 예상일</b></td>
	<td bgcolor="#FFFFFF" height="25"><font color="#464646"><%= boardqna.results(0).FExpectReplyDate %></font></td>


	<td align="center" bgcolor="#FFFFFF"><b>고객 연락처</b></td>
	<td bgcolor="#FFFFFF">
	    <%= boardqna.results(0).userphone %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#FFFFFF"><b>문의제목</b></td>
	<td colspan="3" bgcolor="#FFFFFF" height="25"><font color="#464646"><%= nl2br(db2html(boardqna.results(0).title)) %></font> <input type="button" class="button" value="UTF8보기" onClick="popMyQNAUTF8(<% = boardqna.results(0).id %>)"></td>
</tr>
<tr height="25">
	<td align="center" bgcolor="#FFFFFF"><b>주문상품정보</b></td>
	<td bgcolor="#FFFFFF" colspan="3">
		<% if Not IsNull(boardqna.results(i).Fitemname) and boardqna.results(i).Fitemname <> "" then %>
			<%= boardqna.results(i).Fitemname %>
			<% if (boardqna.results(i).Fitemoption <> "0000") then %>
				<font color="blue">[<%= boardqna.results(i).Fitemoptionname %>]</font>
			<% end if %>
		<% end if %>
	</td>
</tr>
<% if Not IsNull(boardqna.results(i).Fdevice) and boardqna.results(i).Fdevice <> "" then %>
<tr height="25">
	<td align="center" bgcolor="#FFFFFF"><b>시스템 환경</b></td>
	<td bgcolor="#FFFFFF" colspan="3">
			<% If boardqna.results(i).Fdevice="P" Then %>PC : <% Else %>Mobile : <% End If %>
			<%= boardqna.results(i).FOS%>
			<% if (boardqna.results(i).FOSetc <> "") then %>
				 [<%= boardqna.results(i).FOSetc %>]
			<% end if %>
	</td>
</tr>
<% end if %>
<tr>
	<td align="center" bgcolor="#FFFFFF"><b>문의내용</b></td>
	<td colspan="3" bgcolor="#FFFFFF" height="25"><font color="#464646"><%= nl2br(db2html(boardqna.results(0).contents)) %></font></td>
</tr>
<tr>
	<td align="center" bgcolor="#FFFFFF"><b>첨부사진</b></td>
	<td colspan="3" bgcolor="#FFFFFF" height="25">
		<% if boardqna.results(0).FattachFile <> "" then %>
			<a href="javascript:popMyQNA_IMAGE(<%= boardqna.results(0).id %>, 0)"><img id="idAttachFile" src="<%= uploadUrl %><%= boardqna.results(0).FattachFile %>" border="0"></a>
		<% end if %>
		<% if boardqna.results(0).FattachFile2 <> "" then %>
			<a href="javascript:popMyQNA_IMAGE(<%= boardqna.results(0).id %>, 1)"><img id="idAttachFile2" src="<%= uploadUrl %><%= boardqna.results(0).FattachFile2 %>" border="0"></a>
		<% end if %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#FFFFFF"><b>답변만족도</b></td>
	<td colspan="3" bgcolor="#FFFFFF" height="25">
		<% if (boardqna.results(0).FEvalPoint > 0) then %>
			<% for i = 1 to boardqna.results(0).FEvalPoint %><img src="http://fiximage.10x10.co.kr/web2009/mytenbyten/star_red.gif"><% next %>
		<% end if %>
	</td>
</tr>
</table>

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" border="0">
<tr height="25" valign="top" bgcolor="<%= adminColor("tabletop") %>">
    <td colspan="4" valign="middle">
        <img src="/images/icon_arrow_down.gif" align="absbottom">
        <font color="red"><b>답변작성</b></font>
    </td>
</tr>
<tr>
    <td width="90" align="center" bgcolor="#FFFFFF" height="30">질문구분</td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input type="radio" name="replyqadiv" value="01" <% if boardqna.results(0).Freplyqadiv = "01" then response.write "checked" %> > 단순문의
		<input type="radio" name="replyqadiv" value="02" <% if boardqna.results(0).Freplyqadiv = "02" then response.write "checked" %> > 업체불만
		<input type="radio" name="replyqadiv" value="03" <% if boardqna.results(0).Freplyqadiv = "03" then response.write "checked" %> > 배송(CJ)불만
		<input type="radio" name="replyqadiv" value="10" <% if boardqna.results(0).Freplyqadiv = "10" then response.write "checked" %> > 시스템개선요청
		<input type="radio" name="replyqadiv" value="99" <% if boardqna.results(0).Freplyqadiv = "99" then response.write "checked" %> > 기타불만
	</td>
</tr>

<% if boardqna.results(0).replyuser<>"" then %>
	<tr>
	    <td align="center" bgcolor="#FFFFFF">답변제목</td>
		<td colspan="1" bgcolor="#FFFFFF"><input type="text" class="text" name="replytitle" size="55" value="<%= boardqna.results(0).replytitle %>"></td>
		<td colspan="2" bgcolor="#FFFFFF">업체 1차답변</td>
	</tr>
	<tr>
	    <td align="center" bgcolor="#FFFFFF">답변내용</td>
		<td width="800" bgcolor="#FFFFFF">
			<textarea class="textarea" name="replycontents" cols="90" rows="10"><%= db2html(boardqna.results(0).replycontents) %></textarea>
		</td>
		<td valign="top" colspan="2" bgcolor="#FFFFFF">
			브랜드 : <input type="text" class="text" name="targetMakerID" value="<%= boardqna.results(0).FtargetMakerID %>" size="20">
			<input type="button" class="button" value="저장" onClick="jsSetMakerID(false)">
			<% if boardqna.results(0).FtargetMakerID <> "" then %>
			<input type="button" class="button" value="삭제" onClick="jsSetMakerID(true)">
			<% end if %>
			<br /><br />
			선택상품 브랜드 : <%= boardqna.results(0).Fmakerid %><%= CHKIIF(boardqna.results(0).Fisupchebeasong = "Y", "<font color='red'>(업배)</font>", "")%><br /><br />
			<% if boardqna.results(0).FtargetMakerID <> "" then %>
				답변자 : <%= boardqna.results(0).Fupchereplyuser %><br />
				답변일시 : <%= boardqna.results(0).Fupchereplydate %><br /><br />
				답변내용 :<br />
				<%= nl2br(db2html(boardqna.results(0).Fupchereplycontents)) %>
			<% end if %>
		</td>
	</tr>
<% Else %>
	<tr>
	    <td align="center" bgcolor="#FFFFFF">답변제목</td>
		<td colspan="1" bgcolor="#FFFFFF">
			<input type="text" class="text" name="replytitle" value="[텐바이텐] 안녕하세요. 고객님 문의에 대해 답변드립니다." size="55">&nbsp;
			<!-- #include virtual="/cscenter/board/cs_reply_xml_selectbox.asp"-->
		</td>
		<td colspan="2" bgcolor="#FFFFFF">업체 1차답변</td>
	</tr>
	<tr>
	    <td align="center" bgcolor="#FFFFFF">답변내용</td>
		<td width="800" bgcolor="#FFFFFF"><textarea class="textarea" name="replycontents" cols="90" rows="20"></textarea></td>
		<td valign="top" colspan="2" bgcolor="#FFFFFF">
			브랜드 : <input type="text" class="text" name="targetMakerID" value="<%= boardqna.results(0).FtargetMakerID %>" size="20">
			<input type="button" class="button" value="저장" onClick="jsSetMakerID(false)">
			<% if boardqna.results(0).FtargetMakerID <> "" then %>
			<input type="button" class="button" value="삭제" onClick="jsSetMakerID(true)">
			<% end if %>
			<br /><br />
			선택상품 브랜드 : <%= boardqna.results(0).Fmakerid %><%= CHKIIF(boardqna.results(0).Fisupchebeasong = "Y", "<font color='red'>(업배)</font>", "")%><br /><br />
			<% if boardqna.results(0).FtargetMakerID <> "" then %>
				답변자 : <%= boardqna.results(0).Fupchereplyuser %><br />
				답변일시 : <%= boardqna.results(0).Fupchereplydate %><br /><br />
				답변내용 :<br />
				<%= nl2br(db2html(boardqna.results(0).Fupchereplycontents)) %>
			<% end if %>
		</td>
	</tr>
<% End If %>

<tr>
	<td colspan="15" align="center" bgcolor="#FFFFFF">
	    <input type="button" class="button" value=" 답변저장 " onclick="SubmitForm()" id="btnSubmit">
	    <input type="button" class="button" value=" 목록으로 " onclick="PopMyQnaList('', '', 'N')">
	</td>
</tr>
</form>
</table>

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		<font color="red"><strong>예전 상담 목록</strong></font>
	</td>
</tr>

<% if boardqna.results(0).userid <> "" or boardqna.results(0).orderserial <> "" then %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="60">레벨</td>
	    <td width="80">주문번호</td>
	    <td width="50">상품</td>
	    <td>제목</td>
	    <td width="200">구분</td>
	    <td width="90">작성일</td>
	    <td width="70">답변여부</td>
	    <td width="70">답변자</td>
	    <td width="70">답변일</td>
	    <td width="40">삭제</td>
	</tr>

	<% if myqnalist.ResultCount < 0 then %>
	<% else %>
		<% for i = 0 to (myqnalist.ResultCount - 1) %>
		<tr align="center" <% if (myqnalist.results(i).id <> CLng(request("id"))) then %>bgcolor="#FFFFFF"<% else %> class="tr_select" bgcolor="#AFEEEE"<% end if %>>
		    <td><b><%= myqnalist.results(i).GetUserLevelStr %></b></td>
		    <td><%= myqnalist.results(i).orderserial %></td>
		    <td><%= myqnalist.results(i).itemid %></td>
		    <td align="left"><a href="cscenter_qna_board_reply.asp?id=<%= myqnalist.results(i).id %>&reffrom=<%= reffrom %>"><%= myqnalist.results(i).title %></a></td>
		    <td>
		    	<a href="cscenter_qna_board_reply.asp?id=<%= myqnalist.results(i).id %>&reffrom=<%= reffrom %>">
		    	<%= myqnalist.results(i).fqadivname %></a>
		    </td>
		    <td><%= FormatDate(myqnalist.results(i).regdate, "0000-00-00") %></td>
		    <td><% if (myqnalist.results(i).replyuser<>"") then %>답변완료<% end if %></td>
		    <td><% if (myqnalist.results(i).replyuser<>"") then %><%= myqnalist.results(i).replyuser %><% end if %></td>
		    <td><acronym title="<%= myqnalist.results(i).replydate %>"><%= Left(myqnalist.results(i).replydate,10) %></acronym></td>
		    <td><% if (myqnalist.results(i).dispyn="N") then %><font color="red">삭제</font><% end if %></td>
		</tr>
		<% next %>
	<% end if %>
<% end if %>
</table>

<form method="post" name="delform" action="cscenter_qna_board_act.asp" onsubmit="return false">
<input type="hidden" name="id" value="<%= boardqna.results(0).id %>">
<input type="hidden" name="mode" value="del">
</form>

<iframe name="PrefaceFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
<script language="JavaScript">

function TnChangePrefaceNew(SelectGubun){
	PrefaceFrame.location.href="/cscenter/board/preface_select.asp?gubun=" + SelectGubun + "&userid=<%= boardqna.results(0).userid %>&masterid=01";
}

function TnChangeText(str){
	var basictext;
	basictext = "안녕하세요. <%= useridForShow %> 님\n"
	basictext = basictext + "텐바이텐 고객행복센터 <%= session("ssBctCname") %>입니다.\n"
	basictext = basictext + "(내용)\n"
	basictext = basictext + "만족스런답변이 되셨는지요\n\n"

	if(str == ''){
		document.frm.replycontents.value = basictext;
	}
	else{
		document.frm.replycontents.value = str;
	}
}

</script>
<iframe name="ComplimentFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
<script language="JavaScript">

document.onload = getOnload();

function getOnload() {
	<% if IsNull(boardqna.results(0).replyuser) then %>
	// 답변 템플릿 설정
	requestSelectBoxMaster();
	// 기본 인사말 설정
	TnChangePrefaceNew("<%= CHKIIF(boardqna.results(0).Fsitename <> "10x10", "55", "00") %>");
	<% end if %>
}

function fnSelectBoxDetailSelected(v) {
	TnChangePrefaceNew("<%= CHKIIF(boardqna.results(0).Fsitename <> "10x10", "55", "00") %>");
	setTimeout(function(){
		document.frm.replycontents.value = document.frm.replycontents.value.replace('네 고객님', v)
	}, 150);
}

function fnCopyToClipBoard() {
	document.frm.replycontentstr.focus();
	document.frm.replycontentstr.select();

	/*
	if (window.clipboardData && clipboardData.setData) {
		// IE
		clipboardData.setData('text', document.frm.replycontentstr.value);

		alert("복사되었습니다.");
	} else {
		alert("컨트롤C 로 복사하세요.");
	}
	*/
}

</script>

<%
set myqnalist = Nothing
set boardqna = Nothing
set orderinfo = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
