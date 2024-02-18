<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 업체게시판
' History : 2009.04.07 최초생성자모름
'			2016.04.07 한용민 수정
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/classes/board/upche_qnacls.asp" -->
<%
dim idx, mode
	idx = requestCheckvar(request("idx"),10)

if (idx = "") then
	idx = "-1"
end if

dim boardqna
set boardqna = New CUpcheQnADetail
	boardqna.FRectIdx = idx
	boardqna.read

mode = "write"
if (idx <> "") and (idx <> "-1") then
	if (boardqna.Fuserid <> "") then
		mode = "edit"

		If boardqna.Fuserid <> session("ssBctId") Then
			set boardqna = Nothing

			Response.Write "<script type='text/javascript'>alert('잘못된 접근입니다.');location.href='/';</script>"
			dbget.close()
			Response.End
		End If
	end if
end if

if (idx = "-1") then
	idx = ""
end if

%>

<STYLE TYPE="text/css">
<!--
    A:link, A:visited, A:active { text-decoration: none; }
    A:hover { text-decoration:underline; }
    BODY, TD, UL, OL, PRE { font-size: 9pt; }
    INPUT,SELECT,TEXTAREA { border:1 solid #666666; background-color: #CACACA; color: #000000; }
-->
</STYLE>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

function workerlist(){
	var worker = document.frm.workerid.value;
	window.open('PopWorkerList.asp?workerid='+worker+'&idx=<%= idx %>','workerlist','width=590,height=527,scrollbars=yes');
}

function SubmitForm(){
    if (document.frm.gubun.value == "") {
        alert("분류를 선택해주세요.");
        return;
    }
    if (document.frm.workerid.value == "") {
        alert("담당자를 선택해주세요.");
        return;
    }
	if (document.frm.title.value == "") {
        alert("제목을 입력하세요.");
        return;
    }
    if (document.frm.contents.value == "") {
        alert("내용을 입력하세요.");
        return;
    }

    if (confirm("입력이 정확합니까?") == true) {
		document.frm.submit();
	}
}

function pop_10x10_person(){
	var popwin = window.open('/common/pop_10x10_person.asp?board=U','pop_10x10_person','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function jsSetDefaultText() {
	var frm = document.frm;
	var workerid = frm.workerid.value;

	//2016.04.07 한용민 추가
	if (frm.mode.value == "write" && frm.contents.value == "") {
		var str = $.ajax({
				type: "GET",
		        url: "/common/member/ajax_getpartno.asp",
		        data: "userid="+workerid+"&empno=",
		        dataType: "text",
		        async: false
		}).responseText;

		//30:시스템운영팀, 7:시스템개발팀, 10:cs팀
		if (str=='30' || str=='7' || str=='10'){
			frm.contents.value = "[입력사항]\n주문번호 : \n주문자 : \n상품명 : \n통화여부 : \n내용 : \n";
		}
	}
}

</script>

<table width="600" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form method="post" name="frm" action="upche_qna_board_act.asp" onsubmit="return false">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="masterid" value="01">
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">질문유형</td>
	<td bgcolor="#FFFFFF">
	  <select class="select" name="gubun">
		<option value="">선택</option>
		<option value="01" <% if boardqna.Fgubun="01" then response.write "selected" %> >배송문의</option>
		<option value="02" <% if boardqna.Fgubun="02" then response.write "selected" %> >반품문의</option>
		<option value="03" <% if boardqna.Fgubun="03" then response.write "selected" %> >교환문의</option>
		<option value="04" <% if boardqna.Fgubun="04" then response.write "selected" %> >정산문의</option>
		<option value="05" <% if boardqna.Fgubun="05" then response.write "selected" %> >입고문의</option>
		<option value="06" <% if boardqna.Fgubun="06" then response.write "selected" %> >재고문의</option>
		<option value="07" <% if boardqna.Fgubun="07" then response.write "selected" %> >상품등록문의</option>
		<option value="08" <% if boardqna.Fgubun="08" then response.write "selected" %> >이벤트문의</option>
		<option value="20" <% if boardqna.Fgubun="20" then response.write "selected" %> >기타문의</option>
	  </select>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">담당자</td>
  	<td bgcolor="#FFFFFF">
      	<input type="text" class="text" name="workername" value="<%= fnGetMemberName(boardqna.Fworkerid) %>" size="10" readonly>
      	<input type="hidden" name="workerid" value="<%= boardqna.Fworkerid %>">
		<!--
      	<input type="button" value="담당자리스트OLD" onClick="workerlist()">
		-->
		<input type="button" value="담당자리스트" onClick="pop_10x10_person()"><br><br>

		* 주문 취소/반품/반품배송비 등은 고객센터 담당자를 선택하세요.<br>
		* 상품명 변경, 수수료 문의 등은 담당MD 선택하세요.<br>
		* 그밖에 누구를 담당자로 지정할지 모를 경우 고객센터 담당자를 선택하세요.
  	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">제목</td>
	<td bgcolor="#FFFFFF">
      	<input type="text" class="text" name="title" size="50" value="<%= boardqna.Ftitle %>">
  	</td>
</tr>
<tr>
	<td align="center" bgcolor="<%= adminColor("tabletop") %>">내용</td>
  	<td bgcolor="#FFFFFF">
  		<textarea class="textarea" class="textarea" name="contents" cols="80" rows="10" onFocus="jsSetDefaultText()"><%= db2html(boardqna.Fcontents) %></textarea>
  	</td>
</tr>
<tr bgcolor="#FFFFFF">
  	<td colspan="2" align="center">
  	<% if mode = "write" then %>
  	<input type="button" class="button" value="글쓰기" onclick="SubmitForm()"></td>
  	<% else %>
  	<input type="button" class="button" value="수정" onclick="SubmitForm()"></td>
  	<% end if %>
</tr>
</form>
</table>

<%
set boardqna=nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
