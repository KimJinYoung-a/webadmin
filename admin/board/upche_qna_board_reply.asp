<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [업체]게시판
' Hieditor : 2015.05.27 이상구 생성
'			 2020.03.12 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/classes/board/upche_qnacls.asp" -->
<!-- #include virtual="/lib/classes/board/cs_templatecls.asp"-->
<%
dim i, j, reffrom, page, SearchKey, SearchString, gubun, replyYn, usingYn, param, sWorkerId, sWorkerName, selDate, sDate, eDate
dim idx, replytitle
	reffrom = request("reffrom")
	page = requestcheckvar(getNumeric(Request("page")),10)
	SearchKey = Request("SearchKey")
	SearchString = Request("SearchString")
	gubun = requestcheckvar(getNumeric(Request("gubun")),2)
	replyYn = requestcheckvar(Request("replyYn"),1)
	usingYn = requestcheckvar(Request("usingYn"),1)
	selDate		= requestCheckVar(Request("selDate"),1)
	sDate 		= requestCheckVar(Request("sDate"),10)
	eDate 		= requestCheckVar(Request("eDate"),10)
	idx = requestcheckvar(getNumeric(Request("idx")),10)

param = "&SearchKey=" & SearchKey & "&SearchString=" & Server.URLencode(SearchString) & "&gubun=" & gubun & "&replyYn=" & replyYn & "&usingYn=" & usingYn & "&selDate=" & selDate & "&sDate=" & sDate & "&eDate=" & eDate

'==============================================================================
dim boardqna
set boardqna = New CUpcheQnADetail
	boardqna.FRectIdx = idx
	boardqna.read()

sWorkerId = boardqna.Fworkerid
replytitle = boardqna.freplytitle
if replytitle="" or isnull(replytitle) then replytitle="안녕하세요"
If sWorkerId <> "" Then
	sWorkerName = fnGetMemberName(sWorkerId)
End IF
%>

<STYLE TYPE="text/css">
<!--
    A:link, A:visited, A:active { text-decoration: none; }
    A:hover { text-decoration:underline; }
    BODY, TD, UL, OL, PRE { font-size: 9pt; }
    INPUT,SELECT,TEXTAREA { border:1 solid #666666; background-color: #CACACA; color: #000000; }
-->
</STYLE>
<script type="text/javascript">

function workerlist(){
	var worker = document.frm.workerid.value;
	window.open('/designer/board/PopWorkerList.asp?workerid='+worker+'&idx=<%= idx %>','worker','width=590,height=527,scrollbars=yes');
}

function SubmitForm(){
	if (document.frm.workerid) {
        if (document.frm.workerid.value == "") {
                alert("담당자를 선택해주세요.");
                return;
        }
	}

    if (document.frm.replytitle.value == "") {
            alert("제목을 입력하세요.");
            return;
    }
    if (document.frm.replycontents.value == "") {
            alert("내용을 입력하세요.");
            return;
    }

    if (confirm("입력이 정확합니까?") == true) { document.frm.submit(); }
}

function updateqadiv(){
	if (confirm("수정하시겠습니까?")){
		document.updateform.submit();
	}
}

function delqadiv(){
	if (confirm("삭제하시겠습니까?")){
		document.delform.submit();
	}
}

function TnCSTemplateGubunChanged(gubun) {

	CSTemplateFrame.location.href="/cscenter/board/cs_template_select_process.asp?mastergubun=31&gubun=" + gubun;
}

function TnCSTemplateGubunProcess(v, errMSG) {

	if (errMSG != "") {
		alert(errMSG);
		return;
	}

	if(v == "") {
		//
	} else {
		document.frm.replycontents.value = v;
		// alert(v);
	}
}

function workerchange()
{
    if (document.frm.workerid.value == "") {
            alert("담당자를 선택해주세요.");
            return;
    }

	if(confirm("선택하신 담당자로 변경하시겠습니까?") == true) {
		document.frm.mode.value = "edit";
		document.frm.submit();
		return true;
     } else {
     	return false;
     }
}
</script>

<form method="post" name="frm" action="/admin/board/upche_qna_board_act.asp" onsubmit="return false" style="margin:0px;" >
<input type="hidden" name="mode" value="reply">
<input type="hidden" name="idx" value="<%= boardqna.Fidx %>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="SearchKey" value="<%=SearchKey%>">
<input type="hidden" name="SearchString" value="<%=SearchString%>">
<input type="hidden" name="gubun" value="<%=gubun%>">
<input type="hidden" name="replyYn" value="<%=replyYn%>">
<input type="hidden" name="selDate" value="<%=selDate%>">
<input type="hidden" name="sDate" value="<%=sDate%>">
<input type="hidden" name="eDate" value="<%=eDate%>">
<input type="hidden" name="imsitxt">
<table width="800" align="left" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan="15">
		<img src="/images/icon_arrow_down.gif" align="absbottom">
		<font color="red"><b>문의내용</b></font>

	</td>
</tr>
<tr height="25">
	<td width="80" align="center" bgcolor="#FFFFFF"><b>작성자</b></td>
	<td colspan="3" bgcolor="#FFFFFF"><%= boardqna.Fusername %>(<%= boardqna.fuserid %>)</td>
</tr>
<tr height="25">
	<td align="center" bgcolor="#FFFFFF"><b>작성일시</b></td>
	<td colspan="3" bgcolor="#FFFFFF"><%= boardqna.Fregdate %></td>
</tr>
<tr height="25">
	<td align="center" bgcolor="#FFFFFF"><b>문의제목</b></td>
	<td colspan="3" bgcolor="#FFFFFF"><%= nl2br(ReplaceBracket(boardqna.Ftitle)) %></td>
</tr>
<tr height="25">
	<td align="center" bgcolor="#FFFFFF"><b>문의내용</b></td>
	<td colspan="3" bgcolor="#FFFFFF"><%= nl2br(ReplaceBracket(boardqna.Fcontents)) %></td>
</tr>

<tr height="25" valign="top" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan="4" valign="middle">
		<img src="/images/icon_arrow_down.gif" align="absbottom">
		<font color="red"><b>답변작성</b></font>
	</td>
</tr>

<% if not isnull(boardqna.Freplyuser) then %>
	<tr height="25">
		<td align="center" bgcolor="#FFFFFF"><b>답변자</b></td>
		<td colspan="3" bgcolor="#FFFFFF"><%= boardqna.Freplyuser %></td>
	</tr>
	<tr height="25">
		<td align="center" bgcolor="#FFFFFF"><b>답변일시</b></td>
		<td colspan="3" bgcolor="#FFFFFF"><%= boardqna.Freplydate %></td>
	</tr>
<% else %>
	<tr height="25">
		<td align="center" bgcolor="#FFFFFF"><b>담당자</b></td>
		<td colspan="3" bgcolor="#FFFFFF">
			<input type="text" class="text_ro" name="workername" value="<%=sWorkerName%>" size="10" readonly>
			<input type="hidden" name="workerid" value="<%=sWorkerId%>">
			<input type="button" class="button_s" value="담당자리스트" onClick="workerlist()">
			<input type="button" class="button_s" value="담당자변경하기" onClick="workerchange()">
		</td>
	</tr>
<% end if %>

<tr>
	<td align="center" bgcolor="#FFFFFF"><b>답변제목</b></td>
	<td colspan="3" bgcolor="#FFFFFF">
		<input type="text" class="text" name="replytitle" size="30" value="<%= replytitle %>">&nbsp;

		<% if boardqna.freplycontents = "" then  %>
			<% SelectBoxCSTemplateGubunNew "31", "csreg_template", "" %>
			*[CS]각종설정 >> [업체게시판]템플릿관리 참조
			<iframe name="CSTemplateFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
		<% end if  %>
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#FFFFFF"><b>답변내용</b></td>
	<td colspan="3" bgcolor="#FFFFFF"><textarea class="textarea" name="replycontents" cols="80" rows="10"><%= (boardqna.freplycontents) %></textarea></td>
</tr>
<tr>
	<td colspan="15" align="center" bgcolor="#FFFFFF">
		<% if boardqna.fisusing="Y" then %>
			답변이메일발송
			<input type="checkbox" name="csmailsend" >
			<input type="button" class="button" value="답변저장" onclick="SubmitForm()">
		<% end if %>

		<input type="button" class="button" value="목록으로" onclick="window.location='upche_qna_board_list.asp?page=<%=page & Param%>'">
	</td>
</tr>
</table>
</form>

<iframe name="PrefaceFrame" src="" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
<script type="text/JavaScript">

function TnChangePreface(SelectGubun){
	PrefaceFrame.location.href="/cscenter/board/preface_select.asp?gubun=" + SelectGubun + "&userid=<%= boardqna.fuserid %>&masterid=03";
}

 function TnChangeText(str){
	var basictext;
	basictext = "";
	//basictext = "안녕하세요?\n텐바이텐 <%= session("ssBctCname") %> 입니다.\n\n\n수고하세요~\n궁금하신 사항은 아래로 연락주세요^^\n";
	//basictext = basictext + "담당 : <%= session("ssBctCname") %>";

	if(str == ''){
		document.frm.replycontents.value = basictext;
	} else {
		document.frm.replycontents.value = str;
	}
 }

<% if boardqna.freplycontents="" or isNull(boardqna.freplycontents) then %>
	TnChangeText('');
<% end if %>

</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
