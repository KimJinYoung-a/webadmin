<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 공모전리스트
' History : 이상구 생성
'			한용민 수정(isms취약점조치)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/contestCls.asp"-->

<%
	Dim contestdetail, vSubject, vContest, vEntrySDate, vEntryEDate, vVoteSDate, vVoteEDate, vResultDate, vUseYN, vRegdate
	vContest = requestCheckVar(Request("contest"),6)
	vUseYN	 = "y"
	
	If vContest <> "" Then
		Set contestdetail = new ClsContest
		contestdetail.FContest = vContest
		contestdetail.FContestDetail
		
		vContest 		= contestdetail.FOneItem.fcontest
		vSubject 		= ReplaceBracket(contestdetail.FOneItem.fsubject)
		vEntrySDate 	= contestdetail.FOneItem.fentry_sdate
		vEntryEDate 	= contestdetail.FOneItem.fentry_edate
		vVoteSDate 		= contestdetail.FOneItem.fvote_sdate
		vVoteEDate 		= contestdetail.FOneItem.fvote_edate
		vResultDate 	= contestdetail.FOneItem.fresult_date
		vUseYN 			= contestdetail.FOneItem.fuseyn
		vRegdate 		= contestdetail.FOneItem.fregdate
		Set contestdetail = nothing
	End If
%>

<script type='text/javascript'>
function PopMenuHelp(menupos){
	var popwin = window.open('/designer/menu/help.asp?menupos=' + menupos,'PopMenuHelp_a','width=900, height=600, scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<form name="frm" action="contest_proc.asp" method="post">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% If vContest <> "" Then %>
<tr align="center" bgcolor="#E6E6E6" height="30">
	<td width="100">공모전 No.</td>
	<td width="280" bgcolor="#FFFFFF" align="left"><%=vSubject%><input type="hidden" name="contest" value="<%=vContest%>"></td>
</tr>
<% End If %>
<tr align="center" bgcolor="#E6E6E6" height="30">
	<td width="100">주 제</td>
	<td width="280" bgcolor="#FFFFFF" align="left"><input type="text" name="subject" value="<%=vSubject%>" size="38" maxlength="50"></td>
</tr>
<tr align="center" bgcolor="#E6E6E6" height="30">
	<td>응모기간</td>
	<td bgcolor="#FFFFFF" align="left">
        <input id="entry_sdate" name="entry_sdate" value="<%=vEntrySDate%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="entry_sdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
        <input id="entry_edate" name="entry_edate" value="<%=vEntryEDate%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="entry_edate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var ENT_Start = new Calendar({
				inputField : "entry_sdate", trigger    : "entry_sdate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					ENT_End.args.min = date;
					ENT_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var ENT_End = new Calendar({
				inputField : "entry_edate", trigger    : "entry_edate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					ENT_Start.args.max = date;
					ENT_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
</tr>
<tr align="center" bgcolor="#E6E6E6" height="30">
	<td>고객투표기간</td>
	<td bgcolor="#FFFFFF" align="left">
        <input id="vote_sdate" name="vote_sdate" value="<%=vVoteSDate%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="vote_sdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
        <input id="vote_edate" name="vote_edate" value="<%=vVoteEDate%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="vote_edate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var VOT_Start = new Calendar({
				inputField : "vote_sdate", trigger    : "vote_sdate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					VOT_End.args.min = date;
					VOT_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var VOT_End = new Calendar({
				inputField : "vote_edate", trigger    : "vote_edate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					VOT_Start.args.max = date;
					VOT_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
</tr>
<tr align="center" bgcolor="#E6E6E6" height="30">
	<td>당선자발표일</td>
	<td bgcolor="#FFFFFF" align="left">
        <input id="result_date" name="result_date" value="<%=vResultDate%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="result_date_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var RST_Start = new Calendar({
				inputField : "result_date", trigger    : "result_date_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
</tr>
<tr align="center" bgcolor="#E6E6E6" height="30">
	<td width="100">사용여부</td>
	<td width="280" bgcolor="#FFFFFF" align="left">
		<input type="radio" name="useyn" value="y" <% If vUseYN = "y" Then Response.Write "checked" End If %>>사용&nbsp;&nbsp;&nbsp;
		<input type="radio" name="useyn" value="n" <% If vUseYN = "n" Then Response.Write "checked" End If %>>사용안함
	</td>
</tr>
<% If vContest <> "" Then %>
<tr align="center" bgcolor="#E6E6E6" height="30">
	<td width="100">등록일</td>
	<td width="280" bgcolor="#FFFFFF" align="left"><%=vRegdate%></td>
</tr>
<% End If %>
</table>
<table width="380" cellpadding="0" cellspacing="0">
<tr height="30">
	<td align="right">
		<input type="button" class="button" value="저장" onclick="frm.submit();">
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->