<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  파트관리자 상/하위 카테고리 등록/수정폼
' History : 2011.01.25 김진영 생성
'####################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/partpersonCls.asp"-->
<%
Dim mode, idx, sname, sab, cc
mode = requestCheckVar(request("mode"),10)
idx = requestCheckVar(request("idx"),10)
sname = requestCheckVar(request("sname"),30)
sab = requestCheckVar(request("sab"),30)
cc = requestCheckVar(request("cc"),30)

Dim iTotCnt, arrList, intLoop, arrFileList, i
Dim iPageSize, iCurrentpage ,iDelCnt
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt, sDoc_ViewList
Dim iDoc_Idx, sDoc_Id, sDoc_Name, sDoc_Status, sDoc_Start, sDoc_End, sDoc_Type, sDoc_Import, sDoc_Diffi, sDoc_Subj, sDoc_Content
Dim sDoc_WorkerName, sDoc_Worker, sDoc_UseYN, sDoc_Regdate, sDoc_WorkerView

iDoc_Idx		= NullFillWith(requestCheckVar(Request("didx"),10),"")
iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)


If iDoc_Idx = "" Then
	sDoc_Id 		= session("ssBctId")
	sDoc_Name		= session("ssBctCname")
	sDoc_Regdate	= Left(now(),10)
	sDoc_WorkerName	= ""
	sDoc_Worker		= ""
Else
	'####### 맨 처음 확인 날짜 저장 #######
	Call WorkerView(iDoc_Idx)
	'####### 맨 처음 확인 날짜 저장 #######

	Dim cooperateView, cooperateFile
	Set cooperateView = New CCooperate
	cooperateView.FDoc_Idx = iDoc_Idx
	cooperateView.fnGetCooperateView

	sDoc_Id 		= cooperateView.FDoc_Id
	sDoc_Name		= cooperateView.FDoc_Name
	sDoc_Status		= cooperateView.FDoc_Status
	sDoc_Start		= cooperateView.FDoc_Start
	sDoc_End		= cooperateView.FDoc_End
	sDoc_Type		= cooperateView.FDoc_Type
	sDoc_Import		= cooperateView.FDoc_Import
	sDoc_Diffi		= cooperateView.FDoc_Diffi
	sDoc_Subj		= cooperateView.FDoc_Subj
	sDoc_Content	= cooperateView.FDoc_Content
	sDoc_UseYN		= cooperateView.FDoc_UseYN
	sDoc_Regdate	= cooperateView.FDoc_Regdate
	sDoc_WorkerName	= cooperateView.FDoc_WorkerName
	sDoc_Worker		= cooperateView.FDoc_Worker
	sDoc_WorkerView	= cooperateView.FDoc_WorkerViewdate

	Set cooperateFile = new CCooperate
 	cooperateFile.FDoc_Idx = iDoc_Idx
	arrFileList = cooperateFile.fnGetFileList

	For i=0 To UBOUND(Split(sDoc_WorkerName,","))
		if Not(sDoc_WorkerView="" or isNull(sDoc_WorkerView)) then
			'확인일이 없는경우 Pass (2009.06.03;허진원)
			sDoc_ViewList = sDoc_ViewList & "&nbsp;" & Split(sDoc_WorkerName,",")(i) & " : " & Split(sDoc_WorkerView,",")(i) & "<br>"
		end if
	Next
End If
%>

<script language='javascript'>
var openWorker = null;

function workerlist(k)
{
	var worker = iform.doc_worker.value;
	openWorker = window.open('member_pop.asp?worker='+worker+'&idx='+k+'','member_pop','width=570,height=527,scrollbars=yes');
}
function form_check(){
	var frm = document.iform;
	if(frm.category1.value == ""){
		alert("카테고리 이름을 입력하세요");
		frm.category1.focus();
		return;
	}
	frm.action = "partcate_proc.asp?mode=insert";
	frm.submit();
}
function modify(){
	var frm = document.iform;
	if(frm.category1.value == ""){
		alert("카테고리 이름을 입력하세요");
		frm.category1.focus();
		return;
	}
	frm.action = "partcate_proc.asp?mode=modify";
   	frm.submit();
}
function hide(k){
	var frm = document.iform;
	if(confirm("확인버튼 클릭시 하위카테고리도 \n사용하지 못 하게됩니다. \n계속 하시겠습니까?")){
		frm.action = "partcate_proc.asp?mode=hide";
   		frm.submit();
	}
}
function use(k){
	var frm = document.iform;
	if(confirm("확인버튼 클릭시 하위카테고리도 \n사용할 수 있게 됩니다. \n계속 하시겠습니까?")){
		frm.action = "partcate_proc.asp?mode=use";
   		frm.submit();
	}
}
function del(idx){
	if(confirm("확인버튼 클릭시 삭제됩니다. \n계속 하시겠습니까?")){
		var frm = document.iform;
		frm.action = "partcate_proc.asp?idx="+idx+"&mode=del";
		frm.submit();
	}
}
function form_check2(){
	var frm = document.iform;
	if(frm.category1.value == ""){
		alert("카테고리 이름을 입력하세요");
		frm.category1.focus();
		return;
	}
	if(frm.doc_workername.value == ""){
		alert("담당자를 선택하세요");
		frm.doc_workername.focus();
		return;
	}
	frm.action = "partcate_proc.asp?mode=cinsert";
	frm.submit();
}
function cmodify(idx){
	var frm = document.iform;
	if(frm.category1.value == ""){
		alert("카테고리 이름을 입력하세요");
		frm.category1.focus();
		return;
	}
	if(frm.doc_workername.value == ""){
		alert("담당자를 선택하세요");
		frm.doc_workername.focus();
		return;
	}
	frm.action = "partcate_proc.asp?mode=cmodify&cc="+ idx;
	frm.submit();
}
function member_pop(k)
{
	var windowW = 900;
	var windowH = 600;
	var left = Math.ceil( (window.screen.width  - windowW) / 2 );
	var top = Math.ceil( (window.screen.height - windowH) / 2 );
	var member_pop = window.open('./member_pop.asp?idx=' + k, 'member_pop', 'left='+ left + ',top=' + top + ',width=' + windowW + ',height=' + windowH + ',scrollbars=yes,resizable=yes');
	member_pop.focus();
}

</script>
<!-- 상카테고리 등록할 경우 -->
<% If mode = "insert" Then %>
<b><center>신규 카테고리 등록</center></b><p>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<form name="iform" method="post">
<input type=hidden name="mode" value="<%=mode%>">
<tr bgcolor="#DDDDFF">
	<td width="130" height="30">카테고리 이름</td>
	<td bgcolor="#FFFFFF" height="30"><input type="text" class="text" name="category1" value="" size="50" maxlength="50"></td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" align="center" colspan="3" height="30">
		<a href="javascript:form_check()"><img src="/images/icon_confirm.gif" border="0"></a>
	</td>
</tr>
</form>
</table>
<%End If%>

<!-- 상카테고리 수정할 경우 -->
<% If mode = "modify" Then %>
<%
Dim clist, arlist
	Set clist = new Partlist
		clist.idx = idx
		arlist = clist.fnGetmolist
	Set clist = nothing
%>
<b><center>카테고리 수정</center></b><p>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<form name="iform" method="post">
<input type=hidden name="mode" value="<%=mode%>">
<input type=hidden name="idx" value="<%=idx%>">
<tr bgcolor="#DDDDFF">
	<td width="130" height="30">카테고리 이름</td>
	<td bgcolor="#FFFFFF" height="30"><input type="text" class="text" name="category1" value="<%=arlist(1,0)%>" size="50" maxlength="50"></td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" align="center" colspan="3" height="30">
		<img src="/images/icon_modify.gif" onclick="modify();" style="cursor:hand">
	<%
		If arlist(3,0) = "Y" Then
	%>
		<img src="/images/icon_hide.gif" onclick="hide('<%=idx%>');" style="cursor:hand">
	<%
		ElseIf arlist(3,0) = "N" Then
	%>
		<img src="/images/icon_use.gif" onclick="use('<%=idx%>');" style="cursor:hand">
	<%
		End If
	%>
	</td>
</tr>
</form>
</table>
<%End If%>

<!-- 하카테고리 등록할 경우 -->
<% If mode = "cinsert" Then %>
<%
Dim plist, prlist, prlist2
	Set plist = new Partlist
		plist.idx = idx
		prlist2 = plist.fnGetmolist
	Set plist = nothing
%>
<b><center>하위 카테고리 등록</center></b><p>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<form name="iform" method="post">
<input type="hidden" name="idx" value="<%=idx%>">
<tr bgcolor="#FFFFFF">
	<td height="30" colspan="2">상위 카테고리 : <b><%= prlist2(1,0) %></b></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width="130" height="30">카테고리 이름</td>
	<td bgcolor="#FFFFFF" height="30"><input type="text" class="text" name="category1" value="" size="50" maxlength="50"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width="130" height="30">담당자</td>
	<td bgcolor="#FFFFFF" height="30">
		<input type="text" class="text" name="doc_workername" value="<%=sDoc_WorkerName%>" size="60" readonly>
		<input type="hidden" name="doc_worker" value="<%=sDoc_Worker%>">
		<input type="button" class="button" value="담당자리스트" onClick="workerlist('<%=idx%>')">
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" align="center" colspan="2" height="30">
		<a href="javascript:form_check2()"><img src="/images/icon_confirm.gif" border="0"></a>
	</td>
</tr>
</form>
</table>
<%End If%>

<!-- 하카테고리 수정할 경우 -->
<% If mode = "cmodify" Then %>
<%
Dim Fidx
Fidx = requestCheckVar(request("cc"),30)
Dim mlist, mrlist, mrlist2
	Set mlist = new Partlist
		mlist.idx = idx
		mrlist2 = mlist.fnGetmolist3
	Set mlist = nothing
%>
<b><center>하위 카테고리 수정</center></b><p>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<form name="iform" method="post">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="Fidx" value="<%=Fidx%>">
<tr bgcolor="#FFFFFF">
	<td height="30" colspan="2">상위 카테고리 : <b><%= sname %></b></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width="130" height="30">카테고리 이름</td>
	<td bgcolor="#FFFFFF" height="30"><input type="text" class="text" name="category1" value="<%= mrlist2(7,0) %>" size="50" maxlength="50"></td>
</tr>

<tr bgcolor="#DDDDFF">
	<td width="130" height="30">담당자</td>
	<td bgcolor="#FFFFFF" height="30">
		<input type="text" class="text" name="doc_workername" value=
<%
Dim x
For x = 0 to ubound(mrlist2,2)
	If x < ubound(mrlist2,2) Then
		response.write mrlist2(0,x)&","
	Else
		response.write mrlist2(0,x)
	End If
Next
%>
		  size="60" readonly>
		<input type="hidden" name="doc_worker" value=
<%
Dim y
For y = 0 to ubound(mrlist2,2)
	If y < ubound(mrlist2,2) Then
		response.write mrlist2(4,y)&","
	Else
		response.write mrlist2(4,y)
	End If
Next
%>
		>
		<input type="button" class="button" value="담당자리스트" onClick="workerlist('<%=idx%>')"><hr>
		하위카테고리와, 담당자 2가지가 삭제됩니다.
		<input type="button" class="button" value="삭제" onclick="del(<%=idx%>)">
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" align="center" colspan="2" height="30">
		<a href="javascript:cmodify('<%=cc%>')"><img src="/images/icon_confirm.gif" border="0"></a>
	</td>
</tr>
</form>
</table>
<% End If %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->