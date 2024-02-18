<%@ language=vbscript %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/between/betsearchcls.asp" -->
<%
Dim lWord, page, i, isusing, research
page	 = request("page")
isusing	= requestCheckvar(request("isusing"),1)
research = request("research")
If page = "" Then page = 1

If research = "" Then
	isusing = "Y"
End If

SET lWord = new cSearch
	lWord.FPageSize = 20
	lWord.FCurrPage = page
	lWord.FRectIsusing = isusing
	lWord.getLikeWordList
%>
<script type="text/javascript">
function popRegWord(idx){
    var pword = window.open("/admin/etc/between/search/popRegWord.asp?idx="+idx,"popOptionAddPrc","width=400,height=300,scrollbars=yes,resizable=yes");
	pword.focus();
}
function seach_check(){
	var sform = document.fsearch;
	sform.submit();
}
function gosubmit(page){
    var frm = document.fsearch;
    frm.page.value=page;
	frm.submit();
}
function RefreshCaFavKeyWordRec(term){
	var frm;
	frm = document.frmSvArr;

	var chkSel=0;
	var sValue;
	sValue = "";
	try {
		if(frm.cksel.length>1) {
			for(var i=0;i<frm.cksel.length;i++) {
				if(frm.cksel[i].checked) chkSel++;
			}
		} else {
			if(frm.cksel.checked) chkSel++;
		}
		if(chkSel< 5) {
			alert("선택한 검색어가 5개보다 작습니다.");
			return;
		}
	}
	catch(e) {
		alert("검색어가 없습니다.");
		return;
	}

	if (frm.cksel.length > 1){
		for (var i=0;i<frm.cksel.length;i++){
			if (frm.cksel[i].checked){
				sValue = sValue + frm.cksel[i].value + ",";
			}
		}
	}else{
		if (frm.cksel.checked){
			sValue = sValue + frm.cksel.value + ",";
		}
	}
	var AssignReal;
	AssignReal = window.open("<%=mobileUrl%>/chtml/between/make_likeword_xml.asp?idxarr="+sValue, "AssignReal","width=400,height=300,scrollbars=yes,resizable=yes");
	AssignReal.focus();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="fsearch" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		사용유무 :
		<select name="isusing" class="select">
			<option value="">-Choice-</option>
			<option value="Y" <%= Chkiif(isusing = "Y", "selected", "") %>>Y</option>
			<option value="N" <%= Chkiif(isusing = "N", "selected", "") %>>N</option>
		</select>
	</td>
	<td align="right" width="50">
		<img src="/admin/images/search2.gif" border="0" align="absmiddle" style="cursor:hand" onclick="seach_check()">&nbsp;&nbsp;
	</td>
</tr>
</form>
</table>
<% If isusing = "Y" Then %>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td><a href="javascript:RefreshCaFavKeyWordRec();"><img src="/images/icon_reload.gif" align="absmiddle" border="0" alt="html만들기"></a>Real 적용</a></td>
</tr>
</table>
<% End If %>
<table width="30" cellpadding="3" cellspacing="1" class="a">
<tr height="30">
	<td align="left">
		<input type="button" value="등록" class="button_s" onclick="popRegWord('');">
	</td>
</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmSvArr" method="post" onSubmit="return false;" action="" style="margin:0px;">
<tr align="center" bgcolor="#F3F3FF" height="30">
	<td width="30"><input type="checkbox" name="chkAll" onClick="fnCheckAll(this.checked,frmSvArr.cksel);"></td>
	<td width="200">마지막 Real 적용시간</td>
	<td width="60">순서</td>
	<td>추천 검색어</td>
	<td width="150">등록일</td>
	<td width="100">사용유무</td>
</tr>
<%
If lWord.FResultcount > 0 Then
	For i = 0 to lWord.FResultcount -1
%>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" value="<%= lWord.FItemList(i).FIdx %>"></td>
	<td><%= lWord.FItemList(i).FUpdatedate %></td>
	<td><%= lWord.FItemList(i).FRank %></td>
	<td onclick="popRegWord(<%= lWord.FItemList(i).FIdx %>);" style="cursor:pointer;"><%= lWord.FItemList(i).FLikeWord %></td>
	<td><%= lWord.FItemList(i).FRegdate %></td>
	<td><%= lWord.FItemList(i).FIsusing %></td>
</tr>
<% Next %>
<tr height="25" bgcolor="FFFFFF" >
	<td colspan="15" align="center">
       	<% If lWord.HasPreScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= lWord.StartScrollPage-1 %>');">[pre]</a></span>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + lWord.StartScrollPage to lWord.StartScrollPage + lWord.FScrollCount - 1 %>
			<% If (i > lWord.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(lWord.FCurrPage) Then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% End if %>
		<% Next %>
		<% If lWord.HasNextScroll Then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
<% Else %>
<tr height="50" bgcolor="FFFFFF" >
	<td colspan="15" align="center">
		등록된 데이터가 없습니다
	</td>
</tr>
<% End If %>
</table>
<% SET lWord = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->