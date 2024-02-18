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
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/contestCls.asp"-->
<%
	Dim iCurrentpage, contestlist, i, iTotCnt, vSubject, page, vUseYN
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	page 			= NullFillWith(requestCheckVar(request("page"),5),1)
	vSubject		= requestCheckVar(Request("subject"),100)
	vUseYN			= NullFillWith(request("useyn"),"y")
	
	Set contestlist = new ClsContest
	contestlist.FCurrPage = page
	contestlist.FUseYN = vUseYN
	contestlist.FSubject = vSubject
	contestlist.FContestList
	
	iTotCnt = contestlist.ftotalcount
%>

<script type='text/javascript'>

function delproc(gubun,idx)
{
  	document.delprocfrm.gubun.value = gubun;
	document.delprocfrm.idx.value = idx;
	document.delprocfrm.target = "delProc";
	document.delprocfrm.action = "delproc.asp";
	document.delprocfrm.submit();
}

function viewtext(i)
{
	location.href = "detail_list.asp?menupos=<%=Request("menupos")%>&divnum="+i+"";
}

function newRegContest(con){
	var pop_view = window.open('popup_contestdetail.asp?contest='+con+'','popup_contestdetail','width=800,height=700,scrollbars=no,resizable=no');
	pop_view.focus();
}

</script>

<!-- 검색 시작 -->
<form name="frm" action="index.asp" method="get" style="margin:0px;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<input type="hidden" name="menupos" value="<%=Request("menupos")%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			사용여부 : 
			<select name="useyn">
			<option value="y" <% If vUseYN = "y" Then Response.Write "selected" End If %>>y</option>
			<option value="n" <% If vUseYN = "n" Then Response.Write "selected" End If %>>n</option>
			</select>&nbsp;&nbsp;&nbsp;
			제목 : <input type="text" name="subject" value="<%=vSubject%>" size="50">
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left"></td>
	</tr>
</table>
</form>

<br>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="right" style="padding:5px 0 5px 0;"><input type="button" class="button" value="신규등록" onclick="newRegContest('');"></td>
</tr>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= iTotCnt %></b>
	</td>
</tr>
<tr align="center" bgcolor="#E6E6E6" height="25">
	<td>공모전 No.</td>
	<td>주 제</td>
	<td>응모기간</td>
	<td>고객투표기간</td>
	<td>당선자발표일</td>
	<td>사용여부</td>
	<td>비고</td>
</tr>
<%
	If contestlist.FResultCount <> 0 Then
		For i = 0 To contestlist.FResultCount -1
%>
			<tr bgcolor="FFFFFF" height="25" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
				<td width="100" align="center" style="cursor:pointer" onClick="viewtext('<%=contestlist.FItemList(i).fcontest%>');"><%=contestlist.FItemList(i).fcontest%></td>
				<td style="cursor:pointer" onClick="viewtext('<%=contestlist.FItemList(i).fcontest%>');">
					<%= ReplaceBracket(contestlist.FItemList(i).fsubject) %>
				</td>
				<td width="150" align="center" style="cursor:pointer" onClick="viewtext('<%=contestlist.FItemList(i).fcontest%>');"><%=contestlist.FItemList(i).fentry_sdate%> ~ <%=contestlist.FItemList(i).fentry_edate%></td>
				<td width="150" align="center" style="cursor:pointer" onClick="viewtext('<%=contestlist.FItemList(i).fcontest%>');"><%=contestlist.FItemList(i).fvote_sdate%> ~ <%=contestlist.FItemList(i).fvote_edate%></td>
				<td width="100" align="center" style="cursor:pointer" onClick="viewtext('<%=contestlist.FItemList(i).fcontest%>');"><%=contestlist.FItemList(i).fresult_date%></td>
				<td width="50" align="center" style="cursor:pointer" onClick="viewtext('<%=contestlist.FItemList(i).fcontest%>');"><%=contestlist.FItemList(i).fuseyn%></td>
				<td width="50" align="center">
					<input type="button" class="button" value="수정" onclick="newRegContest('<%=contestlist.FItemList(i).fcontest%>');">
				</td>
			</tr>
<%
		Next
	Else
%>
		<tr bgcolor="#FFFFFF" height="30" colspan="20">
			<td width="850" align="center" class="page_link">[데이터가 없습니다.]</td>
		</tr>
<%
	End If
%>
<tr bgcolor="#FFFFFF" height="30">
	<td align="center" colspan="20">
		<table border=0 cellspacing="0" class="a">
		<tr>																		
			<td align="center">
				<a href="?page=1&menupos=<%=Request("menupos")%>"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pageprev02.gif" width="9" height="9" border="0" /></a>
				<% if contestlist.HasPreScroll then %>
					&nbsp;&nbsp;<a href="?page=<%= contestlist.StartScrollPage-1 %>&menupos=<%=Request("menupos")%>"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pageprev01.gif" width="9" height="9" border="0" /></a>
				<% else %>
					&nbsp;&nbsp;<img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pageprev01.gif" width="9" height="9" border="0" />
				<% end if %>																												
				<% 
				for i = 0 + contestlist.StartScrollPage to contestlist.StartScrollPage + contestlist.FScrollCount - 1 
				if (i > contestlist.FTotalpage) then Exit for 
				if CStr(i) = CStr(contestlist.FCurrPage) then 
				%>
					&nbsp;&nbsp;&nbsp;&nbsp;<span class="eng11pxblack"><b><%= i %></b></span>
				<% else %>
					&nbsp;&nbsp;&nbsp;&nbsp;<a href="?page=<%= i %>&menupos=<%=Request("menupos")%>" style="cursor:pointer"><%= i %></a>
				<% 
				end if 
				next 
				%>													
				<% if contestlist.HasNextScroll then %>
					&nbsp;&nbsp;<span class="list_link"><a href="?page=<%= i %>&menupos=<%=Request("menupos")%>"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pagenext01.gif" width="9" height="9" border="0" /></a>
				<% else %>
					&nbsp;&nbsp;<img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pagenext01.gif" width="9" height="9" border="0" />
				<% end if %>																												
				&nbsp;&nbsp;&nbsp;<a href="?page=<%= contestlist.FTotalpage %>&menupos=<%=Request("menupos")%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pagenext02.gif" width="9" height="9" border="0" /></a>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>

<form name="delprocfrm" method="post">
<input type="hidden" name="idx" value="">
<input type="hidden" name="gubun" value="">
</form>
<iframe id="delProc" name="delProc" src="about:blank" frameborder="0" width="0" height="0"></iframe>

<%
	set contestlist = nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->