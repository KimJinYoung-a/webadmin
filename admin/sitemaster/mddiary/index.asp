<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/mddiaryCls.asp"-->

<%
	Dim iCurrentpage, mddiarylist, i, iTotCnt, vEvtCode, vSDate, page
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	page 			= NullFillWith(requestCheckVar(request("page"),5),1)

	
	Set mddiarylist = new Clsmddiary
	mddiarylist.FCurrPage = page
	mddiarylist.FmddiaryList
	
	iTotCnt = mddiarylist.ftotalcount
%>

<script language="javascript">
document.domain = "10x10.co.kr";

function delproc(gubun,idx)
{
  	document.delprocfrm.gubun.value = gubun;
	document.delprocfrm.idx.value = idx;
	document.delprocfrm.target = "delProc";
	document.delprocfrm.action = "delproc.asp";
	document.delprocfrm.submit();
}

function mddiaryWrite(id)
{
	var mddiary = window.open('mddiary_write.asp?mgzId='+id+'','mddiary','width=540,height=527');
	mddiary.focus();
}
</script>

<!-- 리스트 시작 -->
<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="5">
		<table width="100%" cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td>
				Total Count : <b><%= iTotCnt %></b>
			</td>
			<td align="right">
				<input type="button" value="등 록" onClick="mddiaryWrite('')">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr bgcolor="silver">
	<td align="center">Diary No.</td>
	<td align="center">제 목</td>
	<td align="center">오픈일</td>
	<td align="center">사용여부</td>
	<td align="center">미리보기</td>
</tr>
<%
	If mddiarylist.FResultCount <> 0 Then
		For i = 0 To mddiarylist.FResultCount -1
%>
		<tr bgcolor="FFFFFF">
			<td width="70" align="center" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer" onClick="mddiaryWrite(<%=mddiarylist.FItemList(i).fmgzId%>)"><%=mddiarylist.FItemList(i).fmgzId%></td>
			<td width="210" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer" onClick="mddiaryWrite(<%=mddiarylist.FItemList(i).fmgzId%>)"><img src="<%=mddiarylist.FItemList(i).fmenuimg%>"></td>
			<td width="100" align="center" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer" onClick="mddiaryWrite(<%=mddiarylist.FItemList(i).fmgzId%>)"><%=Left(mddiarylist.FItemList(i).fopendate,10) %></td>
			<td width="70" align="center" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer" onClick="mddiaryWrite(<%=mddiarylist.FItemList(i).fmgzId%>)"><%=mddiarylist.FItemList(i).fuseyn %></td>
			<td width="70" align="center"><a href="http://www.10x10.co.kr/event/magazine/?page=1&preview=previewok&mgzId=<%=mddiarylist.FItemList(i).fmgzId %>" target="_blank">[Preview]</a></td>
		</tr>
<%
		Next
	Else
%>
		<tr bgcolor="#FFFFFF" height="30">
			<td colspan="20" align="center" class="page_link">[데이터가 없습니다.]</td>
		</tr>
<%
	End If
%>
<tr bgcolor="#FFFFFF">
	<td align="center" style="padding:10 0 10 0" colspan="5">
		<a href="?page=1"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pageprev02.gif" width="9" height="9" border="0" /></a>
		<% if mddiarylist.HasPreScroll then %>
			&nbsp;&nbsp;<a href="?page=<%= mddiarylist.StartScrollPage-1 %>"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pageprev01.gif" width="9" height="9" border="0" /></a>
		<% else %>
			&nbsp;&nbsp;<img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pageprev01.gif" width="9" height="9" border="0" />
		<% end if %>																												
		<% 
		for i = 0 + mddiarylist.StartScrollPage to mddiarylist.StartScrollPage + mddiarylist.FScrollCount - 1 
		if (i > mddiarylist.FTotalpage) then Exit for 
		if CStr(i) = CStr(mddiarylist.FCurrPage) then 
		%>
			&nbsp;&nbsp;&nbsp;&nbsp;<span class="eng11pxblack"><b><%= i %></b></span>
		<% else %>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="?page=<%= i %>" style="cursor:pointer"><%= i %></a>
		<% 
		end if 
		next 
		%>													
		<% if mddiarylist.HasNextScroll then %>
			&nbsp;&nbsp;<span class="list_link"><a href="?page=<%= i %>"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pagenext01.gif" width="9" height="9" border="0" /></a>
		<% else %>
			&nbsp;&nbsp;<img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pagenext01.gif" width="9" height="9" border="0" />
		<% end if %>																												
		&nbsp;&nbsp;&nbsp;<a href="?page=<%= mddiarylist.FTotalpage %>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pagenext02.gif" width="9" height="9" border="0" /></a>
	</td>
</tr>
</table>

<form name="delprocfrm" method="post">
<input type="hidden" name="idx" value="">
<input type="hidden" name="gubun" value="">
<input type="hidden" name="eC" value="<%=vEvtCode%>">
</form>
<iframe id="delProc" name="delProc" src="about:blank" frameborder="0" width="0" height="0"></iframe>

<%
	set mddiarylist = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->