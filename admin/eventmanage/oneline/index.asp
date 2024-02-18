<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/onelineCls.asp"-->

<%
	Dim iCurrentpage, onelinelist, i, iTotCnt, vEvtCode, vSDate, page
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	page 			= NullFillWith(requestCheckVar(request("page"),5),1)
	vEvtCode		= requestCheckVar(Request("eC"),10)
	vSDate			= requestCheckVar(Request("esday"),10)
	
	Set onelinelist = new ClsOneLine
	onelinelist.FEvtCode = vEvtCode
	onelinelist.FCurrPage = page
	onelinelist.FOneLineList
	
	iTotCnt = onelinelist.ftotalcount
%>

<script language="javascript">
function delproc(gubun,idx)
{
  	document.delprocfrm.gubun.value = gubun;
	document.delprocfrm.idx.value = idx;
	document.delprocfrm.target = "delProc";
	document.delprocfrm.action = "delproc.asp";
	document.delprocfrm.submit();
}
</script>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		Total Count : <b><%= iTotCnt %></b>
	</td>
</tr>
<%
	If onelinelist.FResultCount <> 0 Then
		For i = 0 To onelinelist.FResultCount -1
%>
		<tr bgcolor="FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer">
			<td width="100%">
				<table cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" width="100%">
				<tr bgcolor="FFFFFF">
					<td width="35" align="center"><%=onelinelist.FItemList(i).fidx%></td>
					<td width="20" align="center"><img src="http://fiximage.10x10.co.kr/web2010/oneline/emoticon_0<%=onelinelist.FItemList(i).ficon%>_s.gif" width="20" height="20"></td>
					<td>
						<%
							If onelinelist.FItemList(i).fwinYN = "Y" Then
								Response.Write "[<font color=red><b>" & DatePart("m",vSDate) & "월 " & getWeekSerial(vSDate) & "주차 당첨자</b></font>]&nbsp;"
							End If
						%>
						<%=onelinelist.FItemList(i).fuserid%> (<font color="<%= getUserLevelColorByDate(onelinelist.FItemList(i).fuserlevel, left(onelinelist.FItemList(i).fregdate,10)) %>">
						<b><%= getUserLevelStrByDate(onelinelist.FItemList(i).fuserlevel, left(onelinelist.FItemList(i).fregdate,10)) %></b></font>)
					</td>
					<td width="140" align="center"><%=onelinelist.FItemList(i).fregdate%></td>
					<td width="80" align="center">
						<%
							If onelinelist.FItemList(i).fisusing = "Y" Then
								Response.Write "<input type='button' value='삭제하기' onClick='delproc(0," & onelinelist.FItemList(i).fidx & ");'>"
							Else
								Response.Write "삭제된글<br><input type='button' value='되살리기' onClick='delproc(1," & onelinelist.FItemList(i).fidx & ");'>"
							End IF
						%>
					</td>
				</tr>
				<tr bgcolor="FFFFFF">
					<td colspan="5" width="100%" style="padding:5 3 5 3;"><%=onelinelist.FItemList(i).fcomment%></td>
				</tr>
				</table>
			</td>
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
<tr>
	<td align="center" valign="top" style="padding-bottom:30px">
		<table border=0 cellspacing="5" class="a">
		<tr>																		
			<td align="center">
				<a href="?page=1&eC=<%=vEvtCode%>&esday=<%=vSDate%>"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pageprev02.gif" width="9" height="9" border="0" /></a>
				<% if onelinelist.HasPreScroll then %>
					&nbsp;&nbsp;<a href="?page=<%= onelinelist.StartScrollPage-1 %>&eC=<%=vEvtCode%>&esday=<%=vSDate%>"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pageprev01.gif" width="9" height="9" border="0" /></a>
				<% else %>
					&nbsp;&nbsp;<img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pageprev01.gif" width="9" height="9" border="0" />
				<% end if %>																												
				<% 
				for i = 0 + onelinelist.StartScrollPage to onelinelist.StartScrollPage + onelinelist.FScrollCount - 1 
				if (i > onelinelist.FTotalpage) then Exit for 
				if CStr(i) = CStr(onelinelist.FCurrPage) then 
				%>
					&nbsp;&nbsp;&nbsp;&nbsp;<span class="eng11pxblack"><b><%= i %></b></span>
				<% else %>
					&nbsp;&nbsp;&nbsp;&nbsp;<a href="?page=<%= i %>&eC=<%=vEvtCode%>&esday=<%=vSDate%>" style="cursor:pointer"><%= i %></a>
				<% 
				end if 
				next 
				%>													
				<% if onelinelist.HasNextScroll then %>
					&nbsp;&nbsp;<span class="list_link"><a href="?page=<%= i %>&eC=<%=vEvtCode%>&esday=<%=vSDate%>"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pagenext01.gif" width="9" height="9" border="0" /></a>
				<% else %>
					&nbsp;&nbsp;<img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pagenext01.gif" width="9" height="9" border="0" />
				<% end if %>																												
				&nbsp;&nbsp;&nbsp;<a href="?page=<%= onelinelist.FTotalpage %>&eC=<%=vEvtCode%>&esday=<%=vSDate%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pagenext02.gif" width="9" height="9" border="0" /></a>
			</td>
		</tr>
		</table>
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
	set onelinelist = nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->