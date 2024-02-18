<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/event/contestCls.asp"-->

<%
	Dim iCurrentpage, contestlist, i, iTotCnt, vEvtCode, vSDate, page, vDiv, vUserID
	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	page 			= NullFillWith(requestCheckVar(request("page"),5),1)
	vDiv			= requestCheckVar(Request("divnum"),10)
	vUserID			= requestCheckVar(Request("userid"),32)
	
	Set contestlist = new ClsContest
	contestlist.FCurrPage = page
	contestlist.FDiv = vDiv
	contestlist.FUserID = vUserID
	contestlist.FEntryList
	
	iTotCnt = contestlist.ftotalcount
%>

<script tepe="text/javascript">

function viewtext(i){
	if(document.getElementById("text"+i+"").style.display == "none"){
		document.getElementById("text"+i+"").style.display = "table-row";
	}else{
		document.getElementById("text"+i+"").style.display = "none";
	}
}

function finallist(){
	var pop_view = window.open('popup_finallist.asp?divnum=<%=vDiv%>','pop_view','width=800,height=700,scrollbars=yes,resizable=yes');
	pop_view.focus();
}
</script>


<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#999999">
<form name="frm" action="detail_list.asp" method="get">
<input type="hidden" name="menupos" value="<%=Request("menupos")%>">
<input type="hidden" name="divnum" value="<%=vDiv%>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="70" bgcolor="#EEEEEE">검색 조건</td>
	<td align="left">
		아이디 : <input type="text" name="userid" value="<%=vUserID%>" size="15">
		&nbsp;&nbsp;&nbsp;
		<input type="submit" value="검색" class="button" onfocus="this.blur();">
	</td>
</tr>
</form>
</table>
<br>
<table cellpadding="0" cellspacing="0" class="a">
<tr height="25">
	<td colspan="20">
		<table width="700" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td>Total Count : <b><%= iTotCnt %></b></td>
			<td align="right">
				<input type="button" value="공모전리스트" onClick="location.href='index.asp?menupos=<%=Request("menupos")%>';">&nbsp;&nbsp;
				<input type="button" value="파이널리스트 선정" onClick="finallist()">
			</td>
		</tr>
		</table>
	</td>
</tr>
<%
	If contestlist.FResultCount <> 0 Then
		For i = 0 To contestlist.FResultCount -1
%>
		<tr bgcolor="FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
			<td>
				<table width="700" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<% If i = 0 Then %>
				<tr align="center" bgcolor="#E6E6E6" height="20">
					<td>idx</td>
					<td>공모전</td>
					<td>응모자</td>
					<td>응모일</td>
				</tr>
				<% End If %>
				<tr bgcolor="FFFFFF" style="cursor:pointer" onClick="viewtext('<%=i%>');">
					<td width="50" align="center"><%=contestlist.FItemList(i).fidx%></td>
					<td width="250" align="center"><%=contestlist.FItemList(i).fsubject%></td>
					<td width="200"><%=contestlist.FItemList(i).fusername%>(<%=contestlist.FItemList(i).fuserid%>)</td>
					<td width="200" align="center" style="word-break:break-all;"><%=contestlist.FItemList(i).fregdate%></td>
				</tr>
				<tr bgcolor="FFFFFF" id="text<%=i%>" style="display:none;">
					<td colspan="4" style="padding:5 3 5 3;">
						<%
							If contestlist.FItemList(i).fimgFile1 <> "" Then
								Response.Write "파일 1 : [<a href='//imgstatic.10x10.co.kr/linkweb/enjoy/ContestFileDownload.asp?filename=" & contestlist.FItemList(i).fimgFile1 & "'>받기</a>]"
								If Right(contestlist.FItemList(i).fimgFile1,3) = "jpg" OR Right(contestlist.FItemList(i).fimgFile1,3) = "gif" Then
									Response.Write "&nbsp;[<a href='http://www.10x10.co.kr/common/showimage.asp?img=" & contestlist.FItemList(i).fimgFile1 & "' target='_blank'>바로보기</a>]"
								End IF
							End If
							If contestlist.FItemList(i).fimgFile2 <> "" Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
								Response.Write "파일 2 : [<a href='//imgstatic.10x10.co.kr/linkweb/enjoy/ContestFileDownload.asp?filename=" & contestlist.FItemList(i).fimgFile2 & "'>받기</a>]"
								If Right(contestlist.FItemList(i).fimgFile2,3) = "jpg" OR Right(contestlist.FItemList(i).fimgFile2,3) = "gif" Then
									Response.Write "&nbsp;[<a href='http://www.10x10.co.kr/common/showimage.asp?img=" & contestlist.FItemList(i).fimgFile2 & "' target='_blank'>바로보기</a>]"
								End IF
							End If
							If contestlist.FItemList(i).fimgFile3 <> "" Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
								Response.Write "파일 3 : [<a href='//imgstatic.10x10.co.kr/linkweb/enjoy/ContestFileDownload.asp?filename=" & contestlist.FItemList(i).fimgFile3 & "'>받기</a>]"
								If Right(contestlist.FItemList(i).fimgFile3,3) = "jpg" OR Right(contestlist.FItemList(i).fimgFile3,3) = "gif" Then
									Response.Write "&nbsp;[<a href='http://www.10x10.co.kr/common/showimage.asp?img=" & contestlist.FItemList(i).fimgFile3 & "' target='_blank'>바로보기</a>]"
								End IF
							End If
							If contestlist.FItemList(i).fimgFile4 <> "" Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
								Response.Write "파일 4 : [<a href='//imgstatic.10x10.co.kr/linkweb/enjoy/ContestFileDownload.asp?filename=" & contestlist.FItemList(i).fimgFile4 & "'>받기</a>]"
								If Right(contestlist.FItemList(i).fimgFile4,3) = "jpg" OR Right(contestlist.FItemList(i).fimgFile4,3) = "gif" Then
									Response.Write "&nbsp;[<a href='http://www.10x10.co.kr/common/showimage.asp?img=" & contestlist.FItemList(i).fimgFile4 & "' target='_blank'>바로보기</a>]"
								End IF
							End If
							If contestlist.FItemList(i).fimgFile5 <> "" Then
								Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
								Response.Write "파일 5 : [<a href='//imgstatic.10x10.co.kr/linkweb/enjoy/ContestFileDownload.asp?filename=" & contestlist.FItemList(i).fimgFile5 & "'>받기</a>]"
								If Right(contestlist.FItemList(i).fimgFile5,3) = "jpg" OR Right(contestlist.FItemList(i).fimgFile5,3) = "gif" Then
									Response.Write "&nbsp;[<a href='http://www.10x10.co.kr/common/showimage.asp?img=" & contestlist.FItemList(i).fimgFile5 & "' target='_blank'>바로보기</a>]"
								End IF
							End If
						%>
						<br><br>
						텐바이텐 디자인 공모전을 어떻게 알게 되셨나요?<br>
						->&nbsp;<%=contestlist.FItemList(i).GetOptTypeName%>(<%=contestlist.FItemList(i).foptText%>)<br><br>
						디자인 컨셉 설명<br>
						->&nbsp;<%=contestlist.FItemList(i).fimgContent%>
					</td>
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
				<a href="?page=1&divnum=<%=vDiv%>&userid=<%=vUserID%>&menupos=<%=Request("menupos")%>"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pageprev02.gif" width="9" height="9" border="0" /></a>
				<% if contestlist.HasPreScroll then %>
					&nbsp;&nbsp;<a href="?page=<%= contestlist.StartScrollPage-1 %>&divnum=<%=vDiv%>&userid=<%=vUserID%>&menupos=<%=Request("menupos")%>"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pageprev01.gif" width="9" height="9" border="0" /></a>
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
					&nbsp;&nbsp;&nbsp;&nbsp;<a href="?page=<%= i %>&divnum=<%=vDiv%>&userid=<%=vUserID%>&menupos=<%=Request("menupos")%>" style="cursor:pointer"><%= i %></a>
				<% 
				end if 
				next 
				%>													
				<% if contestlist.HasNextScroll then %>
					&nbsp;&nbsp;<span class="list_link"><a href="?page=<%= i %>&divnum=<%=vDiv%>&userid=<%=vUserID%>&menupos=<%=Request("menupos")%>"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pagenext01.gif" width="9" height="9" border="0" /></a>
				<% else %>
					&nbsp;&nbsp;<img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pagenext01.gif" width="9" height="9" border="0" />
				<% end if %>																												
				&nbsp;&nbsp;&nbsp;<a href="?page=<%= contestlist.FTotalpage %>&divnum=<%=vDiv%>&userid=<%=vUserID%>&menupos=<%=Request("menupos")%>" onfocus="this.blur();"><img src="http://fiximage.10x10.co.kr/web2009/momo/images/btn_pagenext02.gif" width="9" height="9" border="0" /></a>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>

<%
	set contestlist = nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->