<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 촬영 요청 등록페이지
' History : 2012.03.13 김진영 생성
'			2015.07.28 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/photo_req/requestCls.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<%
IF application("Svr_Info")="Dev" THEN
	g_MenuPos   = "1404"		'### 메뉴번호 지정.
Else
	g_MenuPos   = "1419"		'### 메뉴번호 지정.
End If

Dim lPhotoreq, page, i, makerid, cdl, r_use, s_type, num_name, req_status_type, request_name, req_photo_user, req_stylist
Dim iPageSize, iCurrentpage ,iDelCnt, sSearchTeam, sDoc_Status, sDoc_AnsOX, sSearchMine, confirmdate, tmpconfirmdate, j
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
Dim iTotCnt
	page = request("page")

If page = "" Then page = 1

'검색 Get값 들..
makerid 		= request("makerid")
cdl				= request("req_category")
r_use			= request("req_use")
s_type			= request("s_type")
num_name 		= request("num_name")
req_status_type = request("req_status_type")
request_name 	= request("request_name")
req_photo_user 	= request("req_photo_user")
req_stylist		= request("req_stylist")

set lPhotoreq = new Photoreq
	lPhotoreq.FPageSize = 20
	lPhotoreq.FCurrPage = page
	lPhotoreq.FMakerid = makerid
	lPhotoreq.FCdl = cdl
	lPhotoreq.FReq_use = r_use
	lPhotoreq.FS_type = s_type
	lPhotoreq.FNum_Name = num_name
	lPhotoreq.FReq_status_type = req_status_type
	lPhotoreq.FRequest_name = request_name
	lPhotoreq.FReq_photo_user = Trim(req_photo_user)
	lPhotoreq.FReq_stylist = Trim(req_stylist)
	lPhotoreq.fnPhotoreqlist
%>
<script language="javascript">
function code_manage()
{
	window.open('PopManageCode.asp','coopcode','width=410,height=600');
}
function user_manage()
{
	window.open('PopUserList.asp','coopcode','width=410,height=600');
}
function gosubmit(page){
    document.searchfrm.page.value=page;
	document.searchfrm.submit();
}
function goUpdate(didx)
{
	location.href = "/admin/photo_req/request_modi.asp?req_no="+didx+"&udate=A&menupos=<%= menupos %>";
}
</script>
<p>
<!-- height="100%" 이거 때문인지 공지만 나오고 아래 촬영요청리스트가 안나옴. 전종윤. -->
<iframe src="/admin/photo_req/board_list.asp" name="board" width="100%" height="200" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe>
<p>
<!-- //-->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr><td><b>[촬영요청리스트]</b></td></tr>
</table>
<p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="searchfrm" action="request_list.asp" method="get">
	<tr align="center" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("gray") %>" width="100">검색조건</td>
		<td align="left">
			<table width="100%" align="center"  cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr bgcolor="FFFFFF">
				<td width="150">브랜드 : </td>
				<td><%	drawSelectBoxDesignerWithName "makerid", makerid %></td>
				<td width="150">전시카테고리 : </td>
				<td colspan="5">
					<%' call DrawCategoryLarge_disp("req_category", cdl) %>
					<%= fnStandardDispCateSelectBox(1,cdl, "req_category", cdl, "")%>
				</td>
			</tr>
			<tr bgcolor="FFFFFF">
				<td width="150">촬영용도 : </td>
				<td><% call DrawPicGubun("req_use", r_use, "2") %></td>
				<td width="150">no/상품명 : </td>
				<td colspan="5">
					<select name="s_type" class="select">
						<option value="">--no/상품명선택--</option>
						<option value="1" <%If s_type = "1" Then response.write "selected" End If%>>요청서 no</option>
						<option value="2" <%If s_type = "2" Then response.write "selected" End If%>>상품명</option>
					</select>
					<input type="text" class="text" name="num_name" value="<%=num_name%>" size="30" maxlength="100" onKeyPress="if (event.keyCode == 13) document.searchfrm.submit();">
				</td>
			</tr>
			<tr bgcolor="FFFFFF">
				<td width="150">진행상태 : </td>
				<td>
					<select name="req_status_type" class="select">
						<option value="">--진행상태선택--</option>
						<option value="4" <%If req_status_type = "4" Then response.write "selected" End If%>>추가기입 요청</option>
						<option value="1" <%If req_status_type = "1" Then response.write "selected" End If%>>촬영스케줄 지정</option>
						<option value="2" <%If req_status_type = "2" Then response.write "selected" End If%>>촬영중</option>
						<option value="3" <%If req_status_type = "3" Then response.write "selected" End If%>>촬영완료</option>
						<option value="9" <%If req_status_type = "9" Then response.write "selected" End If%>>최종오픈</option>
					</select>
				</td>
				<td width="150">촬영요청자 : </td>
				<td><input type="text" class="text" name="request_name" size="16" maxlength="16" value="<%=request_name%>" onKeyPress="if (event.keyCode == 13) document.searchfrm.submit();"></td>
				<td>담당포토 : </td>
				<td><input type="text" class="text" name="req_photo_user" size="16" maxlength="16" value="<%=req_photo_user%>" onKeyPress="if (event.keyCode == 13) document.searchfrm.submit();"></td>
				<td>담당스타일리스트 : </td>
				<td><input type="text" class="text" name="req_stylist" size="16" maxlength="16" value="<%=req_stylist%>" onKeyPress="if (event.keyCode == 13) document.searchfrm.submit();"></td>
			</tr>
			</table>
		</td>
		<td bgcolor="<%= adminColor("gray") %>" width="100"><input type="button" class="button_s" value="검색" onClick="javascript:document.searchfrm.submit();"></td>
	</tr>
</table>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="새로등록" onClick="location.href='request_write.asp?menupos=<%=menupos%>&iC=<%=iCurrentpage%>'">
		<input type='button' class='button' value='관리' onClick='user_manage()'>
		&nbsp;<font color="red"><ins>* 등록 전 위의 공지사항,필독 부탁 드립니다.</ins></font>
	</td>
	<td align="right">
		<%
			Response.Write "<input type='button' class='button' value='코드관리' onClick='code_manage()'>&nbsp;"
		%>
	</td>
</tr>
</table>
<br>
<!-- 액션 끝 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" >
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">검색결과 : <b><%= lPhotoreq.FTotalCount %></b>&nbsp;&nbsp;&nbsp;&nbsp;페이지 : <b><%=page%>/<%=lPhotoreq.FTotalPage%></b></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
	<td width="60">요청서No</td>
	<td width="100">진행상태</td>
	<td width="200">촬영용도</td>
	<td width="">상품명(기획전명)</td>
	<td width="100">카테고리</td>
	<!--<td width="">브랜드</td>-->
	<td width="130">요청일시</td>
	<td width="260">촬영확정일시</td>
	<td width="60">담당MD<BR>(촬영요청)</td>
	<td width="60">중요도</td>
	<td width="50">완성URL<Br>등록여부</td>
</tr>
<%
	If lPhotoreq.FResultcount = 0 Then
%>
	<tr bgcolor="#FFFFFF" height="30">
		<td colspan="10" align="center" class="page_link">[데이터가 없습니다.]</td>
	</tr>
<%
	Else
		For i = 0 to lPhotoreq.FResultcount -1
%>
	<tr align="center" bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'" style="cursor:pointer" onClick="goUpdate('<%=lPhotoreq.FPhotoreqList(i).FReq_no%>')">
		<td><%=lPhotoreq.FPhotoreqList(i).FReq_no%></td>
		<td>
			<%
			Select Case lPhotoreq.FPhotoreqList(i).FReq_status
				Case "0"	lPhotoreq.FPhotoreqList(i).FReq_status  = ""
				Case "1"	lPhotoreq.FPhotoreqList(i).FReq_status  = "촬영스케줄 지정"
				Case "2"	lPhotoreq.FPhotoreqList(i).FReq_status  = "촬영중"
				Case "3"	lPhotoreq.FPhotoreqList(i).FReq_status  = "촬영완료"
				Case "4"	lPhotoreq.FPhotoreqList(i).FReq_status  = "추가 기입 요청건"
				Case "9"	lPhotoreq.FPhotoreqList(i).FReq_status  = "최종오픈"
			End Select

			Select Case lPhotoreq.FPhotoreqList(i).FFontColor
				Case "R"	response.write "<font color='RED'>"&lPhotoreq.FPhotoreqList(i).FReq_status&"</font>"
				Case "G"	response.write "<font color='GREEN'>"&lPhotoreq.FPhotoreqList(i).FReq_status&"</font>"
				Case Else	response.write "<font color='BLACK'>"&lPhotoreq.FPhotoreqList(i).FReq_status&"</font>"
			End Select
			%>
		</td>
		<td>
			<%=lPhotoreq.FPhotoreqList(i).FReq_use%>
			<%
				If lPhotoreq.FPhotoreqList(i).FReq_use_detail <> "" Then
					response.write "("&lPhotoreq.FPhotoreqList(i).FReq_use_detail&")"
				End If
			%>
		</td>
		<td align="left"><%=DDotFormat(lPhotoreq.FPhotoreqList(i).FReq_prd_name,20)%></td>
		<td><%=lPhotoreq.FPhotoreqList(i).FReq_codenm%></td>
		<!--<td><%'=lPhotoreq.FPhotoreqList(i).FReq_makerid%></td>-->
		<td>
			요청 일시 : <%= Left(lPhotoreq.FPhotoreqList(i).FReq_regdate,10) %><br>
		</td>
		<td>
			<%
			confirmdate = lPhotoreq.FPhotoreqList(i).fconfirmdate
			%>
			<% if confirmdate <> "" then %>
				<% For j = LBound(Split(confirmdate,"|^|")) To UBound(Split(confirmdate,"|^|")) %>
				<%
				tmpconfirmdate = Split(confirmdate,"|^|")(j)
				tmpconfirmdate = Split(tmpconfirmdate,"|*|")
				%>
				<%= left(tmpconfirmdate(0),10) %>
				<% if tmpconfirmdate(2)<>"" or tmpconfirmdate(3)<>"" then %>
					(
					<% if tmpconfirmdate(2) <> "" then %>
						포토 : <%= tmpconfirmdate(2) %>
					<% end if %>
					<% if tmpconfirmdate(3) <> "" then %>
						, 스타일 : <%= tmpconfirmdate(3) %>
					<% end if %>
					)
				<% end if %>
				<br>
				<% next %>
			<% end if %>
		</td>
		<td>
			<%
			If isnull(lPhotoreq.FPhotoreqList(i).FMDid) = "False" Then
				response.write lPhotoreq.FPhotoreqList(i).FMDid&"<br>("& lPhotoreq.FPhotoreqList(i).FReq_name &")"
			ElseIf isnull(lPhotoreq.FPhotoreqList(i).FMDid) = "True" or (lPhotoreq.FPhotoreqList(i).FMDid) = "00" Then
				response.write lPhotoreq.FPhotoreqList(i).FReq_name
			End If
			%>
		</td>
		<td>
			<% for j = 1 to lPhotoreq.FPhotoreqList(i).FImport_level %>★<% next %>
		</td>
		<td>
			<% if lPhotoreq.FPhotoreqList(i).fopencount>0 then %>
				Y
			<% else %>
				N
			<% end if %>
		</td>
	</tr>
<%
		Next
	End If
%>
<tr height="25" bgcolor="FFFFFF" >
	<td colspan="15" align="center">
       	<% If lPhotoreq.HasPreScroll Then %>
			<a href="javascript:gosubmit('<%= lPhotoreq.StartScrollPage-1 %>');">[pre]</a>
		<% Else %>
		[pre]
		<% End If %>
		<% For i = 0 + lPhotoreq.StartScrollPage to lPhotoreq.StartScrollPage + lPhotoreq.FScrollCount - 1 %>
			<% If (i > lPhotoreq.FTotalpage) Then Exit for %>
			<% If CStr(i) = CStr(lPhotoreq.FCurrPage) Then %>
			<font color="red">[<%= i %>]</font>
			<% Else %>
			<a href="javascript:gosubmit('<%= i %>');">[<%= i %>]</a>
			<% End if %>
		<% Next %>
		<% If lPhotoreq.HasNextScroll Then %>
			<a href="javascript:gosubmit('<%= i %>');">[next]</a>
		<% Else %>
		[next]
		<% End If %>
	</td>
</tr>
</form>
</table>
<%set lPhotoreq = nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->