<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cooperate/chk_auth.asp"-->
<!-- #include virtual="/lib/classes/cooperate/cooperateCls.asp"-->
<!-- #include virtual="/lib/classes/approval/eappCls.asp"-->

<% If NOT (C_ADMIN_AUTH) Then %>
<script language="javascript">
<!--  
	 window.open("/admin/cooperate/popIndex.asp","popCooperate","width="+screen.availWidth+", height="+ screen.availHeight +",resizable=yes, scrollbars=yes"); 
//-->
</script>

<%
	dbget.close()
	Response.End
End If

	Dim iTotCnt, arrList,intLoop, vParam
	Dim iPageSize, iCurrentpage ,iDelCnt, sSearchTeam, sDoc_Status, sDoc_AnsOX, sSearchMine, sUserName, sSearching, sContent
	Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt
	Dim sDoc_Type

	iCurrentpage 	= NullFillWith(requestCheckVar(Request("iC"),10),1)
	sSearchTeam		= NullFillWith(requestCheckVar(Request("search_team"),20),"")
	sDoc_Status		= NullFillWith(requestCheckVar(Request("doc_status"),10),"0")
	sDoc_Type		= NullFillWith(requestCheckVar(Request("doc_type"),10),"")
	sDoc_AnsOX		= NullFillWith(requestCheckVar(Request("ans_ox"),1),"")
	sSearchMine		= NullFillWith(requestCheckVar(Request("onlymine"),1),"o")
	sUserName		= NullFillWith(requestCheckVar(Request("username"),10),"")
	sSearching		= NullFillWith(requestCheckVar(Request("searching"),10),"")
	sContent		= NullFillWith(requestCheckVar(Request("content"),100),"")
	iPageSize 		= 20
	iPerCnt 		= 10

	If sSearching = "doc_idx" AND IsNumeric(sContent) = False Then
		Response.Write "<script language='javascript'>alert('협조No 겁색은 숫자로만 입력하셔야 합니다.');history.back();</script>"
	End IF

	vParam = "&iC="&iCurrentpage&"&s_search_team="&sSearchTeam&"&s_status="&sDoc_Status&"&s_type="&sDoc_Type&"&s_ans_ox="&sDoc_AnsOX&"&s_onlymine="&sSearchMine&"&username="&sUserName&"searching="&sSearching&"content="&sContent&""
	'<!-- 읽는 페이지에 같은 파라메터명으로 되어 있는게 있어서 혹시나 모를 일에 대비하여 파라메터명을 바꿔서 주고 받았슴. //-->

	Dim cooperatelist , i
	
		set cooperatelist = new CCooperate
	 	cooperatelist.FCPage = iCurrentpage
	 	cooperatelist.FPSize = iPageSize
	 	cooperatelist.FTeam = sSearchTeam
	 	cooperatelist.FDoc_Status = sDoc_Status
	 	cooperatelist.FDoc_Type = sDoc_Type
	 	cooperatelist.FDoc_AnsOX = sDoc_AnsOX
	 	cooperatelist.FDoc_MineOX = sSearchMine
	 	cooperatelist.FDoc_UserName = sUserName
	 	cooperatelist.FDoc_Searching = sSearching
	 	cooperatelist.FDoc_Content = sContent
		arrList = cooperatelist.fnGetCooperateList
		iTotCnt = cooperatelist.FTotCnt
	
	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1
%>

<script language="javascript">
function code_manage()
{
	window.open('PopManageCode.asp','coopcode','width=410,height=570');
}
function goWrite(didx)
{
	location.href = "cooperate_read.asp?didx="+didx+"<%=vParam%>";
}
function jsGoPage(iP){
	document.frmpage.iC.value = iP;
	document.frmpage.submit();
}
function mine()
{
	if(!(document.frm.onlyminechk.checked))
	{
		document.frm.onlymine.value = "x";
	}
	else
	{
		document.frm.onlymine.value = "o";
	}
}
function issystem(value)
{
}

function popCooperate(){
	 var winCooperate = window.open("/admin/cooperate/popIndex.asp","popCooperate","width="+screen.availWidth+", height="+ screen.availHeight +",resizable=yes, scrollbars=yes"); 
	 winCooperate.focus();
}

//전자결재 품의서 등록 - 업무협조고유번호(scmidx) 
function jsRegEapp(scmidx){ 
	var winEapp = window.open("/admin/approval/eapp/regeapp.asp","popE","width=1000,height=600,scrollbars=yes");
	document.frmEapp.iSL.value = scmidx;   
	document.frmEapp.target = "popE";
	document.frmEapp.submit();
	winEapp.focus();
}

//전자결재 품의서 내용보기
function jsViewEapp(reportidx,reportstate){
	var winEapp = window.open("/admin/approval/eapp/modeapp.asp?iRM=M01"+reportstate+"&iridx="+reportidx,"popE","");
	winEapp.focus();
}
</script>


<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<a href="/admin/notice/cooperate/?menupos=<%=g_MenuPos%>"><u><b>[업무협조리스트]</b></u></a>&nbsp;&nbsp;&nbsp;<a href="/admin/notice/cooperate/my_cooperate.asp?menupos=<%=g_MenuPos%>">[보낸 업무협조]</a>
		</td>
		<td align="right">
		</td>
	</tr>
</table>

<p>
<form name="frmEapp" method="post" action="/admin/approval/eapp/regeapp.asp">
<input type="hidden" name="tC" value="">
<input type="hidden" name="ieidx" value="38">  
<input type="hidden" name="iSL" value="">
</form>
<form name="frm" action="index.asp" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%=g_MenuPos%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<table class="a">
		<tr>
			<td>
				<%
				If g_TeamJang = "o" Then
					Dim vSelect
					vSelect = vSelect & "<select name='search_team' onchange='frm.submit()' class='select'>" & vbCrLf
					vSelect = vSelect & "	<option value=''>-팀선택-</option>" & vbCrLf
					vSelect = vSelect & "	<option value='9,10' "
					If sSearchTeam = "9,10" Then
						vSelect = vSelect & "selected"
					End If
					vSelect = vSelect & "	>운영관리팀</option>" & vbCrLf
					vSelect = vSelect & "	<option value='11,12,14,16,21' "
					If sSearchTeam = "11,12,14,16,21" Then
						vSelect = vSelect & "selected"
					End If
					vSelect = vSelect & "	>텐바이텐온라인사업팀</option>" & vbCrLf
					vSelect = vSelect & "	<option value='7' "
					If sSearchTeam = "7" Then
						vSelect = vSelect & "selected"
					End If
					vSelect = vSelect & "	>시스템팀</option>" & vbCrLf
					vSelect = vSelect & "	<option value='8,20' "
					If sSearchTeam = "8,20" Then
						vSelect = vSelect & "selected"
					End If
					vSelect = vSelect & "	>경영관리팀</option>" & vbCrLf
					vSelect = vSelect & "	<option value='15,19' "
					If sSearchTeam = "15,19" Then
						vSelect = vSelect & "selected"
					End If
					vSelect = vSelect & "	>아이띵소팀</option>" & vbCrLf
					vSelect = vSelect & "	<option value='13,18' "
					If sSearchTeam = "13,18" Then
						vSelect = vSelect & "selected"
					End If
					vSelect = vSelect & "	>오프라인팀</option>" & vbCrLf
					vSelect = vSelect & "	<option value='17' "
					If sSearchTeam = "17" Then
						vSelect = vSelect & "selected"
					End If
					vSelect = vSelect & "	>패션사업팀</option>" & vbCrLf
					vSelect = vSelect & "</select>" & vbCrLf
					
					Response.Write vSelect
				End If
				%>
				&nbsp;
				처리상태:
				<%=CommonCode("w","doc_status","s"&sDoc_Status)%>
		     	&nbsp;
		     	요청구분:
				<%=CommonCode("w","doc_type",sDoc_Type)%>
		     	&nbsp;
		     	답변여부:
		     	<select name="ans_ox" class='select'>
			     	<option value='' selected>전체</option>
			     	<option value='x' <% If sDoc_AnsOX = "x" Then %>selected<% End If %>>미답변</option>
			     	<option value='o' <% If sDoc_AnsOX = "o" Then %>selected<% End If %>>답변완료</option>
		     	</select>
			</td>
			<td rowspan="3" style="padding:0 0 0 30px;" valign="top"><input type="submit" value=" 검  색 " class="button" style="width:70px;height:50px;" onfocus="this.blur();"></td>
		</tr>
		<tr>
			<td>
				보낸사람이름 : <input type="text" name="username" value="<%=sUserName%>" size="10">&nbsp;&nbsp;&nbsp;
				<select name="searching" class="select">
					<option value="">-선택-</option>
					<option value="doc_idx" <%=CHKIIF(sSearching="doc_idx","selected","")%>>협조No</option>
					<option value="title" <%=CHKIIF(sSearching="title","selected","")%>>제목</option>
					<option value="content" <%=CHKIIF(sSearching="content","selected","")%>>내용</option>
				</select>
				<input type="text" name="content" value="<%=sContent%>" size="60">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>
		<input type="hidden" name="onlymine" value="<%=sSearchMine%>">
		<label id="onlymine"><input type="checkbox" name="onlyminechk" onClick="mine()" value="o" <% If sSearchMine = "o" Then %>checked<% End If %>>나의 작업만 보기</label>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		※ 제목, 내용 검색시 상당히 느릴 수 있습니다.
	</td>
</tr>
</table>
</form>
<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="신규등록" onClick="location.href='cooperate_write.asp?menupos=<%=g_MenuPos%>&iC=<%=iCurrentpage%>'">
	</td>
	<td align="right">
		<%
		If CInt(session("ssAdminLsn")) = 1 AND CInt(session("ssAdminPsn")) = 7 Then
			Response.Write "<input type='button' class='button' value='코드관리' onClick='code_manage()'>&nbsp;"
		End If
		%>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			검색결과 : <b><%= iTotCnt %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td width="60" align="center">협조문NO</td>
		<td width="60" align="center">보낸사람</td>
		<td>제목</td>
		<!--<td width="120">구분</td>//-->
		<td width="70">작업자</td>
		<td width="150" align="center">참조</td>
		<td width="70" align="center">중요도</td>
		<td width="70" align="center">등록일</td>
		<td width="80" align="center">처리상태</td>
		<td width="60" align="center">답변여부</td>
		<td  align="center">결재내역</td> 
	</tr>
<%
	IF isArray(arrList) THEN
		For intLoop =0 To UBound(arrList,2)
%>
	<tr align="center" bgcolor="#FFFFFF" height="30" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'">
		<td  style="cursor:pointer" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=arrList(0,intLoop)%></td>
		<td style="cursor:pointer" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=arrList(7,intLoop)%></td>
		<td align="left" style="cursor:pointer" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=db2html(arrList(1,intLoop))%></td>
		<!--<td><%=CommonCode("v","doc_type",arrList(2,intLoop))%></td>//-->
		<td width="70" align="center" style="cursor:pointer" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=arrList(9,intLoop)%></td>
		<td width="150" align="left" style="cursor:pointer" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=arrList(10,intLoop)%></td>
		<td style="cursor:pointer" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=CommonCode("v","doc_important",arrList(3,intLoop))%></td>
		<td style="cursor:pointer" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=FormatDatetime(arrList(6,intLoop),2)%></td>
		<td style="cursor:pointer" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=CommonCode("v","doc_status",arrList(5,intLoop))%></td>
		<td style="cursor:pointer" onClick="goWrite('<%=arrList(0,intLoop)%>')"><%=arrList(8,intLoop)%></td>
			<td nowrap>  <!--'시스템개발 및 수정일때만 결재리스트 보여준다 2014.03.06 정윤정 추가-->
			<%IF (arrList(2,intLoop)="3" )  THEN %>
				<div>
				<% if  isNull(arrList(12,intLoop)) then %>
			  <font color="Gray">품의서 미작성</font>
				<% else %>
				<%=fnGetReportState(arrList(13,intLoop))%>&nbsp; 
				<input type="button" class="button"  value="품의서 보기" onClick="jsViewEapp('<%=arrList(12,intLoop)%>','<%= arrList(13,intLoop)%>');">
				<% end if%> 
			</div>
			<%IF arrList(13,intLoop) = 7 THEN%>
				<div style="padding:3px">
				<% if isNull(arrList(14,intLoop)) then %>
				<input type="button" class="button"  value="▷ 개발계획서 품의" onClick="jsRegEapp('<%=arrList(0,intLoop)%>');" >
				<% else %>
				<%=fnGetReportState(arrList(15,intLoop))%><br>
				<input type="button" class="button"  value="개발계획서 보기" onClick="jsViewEapp('<%=arrList(14,intLoop)%>','<%= arrList(15,intLoop)%>');">
				<% end if%> 
			 </div>
			 <%END IF%>
		<%END IF%>
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


	<!-- 페이징처리 -->
	
	
	<%
	iStartPage = (Int((iCurrentpage-1)/iPerCnt)*iPerCnt) + 1
	
	If (iCurrentpage mod iPerCnt) = 0 Then
		iEndPage = iCurrentpage
	Else
		iEndPage = iStartPage + (iPerCnt-1)
	End If
	%>
	
	<form name="frmpage" method="get" action="/admin/notice/cooperate/index.asp">
	<input type="hidden" name="iC" value="">
	<input type="hidden" name="menupos" value="<%=Request("menupos")%>">
	<input type="hidden" name="search_team" value="<%=sSearchTeam%>">
	<input type="hidden" name="doc_status" value="<%=sDoc_Status%>">
	<input type="hidden" name="doc_type" value="<%=sDoc_Type%>">
	<input type="hidden" name="ans_ox" value="<%=sDoc_AnsOX%>">
	<input type="hidden" name="onlymine" value="<%=sSearchMine%>">
	<input type="hidden" name="username" value="<%=sUserName%>">
	<input type="hidden" name="searching" value="<%=sSearching%>">
	<input type="hidden" name="content" value="<%=sContent%>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			<% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
			<% else %>[pre]<% end if %>
	        <%
				for ix = iStartPage  to iEndPage
					if (ix > iTotalPage) then Exit for
					if Cint(ix) = Cint(iCurrentpage) then
			%>
				<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="red">[<%=ix%>]</font></a>
			<%		else %>
				<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
			<%
					end if
				next
			%>
	    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
			<% else %>[next]<% end if %>
		</td>
	</tr>
    </form>
</table>

<%
	set cooperatelist = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
