<%@ language=vbscript %>
<% option explicit %>
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
<!-- #include virtual="/lib/classes/sitemasterclass/main_ContentsManageCls.asp" -->
<%
dim research,isusing, fixtype, linktype, poscode, validdate, prevDate, gubun
dim page
	isusing = request("isusing")
	research= request("research")
	poscode = request("poscode")
	fixtype = request("fixtype")
	page    = request("page")
	validdate= request("validdate")
	prevDate = request("prevDate")
	gubun = request("gubun")
	
	If gubun = "" Then
		gubun = "gift"
	End If
	
	If gubun = "index" Then
		response.write "<script>alert('10x10메인관리는 [ON]사이트관리>>[메인]페이지관리에서 합니다.');history.back();</script>"
	End If

	if ((research="") and (isusing="")) then 
	    isusing = "Y"
	    validdate = "on"
	end if
	
	if page="" then page=1

dim oposcode
	set oposcode = new CMainContentsCode
	oposcode.FRectPosCode = poscode
	
	if (poscode<>"") then
	    oposcode.GetOneContentsCode
	end if

dim oMainContents
	set oMainContents = new CMainContents
	oMainContents.FPageSize = 10
	oMainContents.FCurrPage = page
	oMainContents.FRectIsusing = isusing
	oMainContents.FRectfixtype = fixtype
	oMainContents.FRectPosCode = poscode
	oMainContents.Fgubun = gubun
	oMainContents.FRectvaliddate = validdate
	if (poscode<>"") then
		if (oposcode.FOneItem.Ffixtype="D") then
		'일자별일때 선택일 미리보기 날짜 지정
		oMainContents.FRectSelDate = prevDate
		end if
	oMainContents.Flinktype = oposcode.FOneItem.Flinktype
	end if
	oMainContents.GetMainContentsList

dim i


	'### 구분별 js 생성파일 ### (기존 index, 핑거스, 베스트어워드는 현재 사용중이어서 그대로 사용. 추후 변경예정.
	Dim vGubun
	If gubun = "my10x10" Then
		vGubun = "_my10x10"
	End IF
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>
function NextPage(page){
    frm.page.value = page;
    frm.submit();
}

function popPosCodeManage(){
    var popwin = window.open('/admin/sitemaster/lib/popmainposcodeedit.asp','mainposcodeedit','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function AddNewMainContents(idx){
    var popwin = window.open('/admin/sitemaster/lib/popmaincontentsedit.asp?gubun=<%=gubun%>&idx=' + idx,'mainposcodeedit','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function AssignReal(){
    if (document.frm.poscode.value == ""){
		alert("적용위치를 선택해주세요");
		document.frm.poscode.focus();
	}
	else{
		 var popwin = window.open('','refreshFrm_Main','');
		 popwin.focus();
		 refreshFrm.target = "refreshFrm_Main";
		 refreshFrm.action = "<%=wwwUrl%>/chtml/make_main_contents_JS.asp?poscode=" + document.frm.poscode.value;
		 refreshFrm.submit();
	}
}

function AssignDailyReal(idx){
	 var popwin = window.open('','refreshFrm_Main','');
	 popwin.focus();
	 refreshFrm.target = "refreshFrm_Main";
	 
	 <% if (application("Svr_Info")	= "Dev") then %>
	 refreshFrm.action = "http://2013www.10x10.co.kr/chtml/gift/giftmain_linkbanner_make.asp?poscode=" + document.frm.poscode.value + "&idx=" + idx;
	 <% else %>
	 refreshFrm.action = "<%=wwwUrl%>/chtml/gift/giftmain_linkbanner_make.asp?poscode=" + document.frm.poscode.value + "&idx=" + idx;
	 <% end if %>
	 
	 refreshFrm.submit();
}

function AssignXmlAppl(term){
    if (!confirm('새로 반영하시겠습니까?')) return;
     
	 var popwin = window.open('','refreshFrm_Main','');
	 popwin.focus();
	 refreshFrm.target = "refreshFrm_Main";
	 <% if (application("Svr_Info")	= "Dev") then %>
	 refreshFrm.action = "http://2013www.10x10.co.kr/chtml/gift/main_make_xml.asp?poscode=" + document.frm.poscode.value + "&term=" + term;
	 <% else %>
	 refreshFrm.action = "http://www1.10x10.co.kr/chtml/gift/main_make_xml.asp?poscode=" + document.frm.poscode.value + "&term=" + term;
	 <% end if %>
	 refreshFrm.submit();
}

function setDefault()
{
	frm.poscode.options[0].selected = true;
	frm.submit();
}
</script>

<!-- 상단 검색폼 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">

	    사용구분
		<select name="isusing" class="select">
		<option value="">전체
		<option value="Y" <% if isusing="Y" then response.write "selected" %> >사용함
		<option value="N" <% if isusing="N" then response.write "selected" %> >사용안함
		</select>
		&nbsp;&nbsp;
		적용구분
		<% call DrawFixTypeCombo ("fixtype", fixtype, "") %>
		
		&nbsp;&nbsp;
		그룹구분
		<% call DrawGroupGubunCombo ("gubun", gubun, "onChange='setDefault()'") %>
		
		&nbsp;&nbsp;
		적용위치
		<% call DrawMainPosCodeCombo("poscode",poscode, "", gubun) %>
		<%
			if (poscode<>"") then
				if (oposcode.FOneItem.Ffixtype="D") then
		%>
        &nbsp;&nbsp;
        지정일자 <input id="prevDate" name="prevDate" value="<%=prevDate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "prevDate", trigger    : "prevDate_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		<%
				end if
			end if
		%>
		
		<br>
	    <input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >종료이전
	    <br>
	    ※ <font color="blue">그룹구분 : index - 10x10 메인</font>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검색">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
    <td></td>
    <td colspan="2">
	    <%
	    	if (poscode<>"") then
	    		if (oposcode.FOneItem.Ffixtype="R") AND gubun = "index" then
	    		'실시간 반영
	    %>
			        <a href="javascript:AssignRealApplALL('header');"><img src="/images/refreshcpage.gif" border="0"> Real 적용(실시간 반영주기 헤더전체)</a>
			        <a href="javascript:AssignRealApplALL('idx');"><img src="/images/refreshcpage.gif" border="0"> Real 적용(실시간 반영주기 인덱스전체)</a>
		<%
				elseif oposcode.FOneItem.Flinktype="F" or oposcode.FOneItem.Flinktype="B" then
					if (oposcode.FOneItem.Ffixtype="D") then
					'플래시 일자별 적용
		%>
						오늘을 포함하여 <input type="text" name="vTerm" value="1" size="1" class="text" style="text-align:right;">일간
						<a href="javascript:AssignFlashDailyReal('<%= poscode %>','<%=oposcode.FOneItem.Flinktype%>',document.all.vTerm.value);"><img src="/images/refreshcpage.gif" border="0"> Flash Real 적용(예약)</a>
		<%
					else
					'플래시 일반 적용
		%>
						<a href="javascript:AssignFlashReal('<%= poscode %>','<%=oposcode.FOneItem.Flinktype%>');"><img src="/images/refreshcpage.gif" border="0"> Flash Real 적용</a>
		<%
					end if
				elseif (oposcode.FOneItem.Ffixtype <> "D") and (oposcode.FOneItem.Ffixtype <> "R") and (oposcode.FOneItem.Flinktype <> "X") then
				'링크 등 일반
		%>
		    	    <!--<a href="javascript:AssignTest('<%= poscode %>');"><img src="/images/icon_search.jpg" border="0"> 미리보기</a> 
		    	    &nbsp;&nbsp;
		    	    <a href="javascript:AssignReal('<%= poscode %>');"><img src="/images/refreshcpage.gif" border="0"> Real 적용</a>//-->
	    <%
	    		elseif oposcode.FOneItem.Flinktype="X" Then
	    			if (oposcode.FOneItem.Ffixtype="D") then
		%>
						오늘을 포함하여 <input type="text" name="vTerm" value="1" size="1" class="text" style="text-align:right;">일간
						<a href="javascript:AssignXmlAppl(document.all.vTerm.value);"><img src="/images/refreshcpage.gif" border="0"> XML Real 적용(예약)</a>
		<%
					else
		%>
						<a href="javascript:AssignXmlAppl('');"><img src="/images/refreshcpage.gif" border="0"> XML Real 적용</a>
		<%
					end if
				end if
	    	end if
	    %>
    </td>
    <td colspan="10" align="right">
    	<% if C_ADMIN_AUTH then %>
		<input type="button" class="button" value="코드관리" onClick="popPosCodeManage();">&nbsp;
		<% end if %>
    	<a href="javascript:AddNewMainContents('0');"><img src="/images/icon_new_registration.gif" border="0" align="absmiddle"></a>
    </td>
</tr>
</table>
<!-- 액션 끝 -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%=oMainContents.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oMainContents.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>idx</td>
    <td>구분명</td>
    <td>이미지/텍스트</td>
    <td>링크<br>구분</td>
    <td>반영<br>주기</td>
    <td>시작일</td>
    <td>종료일</td>
    <td>사용여부</td>
    <td>우선순위</td>
    <td>등록자</td>
    <td>작업자</td>
    <td></td>
</tr>
<%
	for i=0 to oMainContents.FResultCount - 1
%>
<% if (oMainContents.FItemList(i).IsEndDateExpired) or (oMainContents.FItemList(i).FIsusing="N") then %>
<tr bgcolor="#DDDDDD">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
    <td align="center"><%= "<a href=""javascript:AddNewMainContents('" & oMainContents.FItemList(i).Fidx & "');"">" & oMainContents.FItemList(i).Fidx & "</a>" %></td>
    <td align="center"><a href="?gubun=<%=gubun%>&poscode=<%= oMainContents.FItemList(i).Fposcode %>"><%= oMainContents.FItemList(i).Fposname %></a></td>
    <td>
	<%
		'텍스트 링크타입이면 텍스트 표시 - 아니면 기존대로 이미지
		if oMainContents.FItemList(i).Flinktype="T" then
			Response.Write "<a href=""javascript:AddNewMainContents('" & oMainContents.FItemList(i).Fidx & "');"">" & oMainContents.FItemList(i).FlinkText & "</a>"
		Else
			'이미지 사이즈에 따라 표시(제한 300px)
			if oMainContents.FItemList(i).Fimagewidth>300 then
	%>
    	<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><img src="<%= oMainContents.FItemList(i).getImageUrl %>" border="0" width=300></a>
	<%		else %>
    	<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><img src="<%= oMainContents.FItemList(i).getImageUrl %>" border="0"></a>
    <%
    		end if
    	end if
    %>
    </td>
    <td align="center"><%= oMainContents.FItemList(i).getlinktypeName %></td>
    <td align="center"><%= oMainContents.FItemList(i).getfixtypeName %></td>
    <td align="center"><%= oMainContents.FItemList(i).FStartdate %></td>
    <td align="center">
    <% if (oMainContents.FItemList(i).IsEndDateExpired) then %>
    <font color="#777777"><%= Left(oMainContents.FItemList(i).FEnddate,10) %></font>
    <% else %>
    <%= Left(oMainContents.FItemList(i).FEnddate,10) %>
    <% end if %>
    </td>
    <td align="center"><%= oMainContents.FItemList(i).FIsusing %></td>
    <td align="center">
    	<%
    	'// 지정된 적용위치에만 우선순위 출력
    	'Select Case poscode
    	'	Case "400", "401", "402", "403", "404", "405", "420", "421", "428"
    			response.write oMainContents.FItemList(i).forderidx 
    	'end Select
    	%>
    </td>
    <td align="center"><%= oMainContents.FItemList(i).Fregname %></td>
    <td align="center"><%= oMainContents.FItemList(i).Fworkername %></td>
    <td>
    <% if (oMainContents.FItemList(i).Ffixtype="R") then %>   
    <a href="javascript:AssignRealAppl('<%= oMainContents.FItemList(i).Fidx %>');"><img src="/images/refreshcpage.gif" border="0"> Real 적용</a>
    <% elseif Not(oMainContents.FItemList(i).IsEndDateExpired or oMainContents.FItemList(i).FIsusing="N" or oMainContents.FItemList(i).Flinktype="F" or oMainContents.FItemList(i).Flinktype="B" or oMainContents.FItemList(i).Ffixtype="R") then %>
    <!--<a href="javascript:AssignDailyTest('<%= oMainContents.FItemList(i).Fidx %>');"><img src="/images/icon_search.jpg" border="0"> 미리보기</a> //-->
    	<% If oMainContents.FItemList(i).Flinktype <> "X" AND poscode <> "" Then %>
    		&nbsp;
    		<a href="javascript:AssignDailyReal('<%= oMainContents.FItemList(i).Fidx %>');"><img src="/images/refreshcpage.gif" border="0"> Real 적용</a>
    	<% Else %>
    		
    	<% End If %>
    <% else %>
    &nbsp;
    <% end if %> 
    </td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="15" align="center">
    <% if oMainContents.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oMainContents.StarScrollPage-1 %>');">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oMainContents.StarScrollPage to oMainContents.FScrollCount + oMainContents.StarScrollPage - 1 %>
		<% if i>oMainContents.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>');">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oMainContents.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>');">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>

<%
set oposcode = Nothing
set oMainContents = Nothing
%>

<form name="refreshFrm" method="post">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->