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
<!-- #include virtual="/admin/pcmain/inc_mainhead.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/main_ContentsManageCls.asp" -->
<%
'###############################################
' PageName : pcmain_manager.asp
' Discription : 사이트 메인 관리
' History : 2018-03-05 이종화
'###############################################

dim research,isusing, fixtype, linktype, poscode, validdate, prevDate, gubun, targetUser , prevTime
targetUser = "전체"
dim page,strParm, datediv
	isusing = request("isusing")
	research= request("research")
	poscode = request("poscode")
	fixtype = request("fixtype")
	page    = request("page")
	validdate= request("validdate")
	prevDate = request("prevDate")
	gubun = request("gubun")
	datediv = request("datediv")
	prevTime = request("prevTime")

	If gubun = "" Then
		gubun = "index"
	End If

	if ((research="") and (isusing="")) then
	    isusing = "Y"
	    validdate = "on"
	end if

	if prevTime = "" then prevTime = "00"

	if page="" then page=1
strParm = "isusing="&isusing&"&poscode="&poscode&"&fixtype="&fixtype&"&validdate="&validdate&"&prevDate="&prevDate&"&gubun="&gubun
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
		if (oposcode.FOneItem.Ffixtype="D" Or poscode="714" Or poscode="710") then
		'일자별일때 선택일 미리보기 날짜 지정
		oMainContents.FRectDateDiv = datediv

		end if
	oMainContents.Flinktype = oposcode.FOneItem.Flinktype
	oMainContents.FRectSelDate = prevDate
	oMainContents.FRectSelDateTime = prevTime
	end if
	oMainContents.GetMainContentsList

dim i


	'### 구분별 js 생성파일 ### (기존 index, 핑거스, 베스트어워드는 현재 사용중이어서 그대로 사용. 추후 변경예정.
	Dim vGubun
	If gubun = "my10x10" Then
		vGubun = "_my10x10"
	End IF
%>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript">
function NextPage(page){
    frm.page.value = page;
    frm.submit();
}

function popPosCodeManage(){
    var popwin = window.open('/admin/sitemaster/lib/popmainposcodeedit.asp','mainposcodeedit','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function AddNewMainContents(idx){
    var popwin = window.open('/admin/sitemaster/lib/popmaincontentsedit.asp?idx=' + idx+'&<%=strParm%>','mainposcodeedit','width=800,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function setDefault()
{
	frm.poscode.options[0].selected = true;
	frm.submit();
}
</script>

<!-- 상단 검색폼 시작 -->
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= request("menupos") %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
	    사용구분
		<select name="isusing" class="select">
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

		<% If  poscode="714" Or poscode="710" Then %>
		<select name="datediv" class="select">
		<option value="1" <% if datediv="1" then response.write "selected" %> >지정일
		<option value="2" <% if datediv="2" then response.write "selected" %> >시작일
		</select>
		<input id="prevDate" name="prevDate" value="<%=prevDate%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "prevDate", trigger    : "prevDate_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		<% Else %>
        &nbsp;&nbsp;
        지정일자 <input id="prevDate" name="prevDate" value="<%=prevDate%>" class="text" size="10" maxlength="10" /><img src="http://scm.10x10.co.kr/images/calicon.gif" id="prevDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "prevDate", trigger    : "prevDate_trigger",
				onSelect: function() {this.hide();}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		<% if prevDate <> "" then %>
		시간 <input type="input" name="prevTime" value="<%=prevTime%>" class="text" size="2" maxlength="2" /> 시~
		<% end if %>
		<% End If %>
		<br>
	    <input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >종료이전
	    <br>
	    ※ <font color="blue">그룹구분 : index - 10x10 메인</font><br/>
	    ※ <font color="blue">그룹구분 : PCbanner - 10x10 PC 배너</font><br/>
	    ※ <font color="blue">그룹구분 : MAbanner - 10x10 M/A 배너</font>
		</font>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검색">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
<!--<td><a href="http://www.10x10.co.kr/index_preview.asp?yyyymmdd=<%= Left(CStr(now()),10) %>" target="refreshFrm_Main">현재상태</a></td>-->
    <td colspan="13" align="right">
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
    <td>노출카테고리</td>	
    <td>링크<br>구분</td>
    <td>반영<br>주기</td>
    <td>시작일</td>
    <td>종료일</td>
    <td>사용여부</td>
	<td>노출등급</td>
    <td>우선순위</td>
    <td>등록자</td>
    <td>작업자</td>
</tr>
<%
	for i=0 to oMainContents.FResultCount - 1	
	
	if not isnull(oMainContents.FItemList(i).FtargetType) then
		Select Case cstr(oMainContents.FItemList(i).FtargetType)
			Case ""
				targetUser = "모든고객"	
			Case "0"
				targetUser = "white"
			Case "1"
				targetUser = "red"			
			Case "2"
				targetUser = "vip"			
			Case "3"
				targetUser = "vip gold"			
			Case "4"
				targetUser = "vvip"
			Case "4"
				targetUser = "vvip"
			Case "7"
				targetUser = "STAFF"
			Case "8"
				targetUser = "FAMILY"
			Case "9"
				targetUser = "BIZ"
			case "00"
				targetUser = "회원전체"
			case "99"
				targetUser = "비회원"
		end select
	else
		targetUser = "모든고객"
	end if 
%>
<% if (oMainContents.FItemList(i).IsEndDateExpired) or (oMainContents.FItemList(i).FIsusing="N") then %>
<tr bgcolor="#DDDDDD">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
    <td align="center"><%= "<a href=""javascript:AddNewMainContents('" & oMainContents.FItemList(i).Fidx & "');"">" & oMainContents.FItemList(i).Fidx & "</a>" %></td>
    <td align="center"><a href="?gubun=<%=gubun%>&poscode=<%= oMainContents.FItemList(i).Fposcode %>"><%= oMainContents.FItemList(i).Fposname %></a></td>
    <td align="center">
	<%
		'텍스트 링크타입이면 텍스트 표시 - 아니면 기존대로 이미지
		if oMainContents.FItemList(i).Flinktype="T" then
			Response.Write "<a href=""javascript:AddNewMainContents('" & oMainContents.FItemList(i).Fidx & "');"">" & oMainContents.FItemList(i).FlinkText & "</a>"
		Else
			If oMainContents.FItemList(i).Fposcode = "714" Then
	%>
    	<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><img src="<%= oMainContents.FItemList(i).Fcultureimage %>" border="0" width=160 height=238 alt="<%= oMainContents.FItemList(i).Faltname %>"></a>
    <% ElseIf oMainContents.FItemList(i).Fposcode = "706" Then %>   
			(이미지 <%=oMainContents.FItemList(i).Fbannertype%>개)&nbsp;&nbsp;<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><img src="<%= oMainContents.FItemList(i).getImageUrl %>" width=300 border="0"  alt="<%= oMainContents.FItemList(i).Faltname %>"></a>
	<% Else %>
    	<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><img src="<%= oMainContents.FItemList(i).getImageUrl %>" width=300 border="0"  alt="<%= oMainContents.FItemList(i).Faltname %>"></a>
		<% If oMainContents.FItemList(i).GetImageUrl2 <> "" Then %>
    		<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><img src="<%= oMainContents.FItemList(i).GetImageUrl2 %>" width=300 border="0"  alt="<%= oMainContents.FItemList(i).Faltname2 %>"></a>
		<% End If %>
		<% If oMainContents.FItemList(i).GetImageUrl3 <> "" Then %>
    		<a href="javascript:AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>');"><img src="<%= oMainContents.FItemList(i).GetImageUrl3 %>" width=300 border="0"  alt="<%= oMainContents.FItemList(i).Faltname3 %>"></a>
		<% End If %>		
    <%
			End If
    	end if
    %>
    </td>
    <td align="center"><%= oMainContents.FItemList(i).getDispCateListName %></td>
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
	<td align="center"><%= targetUser %></td>
    <td align="center">
    	<%
			response.write oMainContents.FItemList(i).forderidx
    	%>
    </td>
    <td align="center"><%= oMainContents.FItemList(i).Fregname %></td>
    <td align="center"><%= oMainContents.FItemList(i).Fworkername %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="13" align="center" height="30">
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
