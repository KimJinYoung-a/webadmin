<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 매뉴관리
' History : 서동석 생성
'			2021.10.19 한용민 수정(수정로그 저장)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%
Dim page, SearchKey, SearchString, pid, strUse, useSslYN, criticinfo, saveLog
dim part_sn, level_sn, lv1customerYN, lv2partnerYN, lv3InternalYN
	pid     = RequestCheckvar(Request("pid"),10)
	page    = RequestCheckvar(Request("page"),10)
	SearchKey = RequestCheckvar(Request("SearchKey"),32)
	SearchString = Request("SearchString")
	strUse =   RequestCheckvar(Request("strUse"),10)
	useSslYN = RequestCheckvar(Request("useSslYN"),10)
	criticinfo = RequestCheckvar(Request("criticinfo"),10)
	saveLog = RequestCheckvar(Request("saveLog"),10)
	part_sn = RequestCheckvar(Request("part_sn"),10)
	level_sn = RequestCheckvar(Request("level_sn"),10)
	lv1customerYN 	= requestCheckvar(request("lv1customerYN"),1)
	lv2partnerYN 	= requestCheckvar(request("lv2partnerYN"),1)
	lv3InternalYN 	= requestCheckvar(request("lv3InternalYN"),1)
	
	if page="" then	page=1
	if pid="" then pid=0
	if strUse="" then strUse="Y"


	'// 내용 접수
	dim oMenu, lp
	Set oMenu = new CMenuList

	oMenu.FPagesize = 100
	oMenu.FCurrPage = page
	oMenu.FRectsearchKey = searchKey
	oMenu.FRectsearchString = searchString
	oMenu.FRectPid = pid
	oMenu.FRectisUsing = strUse
	oMenu.FRectuseSslYN=useSslYN
	oMenu.FRectcriticinfo=criticinfo
	oMenu.FRectSaveLog = saveLog
	oMenu.FRectlv1customerYN = lv1customerYN
	oMenu.FRectlv2partnerYN = lv2partnerYN
	oMenu.FRectlv3InternalYN = lv3InternalYN
	oMenu.FRectPart_sn = part_sn
	oMenu.FRectLevel_sn = level_sn
	oMenu.GetMenuListNew
%>
<!-- 검색 시작 -->
<script language="javascript">
<!--
	// 페이지 이동
	function goPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.action="menu_list.asp";
		document.frm.submit();
	}

	// 하위메뉴 이동
	function goChild(pid)
	{
		document.frm.pid.value=pid;
		document.frm.action="menu_list.asp";
		document.frm.submit();
	}


	// 메뉴 상세정보(수정) 페이지 이동
	function goEdit(mid)
	{
	    //document.frm.mid.value=mid;
		//document.frm.page.value='<%= page %>';
		//document.frm.action="menu_edit.asp";
		//document.frm.submit();

	    var popwin=window.open('menu_edit.asp?mid='+mid,'popmenu_edit','width=1200,height=800,scrollbars=yes,resizable=yes');
	    popwin.focus();

	}

	// 신규등록 페이지로 이동
	function goAddItem()  {
		self.location="menu_add.asp?menupos=<%=menupos%>&pid=<%=pid%>";
	}
//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="" action="menu_list.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="mid" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			상위메뉴 <%=printRootMenuOption("pid",pid, "Action")%>
			&nbsp;
			권한분류 :
			<%= printPartOption("part_sn", part_sn) %>
			&nbsp;
			권한등급 :
			<%= printLevelOption("level_sn", level_sn) %>
		</td>

		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="submit" class="button_s" value="검색">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		사용여부 :
		<select class="select" name="strUse">
			<option value="all">전체</option>
			<option value="Y">사용</option>
			<option value="N">삭제</option>
		</select>
		&nbsp;
		검색 :
		<select class="select" name="SearchKey">
			<option value="">::구분::</option>
			<option value="id">메뉴번호</option>
			<option value="menuname">메뉴명</option>
		</select>
		<input type="text" class="text" name="SearchString" size="20" value="<%=SearchString%>">

		<script language="javascript">
			document.frm.SearchKey.value="<%=SearchKey%>";
			document.frm.strUse.value="<%=strUse%>";
		</script>

		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			메뉴등급 :
			<%' Call DrawSelectBoxCriticInfoMenu("criticinfo", criticinfo) %>
			<input type="checkbox" name="lv1customerYN" value="Y" <% if lv1customerYN = "Y" then %>checked<% end if %> >LV1(고객정보)
			<input type="checkbox" name="lv2partnerYN" value="Y" <% if lv2partnerYN = "Y" then %>checked<% end if %> >LV2(파트너정보)
			<input type="checkbox" name="lv3InternalYN" value="Y" <% if lv3InternalYN = "Y" then %>checked<% end if %> >LV3(내부정보)			
			&nbsp;
			SSL 여부 :
			<select class="select" name="useSslYN">
				<option value="">전체</option>
				<option value="Y" <%=CHKIIF(useSslYN="Y","selected","")%> >SSL 사용</option>
				<option value="N" <%=CHKIIF(useSslYN="N","selected","")%> >SSL 사용안함</option>
			</select>
			&nbsp;
			접속로그 저장 :
			<select class="select" name="saveLog">
				<option value="">전체</option>
				<option value="1" <%=CHKIIF(saveLog="1","selected","")%> >저장</option>
				<option value="0" <%=CHKIIF(saveLog="0","selected","")%> >저장안함</option>
			</select>
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<% if pid<>0 then %>
	<td><input type="button" class="button" value="메뉴루트" onClick="goChild(0)"></td>
	<% end if %>
	<td align="right">
		<input type="button" class="button" value="신규등록" onClick="goAddItem()">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
		검색결과 : <b><%=oMenu.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oMenu.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50">ID</td>
	<td>상위메뉴</td>
	<td>하위메뉴</td>
	<td>상위권한</td>
	<td>링크</td>
	<td>권한</td>
	<td width="30">순서</td>
	<td width="30">사용</td>
	<td width="30">LV1<br>고객<br>정보</td>
	<td width="40">LV2<br>파트너<br>정보</td>
	<td width="30">LV3<br>내부<br>정보</td>
	<!--td width="30">SSL</td-->
	<!--td width="30">로그<br>저장</td-->
	<td width="80">수정</td>
</tr>
<%
	if oMenu.FResultCount=0 then
%>
<tr>
	<td colspan="16" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 메뉴가 없습니다.</td>
</tr>
<%
	else
		for lp=0 to oMenu.FResultCount - 1
%>
<tr align="center" bgcolor="<% if oMenu.FitemList(lp).Fmenu_isUsing="Y" then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>">
	<td><%=oMenu.FitemList(lp).Fmenu_id%></td>
	<td align="left">
		&nbsp;
		<%
		response.Write "<a href='javascript:goChild(" & oMenu.FitemList(lp).Fmenu_id & ")'>" & oMenu.FitemList(lp).Fmenu_name_parent & "</a>"

		if Not(isNull(oMenu.FitemList(lp).Fmenu_cnt)) then
			response.Write "<span style='color:#AA5555;font-size:10px'> [" & oMenu.FitemList(lp).Fmenu_cnt & "]</span>"
		end if
		%>
	</td>
	<td align="left">
		&nbsp;
		<%
		response.Write oMenu.FitemList(lp).Fmenu_name
		%>
	</td>
	<td><%=oMenu.FitemList(lp).getOldMenuDivStr%></td>
	<td align="left"><%=oMenu.FitemList(lp).Fmenu_linkurl%></td>
	<td align="left"><%=getPartLevelInfo(oMenu.FitemList(lp).Fmenu_id, "list")%></td>
	<td><%=oMenu.FitemList(lp).Fmenu_viewIdx%></td>
	<td><%=oMenu.FitemList(lp).Fmenu_isUsing%></td>
	<!--td><%'GetCriticInfoMenuLevelName(oMenu.FitemList(lp).Fmenu_criticinfo) %></td-->
	<td><%=oMenu.FitemList(lp).Flv1customerYN%></td>
	<td><%=oMenu.FitemList(lp).Flv2partnerYN%></td>
	<td><%=oMenu.FitemList(lp).Flv3InternalYN%></td>
	<!--td><%'oMenu.FitemList(lp).Fmenu_useSslYN%></td-->
	<!--td><%'oMenu.FitemList(lp).Fmenu_saveLog%></td-->
	<td><input type="button" value="수정" class="button" onClick="goEdit(<%=oMenu.FitemList(lp).Fmenu_id%>)"></td>
</tr>
<%
		next
	end if
%>
<!-- 메인 목록 끝 -->
<!-- 페이지 시작 -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16" align="center">
	<!-- 페이지 시작 -->
	<%
		if oMenu.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & oMenu.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + oMenu.StartScrollPage to oMenu.FScrollCount + oMenu.StartScrollPage - 1

			if lp>oMenu.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>[" & lp & "]</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>[" & lp & "]</a> "
			end if

		next

		if oMenu.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
	%>
	<!-- 페이지 끝 -->
	</td>
</tr>
</table>
<!-- 페이지 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
