<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%
Dim page, SearchKey, SearchString, pid, strUse, useSslYN, criticinfo
dim part_sn, level_sn, lv1customerYN, lv2partnerYN, lv3InternalYN

function getLevelStr(level_sn)
    getLevelStr = ""

	if IsNull(level_sn) then
		exit function
	end if

	''select case level_sn
	''	case 1
	''		getLevelStr = "관리자"
	''	case 2
	''		getLevelStr = "마스터"
	''	case 3
	''		getLevelStr = "파트선임"
	''	case 4
	''		getLevelStr = "파트구성원"
	''	case 5
	''		getLevelStr = "파트임시직"
	''	case 6
	''		getLevelStr = "매장점장"
	''	case 7
	''		getLevelStr = "매장직원"
	''	case 9
	''		getLevelStr = "개인정보조회"
	''	case else
	''		getLevelStr = "ERR(" & level_sn & ")"
	''end select

	select case level_sn
		case 1
			getLevelStr = "A"
		case 2
			getLevelStr = "B"
		case 3
			getLevelStr = "C"
		case 4
			getLevelStr = "D"
		case 5
			getLevelStr = "E"
		case 6
			getLevelStr = "F"
		case 7
			getLevelStr = "G"
		case 9
			getLevelStr = "H"
		case else
			getLevelStr = "ERR(" & level_sn & ")"
	end select
end function

	pid     = RequestCheckvar(Request("pid"),10)
	page    = RequestCheckvar(Request("page"),10)
	SearchKey = RequestCheckvar(Request("SearchKey"),32)
	SearchString = Request("SearchString")
	strUse =   RequestCheckvar(Request("strUse"),10)
	useSslYN = RequestCheckvar(Request("useSslYN"),10)
	criticinfo = RequestCheckvar(Request("criticinfo"),10)
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

	oMenu.FPagesize = 300
	oMenu.FCurrPage = page
	oMenu.FRectsearchKey = searchKey
	oMenu.FRectsearchString = searchString
	oMenu.FRectPid = pid
	oMenu.FRectisUsing = strUse
	oMenu.FRectuseSslYN=useSslYN
	oMenu.FRectcriticinfo=criticinfo
	oMenu.FRectlv1customerYN = lv1customerYN
	oMenu.FRectlv2partnerYN = lv2partnerYN
	oMenu.FRectlv3InternalYN = lv3InternalYN
	oMenu.FRectPart_sn = part_sn
	oMenu.FRectLevel_sn = level_sn
	oMenu.GetMenuPrivList
%>
<!-- 검색 시작 -->
<script language="javascript">
<!--
	// 페이지 이동
	function goPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.action="menu_priv_list.asp";
		document.frm.submit();
	}

	// 하위메뉴 이동
	function goChild(pid)
	{
		document.frm.pid.value=pid;
		document.frm.action="menu_priv_list.asp";
		document.frm.submit();
	}


	// 메뉴 상세정보(수정) 페이지 이동
	function goEdit(mid)
	{
	    //document.frm.mid.value=mid;
		//document.frm.page.value='<%= page %>';
		//document.frm.action="menu_edit.asp";
		//document.frm.submit();

	    var popwin=window.open('menu_edit.asp?mid='+mid,'popmenu_edit','width=900,height=700,scrollbars=yes,resizable=yes');
	    popwin.focus();

	}

	// 신규등록 페이지로 이동
	function goAddItem()  {
		self.location="menu_add.asp?menupos=<%=menupos%>&pid=<%=pid%>";
	}
//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="" action="menu_priv_list.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="mid" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			상위메뉴 <%=printRootMenuOption("pid",pid, "Action")%>
			&nbsp;
			권한분류 :
			<%= printPartOption("part_sn", part_sn) %>
			&nbsp;
			권한등급 :
			<%= printLevelOption("level_sn", level_sn) %>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="submit" class="button_s" value="검색">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		사용여부
		<select class="select" name="strUse">
			<option value="all">전체</option>
			<option value="Y">사용</option>
			<option value="N">삭제</option>
		</select>
		/ 검색
		<select class="select" name="SearchKey">
			<option value="">::구분::</option>
			<option value="id">메뉴번호</option>
			<option value="menuname">메뉴명</option>
		</select>
		<input type="text" class="text" name="SearchString" size="20" value="<%=SearchString%>">
		/ SSL 여부
		<select class="select" name="useSslYN">
			<option value="">전체</option>
			<option value="Y" <%=CHKIIF(useSslYN="Y","selected","")%> >SSL 사용</option>
			<option value="N" <%=CHKIIF(useSslYN="N","selected","")%> >SSL 사용안함</option>
		</select>
		/
		메뉴등급 :
		<%' Call DrawSelectBoxCriticInfoMenu("criticinfo", criticinfo) %>
		<input type="checkbox" name="lv1customerYN" value="Y" <% if lv1customerYN = "Y" then %>checked<% end if %> >LV1(고객정보)
		<input type="checkbox" name="lv2partnerYN" value="Y" <% if lv2partnerYN = "Y" then %>checked<% end if %> >LV2(파트너정보)
		<input type="checkbox" name="lv3InternalYN" value="Y" <% if lv3InternalYN = "Y" then %>checked<% end if %> >LV3(내부정보)
		<script language="javascript">
			document.frm.SearchKey.value="<%=SearchKey%>";
			document.frm.strUse.value="<%=strUse%>";
		</script>

		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<% if pid<>0 then %>
	<td><input type="button" class="button" value="메뉴루트" onClick="goChild(0)"></td>
	<% end if %>
	<td align="right">

	</td>
</tr>
</table>
<!-- 액션 끝 -->

<p>

* A : 관리자 권한 / B : 마스터 권한 / C : 파트선임 권한 / D : 파트구성원 권한 / E : 파트임시직 권한 / F : 매장점장 권한 / G : 매장직원 권한 / H : 내정보조회권한

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="50">
		검색결과 : <b><%=oMenu.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oMenu.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50">ID</td>
	<td>상위메뉴</td>
	<td>하위메뉴</td>
	<td width="80">기존권한</td>
	<td>부서전체</td>
	<td>기업문화</td>
	<td>마케팅</td>
	<td>입점제휴</td>
	<td>온라인MD운영</td>
	<td>온라인MD수입</td>
	<td>온라인WD</td>
	<td>컨텐츠</td>
	<td>오프라인본사</td>
	<td>오프라인직영점</td>
	<td>운영기획</td>
	<td>시스템</td>
	<td>물류</td>
	<td>CS</td>
	<td>재무회계</td>
	<td>인사총무</td>
	<td>관계사</td>
	<td>추가01</td>
	<td>추가02</td>
	<td>기타</td>
	<td width="50">순서</td>
	<td width="50">사용</td>
	<!--td width="50">SSL</td-->
	<!--td width="100">메뉴등급</td-->
	<td width="30">LV1<br>고객<br>정보</td>
	<td width="40">LV2<br>파트너<br>정보</td>
	<td width="30">LV3<br>내부<br>정보</td>
	<td>수정</td>
</tr>
<%
	if oMenu.FResultCount=0 then
%>
<tr>
	<td colspan="50" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 메뉴가 없습니다.</td>
</tr>
<%
	else
		for lp=0 to oMenu.FResultCount - 1
%>
<% if (oMenu.FitemList(lp).Fmenu_isUsing = "Y") then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
<% else %>
<tr align="center" bgcolor="<%= adminColor("gray") %>" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="<%= adminColor("gray") %>";>
<% end if %>
	<td><%=oMenu.FitemList(lp).Fmenu_id%></td>
	<td align="left">
		<%
		response.Write "<a href='javascript:goChild(" & oMenu.FitemList(lp).Fmenu_id & ")'>" & oMenu.FitemList(lp).Fmenu_name_parent & "</a>"
		%>
	</td>
	<td align="left">
		<%
		response.Write oMenu.FitemList(lp).Fmenu_name
		%>
	</td>
	<td><%= oMenu.FitemList(lp).getOldMenuDivStr %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn1) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn16) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn14) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn22) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn11) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn21) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn12) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn23) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn13) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn24) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn30) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn7) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn9) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn10) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn8) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn20) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn17) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn33) %></td>
	<td><%= getLevelStr(oMenu.FitemList(lp).Fmenu_part_sn25) %></td>
	<td><%= oMenu.FitemList(lp).Fmenu_part_sn_etc %></td>
	<td><%= oMenu.FitemList(lp).Fmenu_viewIdx %></td>
	<td><%= oMenu.FitemList(lp).Fmenu_isUsing %></td>
	<!--td><%'oMenu.FitemList(lp).Fmenu_useSslYN %></td-->
	<!--td><%'GetCriticInfoMenuLevelName(oMenu.FitemList(lp).Fmenu_criticinfo) %></td-->
	<td><%= oMenu.FItemList(lp).Flv1customerYN %></td>
	<td><%= oMenu.FItemList(lp).Flv2partnerYN %></td>
	<td><%= oMenu.FItemList(lp).Flv3InternalYN %></td>
	<td><input type="button" value="수정" class="button" onClick="goEdit(<%=oMenu.FitemList(lp).Fmenu_id%>)"></td>
</tr>
<%
		next
	end if
%>
<!-- 메인 목록 끝 -->
<!-- 페이지 시작 -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="50" align="center">
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
