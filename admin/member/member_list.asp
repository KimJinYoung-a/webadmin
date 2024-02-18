<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/MemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%
	Dim page, SearchKey, SearchString, isUsing, part_sn, research
	Dim puserdiv, ilevel_sn, criticinfouser, posit_sn, job_sn

	page        = requestCheckvar(Request("page"),10)
	isUsing     = requestCheckvar(Request("isUsing"),10)
	SearchKey   = requestCheckvar(Request("SearchKey"),32)
	SearchString = Request("SearchString")
	part_sn     = requestCheckvar(Request("part_sn"),10)
	research    = requestCheckvar(Request("research"),10)
	puserdiv    = requestCheckvar(Request("puserdiv"),10)
	ilevel_sn   = requestCheckvar(Request("ilevel_sn"),10)
	criticinfouser = requestCheckvar(Request("criticinfouser"),10)
	posit_sn    = requestCheckvar(Request("posit_sn"),10)
	job_sn      = requestCheckvar(Request("job_sn"),10)
	
	if isUsing="" and research="" then isUsing="Y"
	if page="" then page=1

	'// 로그인정보(등급)에 따라 기본 부서 설정(마스터 이상:2 및 시스템팀:7 제외)
	if Not(session("ssAdminLsn")<=2 or session("ssAdminPsn")=7) then
		part_sn = session("ssAdminPsn")
	end if 

	'// 내용 접수
	dim oMember, lp
	Set oMember = new CMember

	oMember.FPagesize = 20
	oMember.FCurrPage = page
	oMember.FRectsearchKey = searchKey
	oMember.FRectsearchString = searchString
	oMember.FRectisUsing = isUsing
	oMember.FRectpart_sn = part_sn	
	oMember.FRectuserdiv = puserdiv
	oMember.FRectLevelsn = ilevel_sn
	oMember.FRectPositsn = posit_sn
	oMember.FRectJobsn   = job_sn
	
	oMember.FRectcriticinfouser = criticinfouser
	oMember.GetMemberList
	
	
	dim oaddlevel,jj
	
%>
<!-- 검색 시작 -->
<script language="javascript">
<!--
// 신규 사용자 등록
	function AddItem()
	{
	    alert('사용 불가 메뉴');
	    return;
		//window.open("pop_Member_add.asp","popAddIem","width=378,height=410,scrollbars=yes");
	}

	// 사용자 수정/삭제
	function ModiItem(empno)
	{
		//window.open("pop_member_add.asp?id="+uid,"popModiIem","width=378,height=410,scrollbars=yes");
		var w = window.open("/admin/member/tenbyten/pop_member_modify.asp?sEPN="+empno,"ModiItem","width=700,height=800,scrollbars=yes");
		w.focus();
	}
	
	//어드민 권한관리
	function jsMngAuth(empno){
		var w = window.open("/admin/member/tenbyten/popAdminAuth.asp?sEPN="+empno,"popAuth","width=700,height=300,scrollbars=yes");
		w.focus();
	}

	// 페이지 이동
	function goPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.submit();
	}

//-->
</script>
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">

<tr height="25" align="center" bgcolor="#FFFFFF" >
    <td rowspan="2" width="50" height="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
	    등급(어드민권한)
	    <%=printLevelOption("ilevel_sn", ilevel_sn)%> /
	    
	    <% if (FALSE) then %>
	    (기준)권한
	    <% call DrawAuthBoxSimple("puserdiv",puserdiv,"") %> / 
	    <% end if %>
	    
		<% if session("ssAdminLsn")<=2 then %>
		부서
		<%=printPartOption("part_sn", part_sn)%> /
		<% end if %>		
		
		<!--
		직급:
		<%=printPositOptionIN90("posit_sn", posit_sn)%> /
		-->
		직책:
		<%=printJobOption("job_sn", job_sn)%>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr height="25" align="center" bgcolor="#FFFFFF" >
    <td align="left">
        개인정보취급권한자
		<select name="criticinfouser">
			<option value="">전체</option>
			<option value="1" <%=CHKIIF(criticinfouser="1","selected","")%> >권한자</option>
			<option value="0" <%=CHKIIF(criticinfouser="0","selected","")%> >비권한자</option>
		</select> /
		사용여부
		<select name="isUsing">
			<option value="">전체</option>
			<option value="Y">사용</option>
			<option value="N">삭제</option>
		</select> /
		검색
		<select name="SearchKey">
			<option value="">::구분::</option>
			<option value="userid">아이디</option>
			<option value="username">사용자명</option>
		</select>
		<script language="javascript">		
			document.frm.isUsing.value="<%= isUsing %>";
			document.frm.SearchKey.value="<%= SearchKey %>";
		</script>
		<input type="text" name="SearchString" size="12" value="<%=SearchString%>">
    </td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<p>
<!-- 상단 띠 시작 -->
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr><td height="1" colspan="15" bgcolor="#BABABA"></td></tr>
<tr height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="right">
		<table width="100%" border=0 cellspacing=0 cellpadding=0 class="a">
		<tr>
			<td>총 <%=oMember.FtotalCount%> 명</td>
			<td align="right">page : <%= page %>/<%=oMember.FtotalPage%></td>
		</tr>
		</table>
	</td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 상단 띠 끝 -->
<!-- 메인 목록 시작 -->
<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#E6E6E6">
	<td width="80">아이디</td>
	<td width="60">사번</td>
	<td width="60">이름</td>
	<!--<td width="50">직급</td>-->
	<td width="50">직책</td>
	<td width="190">부서</td>
	<td width="100">등급</td>
    <% if (FALSE) then %><td width="100">(기준)권한</td><% end if %>
	<td width="60">개인정보<br>취급</td>
	<td width="60">사용여부</td>
</tr>
<%
	if oMember.FResultCount=0 then
%>
<tr>
	<td colspan="11" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 사용자가 없습니다.</td>
</tr>
<%
	else
		for lp=0 to oMember.FResultCount - 1
		    '' 추가권한
		    if (oMember.FitemList(lp).FAddLevelCnt>0) then
		        set oaddlevel = new CPartnerAddLevel 
		        oaddlevel.FRectUserID = oMember.FitemList(lp).Fid
		        oaddlevel.FRectOnlyAdd= "on"
		        oaddlevel.getUserAddLevelList
		    end if
%>
<tr align="center" bgcolor="<% if oMember.FitemList(lp).FisUsing="Y" then Response.Write "#FFFFFF": else Response.Write "#F0F0F0": end if %>">
	<td><%=oMember.FitemList(lp).Fid%></td>
	<td><%=oMember.FitemList(lp).Fempno%></td>
	<td><a href="javascript:jsMngAuth('<%=oMember.FitemList(lp).Fempno%>')"><%=oMember.FitemList(lp).Fusername%></a></td>
	<!--<td><%=oMember.FitemList(lp).Fposit_name%></td>-->
	<td><%=oMember.FitemList(lp).Fjob_name%></td>
	<td><%=oMember.FitemList(lp).Fpart_name%>
	<% if (oMember.FitemList(lp).FAddLevelCnt>0) then %>
	    <% for jj=0 to oaddlevel.FresultCount-1 %>
	    <br><font color="blue"><%= oaddlevel.FitemList(jj).Fpart_name %></font>
	    <% next %>
	<% end if %>
	</td>
	<td><%=oMember.FitemList(lp).Flevel_name%>
	<% if (oMember.FitemList(lp).FAddLevelCnt>0) then %>
	    <% for jj=0 to oaddlevel.FresultCount-1 %>
	    <br><font color="blue"><%= oaddlevel.FitemList(jj).Flevel_name %></font>
	    <% next %>
	<% end if %>
	</td>
	<% if (FALSE) then %><td><%= oMember.FitemList(lp).getPartnerUserDivName %></td><% end if %>
	<td><%= GetCriticInfoUserLevelName(oMember.FitemList(lp).Fcriticinfouser)%></td>
	<td><%=oMember.FitemList(lp).FisUsing%></td>
</tr>
<%
            if (oMember.FitemList(lp).FAddLevelCnt>0) then
                set oaddlevel = Nothing
            end if
		next
		
	end if
%>
</table>
<!-- 메인 목록 끝 -->
<!-- 페이지 시작 -->
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr valign="bottom" height="25">
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr valign="bottom">
			<td align="center">
			<!-- 페이지 시작 -->
			<%
				if oMember.HasPreScroll then
					Response.Write "<a href='javascript:goPage(" & oMember.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
				else
					Response.Write "[pre] &nbsp;"
				end if

				for lp=0 + oMember.StartScrollPage to oMember.FScrollCount + oMember.StartScrollPage - 1

					if lp>oMember.FTotalpage then Exit for
	
					if CStr(page)=CStr(lp) then
						Response.Write " <font color='red'>[" & lp & "]</font> "
					else
						Response.Write " <a href='javascript:goPage(" & lp & ")'>[" & lp & "]</a> "
					end if

				next

				if oMember.HasNextScroll then
					Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
				else
					Response.Write "&nbsp; [next]"
				end if
			%>
			<!-- 페이지 끝 -->
			</td>
						 
		</tr>
		</table>
	</td>
</tr>

</table>
<!-- 페이지 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->