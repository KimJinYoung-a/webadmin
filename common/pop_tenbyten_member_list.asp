<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 직원리스트
' History : 2017.04.10 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
	Dim page, SearchKey, SearchString, isUsing, part_sn, research, orderby
	Dim job_sn, posit_sn, continuous_service_year, employeeonly
	Dim iTotCnt,iPageSize, iTotalPage
	dim department_id, inc_subdepartment,nodepartonly

	iPageSize = 20
	page = requestCheckVar(Request("page"),10)
	isUsing = requestCheckVar(Request("isUsing"),10)
	SearchKey = requestCheckVar(Request("SearchKey"),32)
	SearchString = requestCheckVar(Request("SearchString"),32)
	part_sn = requestCheckVar(Request("part_sn"),10)
	job_sn = requestCheckVar(Request("job_sn"),10)
	posit_sn = requestCheckVar(Request("posit_sn"),10)
	research = requestCheckVar(Request("research"),2)

	department_id = requestCheckvar(Request("department_id"),10)
	inc_subdepartment = requestCheckvar(Request("inc_subdepartment"),1)
	nodepartonly = requestCheckvar(Request("nodepartonly"),1)

	orderby = requestCheckvar(Request("orderby"),10)

	if isUsing="" and research="" then isUsing="Y"
	if page="" then page=1
	'if posit_sn ="" then posit_sn = 99
	'// 내용 접수
	dim oMember, arrList,intLoop
	Set oMember = new CTenByTenMember

	oMember.FPagesize 	= iPageSize
	oMember.FCurrPage 	= page
	oMember.FSearchType 	= searchKey
	oMember.FSearchText 	= searchString
	oMember.Fstatediv 		= isUsing
	oMember.Fpart_sn 		= part_sn
	oMember.Fjob_sn 		= job_sn
	oMember.Fposit_sn 	= posit_sn
	oMember.Forderby 		= orderby

	oMember.Fdepartment_id 		= department_id
	oMember.Finc_subdepartment 	= inc_subdepartment
	oMember.FRectNoDepartOnly 	= nodepartonly

	arrList = oMember.fnGetMemberList
	iTotCnt = oMember.FTotCnt
	set oMember = nothing

	iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
%>
<!-- 검색 시작 -->
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
<!--
	// 페이지 이동
	function jsGoPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.submit();
	}
	function mywork_update(emp){
		var popwin = window.open('pop_mywork_update.asp?empno=' + emp,'pop','width=500,height=200,scrollbars=yes,resizable=yes');
		popwin.focus();
	}

	$(function(){
		$(".colName").mouseenter(function(){
			$(this).find(".colPhoto").show();
		}).mousemove(function(e){
			$(this).find(".colPhoto").css("top",e.pageY-20).css("left",e.pageX+20)
		}).mouseleave(function(){
			$(this).find(".colPhoto").hide();
		});
	});
//-->
</script>
<style type="text/css">
body{ margin:0; }
</style>
<title>비상연락망</title>

<table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#FFFFFF">
<tr>
	<td width="30%">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="#CCCCCC">
		<tr>
			<td height="26" align="center" width="50%" bgcolor="#FFFFFF" style="cursor:pointer;" onClick="location.href='/common/pop_organization_chart.asp';"><font size="2">조직도</font></td>
			<td align="center" width="50%" bgcolor="#EDEDED" style="cursor:pointer;" onClick="location.href='/common/pop_tenbyten_member_list.asp';"><strong><font size="2">비상연락망</font></strong></td>
		</tr>
		</table>
	</td>
	<td width="70%" style="border-bottom: 1px solid #CCCCCC;"></td>
</tr>
<tr>
	<td colspan="2" height="10"></td>
</tr>
</table>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" height="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			부서NEW:
			<%= drawChSelectBoxDepartment("department_id", department_id,"") %>
			<input type="checkbox" name="inc_subdepartment" value="N" <% if (inc_subdepartment = "N") then %>checked<% end if %> > 하위 부서직원 제외&nbsp;
			<% if C_ADMIN_AUTH or C_PSMngPart then %>
			직급:
			<%=printPositOptionIN90("posit_sn", posit_sn)%>&nbsp;
			<% end if %>
			직책:
			<%=printJobOption("job_sn", job_sn)%>&nbsp;

			<br>

			사용여부:
			<select name="isUsing" class="select">
				<option value="">전체</option>
				<option value="Y">재직</option>
				<option value="N">퇴사</option>
			</select>
			&nbsp;
			검색:
			<select name="SearchKey" class="select">
				<option value="">::구분::</option>
				<option value="1">아이디</option>
				<option value="2">사용자명</option>
				<option value="3">사번</option>
				<option value="4">핸드폰</option>
			</select>
			<input type="text" class="text" name="SearchString" size="17" value="<%=SearchString%>">
			&nbsp;
			정렬:
			<select name="orderby" class="select">
				<option value="">이름</option>
				<option value="6">입사일(최근순)</option>
				<option value="5">퇴사일(최근순)</option>
				<!--<option value="2">직급</option>-->
				<option value="3">직책</option>
				<option value="4">내선</option>
			</select>
			&nbsp;
			<input type="checkbox" name="nodepartonly" value="Y" <% if (nodepartonly = "Y") then %>checked<% end if %> > 부서NEW 미지정만 (재직인 경우 부서 사용안함 포함)

			<script type="text/javascript">
				document.frm.isUsing.value="<%= isUsing %>";
				document.frm.SearchKey.value="<%= SearchKey %>";
				document.frm.orderby.value="<%= orderby %>";
			</script>

		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
</table>
<!-- 검색 끝 -->

<br><p>

<!-- 상단 띠 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%=iTotCnt%></b>
			&nbsp;
			페이지 : <b><%= page %> / <%=iTotalPage%></b>
			&nbsp;&nbsp;&nbsp;
			※ 이름에 마우스를 가져가면 사진이 나타납니다.
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="80">직책</td>
		<td width="130">이름</td>
		<% if C_ADMIN_AUTH or C_PSMngPart then %><td width="80">직급</td><% end if %>
		<td>부서</td>
		<td>담당업무</td>
		<td width="90">핸드폰번호</td>
		<td width="85">회사전화</td>
		<td width="35">내선</td>
		<td width="110">직통번호(070)</td>
		<td>이메일</td>
    </tr>
	<% if not isArray(arrList)  then %>
	<tr>
		<td colspan="15" align="center" bgcolor="#FFFFFF">등록(검색)된 사용자가 없습니다.</td>
	</tr>
	<% else %>

	<% for intLoop = 0 To UBound(arrList,2) %>
	<tr height=30 align="center" bgcolor="<% if  (arrList(15,intLoop)="Y") then Response.Write "#FFFFFF": else Response.Write adminColor("gray"): end if %>" onmouseout="this.style.backgroundColor='#FFFFFF'" onmouseover="this.style.backgroundColor='#F1F1F1'"  >
		<td><%=arrList(14,intLoop)%></td>
		<td class="colName" alt="<%=intLoop%>">
			<b><%=arrList(1,intLoop)%></b>
			<% if arrList(16,intLoop)<>"" then %>
			<div class="colPhoto" id="lyEmpPhoto<%=intLoop%>" style="background-color:white; position:absolute; left:10; top:10; z-index:1; display:none">
				<img src="<%=replace(arrList(16,intLoop),"http://webimage.10x10.co.kr","/webimage")%>" alt="<%=arrList(1,intLoop)%>" height="110" />
			</div>
			<% end if %>
		</td>
		<% if C_ADMIN_AUTH or C_PSMngPart then %><td><%=arrList(13,intLoop)%></td><% end if %>
		<td align="left"><%=arrList(27,intLoop)%></td>
		<td width="130"><%=arrList(20,intLoop)%><!--% If (session("ssAdminPsn") = "7" or (session("ssAdminPOsn") > 0 and session("ssAdminPOsn") =< "3")) or ((session("ssAdminPOsn") > "0" and session("ssAdminPOsn") < "6") and (session("ssAdminPsn") = arrList(5,intLoop)))  Then %--><!--input type ="button" name="mywork" value="수정" onclick="mywork_update('<%=arrList(0,intLoop)%>')" class="button"--><!--%End If%--></td>
		<td><%=arrList(17,intLoop)%></td>
		<td><%=arrList(9,intLoop)%></td>
		<td><b><%=arrList(10,intLoop)%></b></td>
		<td><b><%=arrList(11,intLoop)%></b></td>
		<td><%=arrList(8,intLoop)%></td>
	</tr>
	<% next %>

	<% end if %>
<!-- 메인 목록 끝 -->

<!-- 페이지 시작 -->
<%
Dim iStartPage,iEndPage,iX,iPerCnt
iPerCnt = 10

iStartPage = (Int((page-1)/iPerCnt)*iPerCnt) + 1

If (page mod iPerCnt) = 0 Then
	iEndPage = page
Else
	iEndPage = iStartPage + (iPerCnt-1)
End If
%>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
			    <tr valign="bottom" height="25">
			        <td valign="bottom" align="center">
			         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
					<% else %>[pre]<% end if %>
			        <%
						for ix = iStartPage  to iEndPage
							if (ix > iTotalPage) then Exit for
							if Cint(ix) = Cint(page) then
					%>
						<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong>[<%=ix%>]</strong></font></a>
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
			</table>
		</td>
	</tr>
	</form>
</table>
<!-- 페이지 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
