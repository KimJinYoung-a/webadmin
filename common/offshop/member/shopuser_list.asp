<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 직영매장어드민권한설정
' Hieditor : 2011.01.10 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/shopmaster/shopuser_cls.asp"-->

<%
dim omember , i , part_sn , page , adminyn ,SearchKey ,SearchString
	part_sn = request("part_sn")
	page = request("page")
	adminyn = request("adminyn")
	SearchKey = request("SearchKey")
	SearchString = request("SearchString")

if page="" then page=1
if adminyn = "" then adminyn = "Y"
			
set omember = new cshopuser_list
	omember.FPageSize = 50
	omember.FCurrPage = page
	omember.frectpart_sn = part_sn
	omember.frectadminyn = adminyn
	omember.frectSearchKey = SearchKey
	omember.frectSearchString = SearchString
	omember.getshopuser_list()
%>

<script language="javascript">
	
	function reg(page){
		frm.page.value = page;
		frm.submit();
	}
	
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="editor_no">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 부서:
		<%=printPartOption("part_sn", part_sn)%>
		&nbsp;&nbsp;	
		* 어드민사용여부 : <% Call drawSelectBoxUsingYN("adminyn",adminyn) %>
		&nbsp;&nbsp;
		* <select name="SearchKey" class="select">
			<option value="" <% if SearchKey = "" then response.write " selected" %>>::구분::</option>
			<option value="1" <% if SearchKey = "1" then response.write " selected" %>>아이디</option>
			<option value="2" <% if SearchKey = "2" then response.write " selected" %>>사용자명</option>
			<option value="3" <% if SearchKey = "3" then response.write " selected" %>>사번</option>
		</select>
		<input type="text" class="text" name="SearchString" size="17" value="<%=SearchString%>">
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<br>		
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

※ 직원별로 관리매장을 여러개 등록 하시면, 직원이 로그인후 오른쪽 상단 매뉴에서 해당 매장들을 선택하여 관리하실수 있습니다.
<br>대표담당매장은 첫 로그인시 자동으로 선택되는 매장을 말합니다, 반드시 지정 부탁드립니다.
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= omember.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= Page %> / <%= omember.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>부서</td>
	<td>사원번호</td>
	<td>ID</td>
	<td>이름</td>
	<td>대표매장(관리매장수)</td>
	<td>비고</td>
</tr>
<% if omember.fresultcount > 0 then %>
<% for i=0 to omember.fresultcount - 1 %>

<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>
	<td>
		<%= omember.FItemList(i).fpart_name %>
	</td>
	<td>
		<%= omember.FItemList(i).fempno %>
	</td>
	<td>
		<%= omember.FItemList(i).fid %>
	</td>		
	<td>
		<%= omember.FItemList(i).fcompany_name %>
	</td>
	<td align="left">
		<%
		if omember.FItemList(i).fshopfirst = "" or isnull(omember.FItemList(i).fshopfirst) then
			response.write "지정없음"
		else
			response.write omember.FItemList(i).fshopfirst&"/"&omember.FItemList(i).fshopname
		end if
		%>
		(<%= omember.FItemList(i).fshopcount %>개)
	</td>
	<td width=70>
		<input type="button" onclick="shopreg('<%= omember.FItemList(i).fempno %>');" value="수정" class="button">
	</td>	
</tr>   
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="20" align="center">
	<% if omember.HasPreScroll then %>
		<a href="javascript:reg('<%= omember.StartScrollPage-1 %>')">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + omember.StartScrollPage to omember.FScrollCount + omember.StartScrollPage - 1 %>
		<% if i>omember.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:reg('<%= i %>')">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if omember.HasNextScroll then %>
		<a href="javascript:reg('<%= i %>')">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
set omember = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->