<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 우수회원샵
' Hieditor : 2009.12.28 한용민 생성
'			 2022.07.06 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/specialshop/specialshop_cls.asp"-->

<%
dim ospecialshop ,i,page , id , status , isusing
	menupos = requestCheckVar(getNumeric(request("menupos")),10)
	page = requestCheckVar(getNumeric(request("page")),10)
	id = requestCheckVar(getNumeric(request("id")),10)
	status = requestCheckVar(request("status"),1)
	isusing = requestCheckVar(request("isusing"),1)
	if page = "" then page = 1

set ospecialshop = new cspecialshop_list
	ospecialshop.FPageSize = 20
	ospecialshop.FCurrPage = page
	ospecialshop.frectid = id
	ospecialshop.frectisusing = isusing
	ospecialshop.frectstatus = status
	ospecialshop.fspecialshop_list()
	
%>

<script type='text/javascript'>

// 등록&수정
function reg(id){
	var reg = window.open('/admin/shopmaster/specialshop/specialshop_edit.asp?id='+id,'reg','width=1200,height=600,scrollbars=yes,resizable=yes');
	reg.focus();
}

//상품 등록&수정
function regitem(id){
	var regitem = window.open('/admin/shopmaster/specialshop/specialshop_edititem.asp?id='+id,'regitem','width=1400,height=700,scrollbars=yes,resizable=yes');
	regitem.focus();
}

//이벤트상태 실서버 적용
function statuschange(){
	var statuschange = window.open('/admin/shopmaster/specialshop/specialshop_process.asp?mode=statuschange','statuschange','width=50,height=50,scrollbars=yes,resizable=yes');
	statuschange.focus();
}

//상품 실서버 적용
function itemupdate(){
	var itemupdate = window.open('/admin/shopmaster/specialshop/specialshop_process.asp?mode=itemupdate','itemupdate','width=50,height=50,scrollbars=yes,resizable=yes');
	itemupdate.focus();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method=get action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			ID:<input type="text" name="id" value="<%=id%>" size=5>
			&nbsp;상태:<% drawstatus "status" , status ,"" %>
			&nbsp;사용여부:<select name="isusing" value="<%=isusing%>">
				<option value="" <% if isusing = "" then response.write " selected" %>>선택</option>
				<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
			</select>						
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			
		</td>
	</tr>
</table>
</form>
<!-- 검색 끝 -->

<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="진행중인상품실서버적용" onclick="itemupdate();">
		<% if C_ADMIN_AUTH then %>
			<input type="button" class="button" value="[관리자권한]실서버상태값초기화(현재날짜기준으로초기화.되도록화요일에만사용권장)" onclick="statuschange();"><br>
		<% end if %>
	</td>
	<td align="right">	
		<input type="button" class="button" value="신규등록" onclick="reg('');">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= ospecialshop.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= ospecialshop.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">   		
		<td align="center">ID</td>
		<td align="center">테마</td>
		<td align="center">오픈일</td>
		<td align="center">종료일</td>
		<td align="center">상태</td>			
		<td align="center" width=150>비고</td>	
    </tr>
	<% if ospecialshop.FresultCount>0 then %>
		<% for i=0 to ospecialshop.FresultCount-1 %>
		<form action="" name="frmBuyPrc<%=i%>" method="get" style="margin:0px;">			
		
		<% if ospecialshop.FItemList(i).fisusing = "N" then %>    
		<tr align="center" bgcolor="#FFFFaa">
		<% else %>    
		<tr align="center" bgcolor="#FFFFFF">
		<% end if %>
			<td align="center">
				<%= ospecialshop.FItemList(i).fid %>			
			</td>
			<td align="center">
				<%= ReplaceBracket(ospecialshop.FItemList(i).Ftitle) %>			
			</td>
			<td align="center">
				<%= FormatDate(ospecialshop.FItemList(i).fopenDate,"0000-00-00") %>
			</td>
			<td align="center">
				<%
					If ospecialshop.FItemList(i).FendDate <> "" then
						Response.Write FormatDate(ospecialshop.FItemList(i).FendDate,"0000-00-00")
					end if
				%>
			</td>
			<td align="center">
				<%= ospecialshop.FItemList(i).fstatusstr %>
			</td>
			<td align="center">
				<input type="button" onclick="reg(<%= ospecialshop.FItemList(i).fid %>)" value="수정" class="button">
				<input type="button" onclick="regitem(<%= ospecialshop.FItemList(i).fid %>)" value="상품등록[<%= ospecialshop.FItemList(i).fitemcount %>개]" class="button">
			</td>
		</tr>   
		</form>
		<% next %>

		<tr height="25" bgcolor="FFFFFF">
			<td colspan="15" align="center">
				<% if ospecialshop.HasPreScroll then %>
					<span class="list_link"><a href="?page=<%= ospecialshop.StartScrollPage-1 %>&id=<%=id%>&status=<%=status%>">[pre]</a></span>
				<% else %>
				[pre]
				<% end if %>
				<% for i = 0 + ospecialshop.StartScrollPage to ospecialshop.StartScrollPage + ospecialshop.FScrollCount - 1 %>
					<% if (i > ospecialshop.FTotalpage) then Exit for %>
					<% if CStr(i) = CStr(ospecialshop.FCurrPage) then %>
					<span class="page_link"><font color="red"><b><%= i %></b></font></span>
					<% else %>
					<a href="?page=<%= i %>&id=<%=id%>&status=<%=status%>" class="list_link"><font color="#000000"><%= i %></font></a>
					<% end if %>
				<% next %>
				<% if ospecialshop.HasNextScroll then %>
					<span class="list_link"><a href="?page=<%= i %>&id=<%=id%>&status=<%=status%>">[next]</a></span>
				<% else %>
				[next]
				<% end if %>
			</td>
		</tr>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
</table>

<%
set ospecialshop = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
