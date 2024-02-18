<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : HITCHHIKER ADMIN(페이지관리->이슈영역)
'	History		: 2014.07.09 유태욱 생성
'#############################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/hitchhiker/about_hitchhikerCls.asp"-->

<%
Dim FResultCount, iTotCnt, idx
dim iCurrentpage
dim SearchTitle, SearchUsing, SearchGubun, validdate, research
dim i
	idx = request("idx")
	research= request("research")
	validdate= request("validdate")
	SearchTitle = request("evt_title")
	SearchUsing = request("SearchUsing")
	SearchGubun = request("SearchGubun")
	iCurrentpage = NullFillWith(requestCheckVar(Request("IC"),10),1)
	
if iCurrentpage="" then iCurrentpage=1
	
if ((research="") and (SearchUsing="")) then 
    SearchUsing = "Y"
    validdate = "on"
end if

dim opart
set opart = new CAbouthitchhiker
	opart.FCurrPage = iCurrentpage
	opart.FPageSize = 15
	opart.FEvt_title = SearchTitle
	opart.FIsusing = SearchUsing
	opart.Fgubun = SearchGubun
	opart.FValiddate = validdate
	opart.fnGetHitchhikerList
iTotCnt = opart.FTotalCount
%>

<script type="text/javascript">
function hicwrite(idx){
	var hicwrite = window.open('/admin/hitchhiker/issuearea/about_hitchhiker_write.asp?idx='+idx,'about_hitchhiker_write','width=650,height=530,scrollbars=yes,resizable=yes');
	hicwrite.focus();
}

function searchFrm(p){
	frm.iC.value = p;
	frm.submit();
}

</script>
<!-- #include virtual="/admin/hitchhiker/inc_HichHead.asp"-->
<img src="/images/icon_arrow_link.gif"> <b>이슈영역</b>
<!--검색------------------------------------------------------------------------------------------------->
<form name="frm" action="about_list.asp" method="get">
<input type="hidden" name="iC" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>" >
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%=admincolor("tablebg")%>">
	<tr align="center" bgcolor="#FFFFFF">
		<td lowsapn="2" width="100" bgcolor="<%=admincolor("gray")%>"> <b>검색조건</b> </td>
		<td align="left">
			<select name="SearchGubun">
				<option value ="" style="color:blue">구 분</option>
				<option value="1" <% If "1" = cstr(SearchGubun) Then%> selected <%End if%>>발간</option>
				<option value="2" <% If "2" = cstr(SearchGubun) Then%> selected <%End if%>>에디터모집</option>
				<option value="3" <% If "3" = cstr(SearchGubun) Then%> selected <%End if%>>기타</option>
			</select>
			&nbsp;&nbsp;
			<b> 사 용 : </b>
			<select name="SearchUsing">
				<option value ="" style="color:blue">전 체</option>
				<option value="Y" <% If "Y" = cstr(SearchUsing) Then%> selected <%End if%>>Y</option>
				<option value="N" <% If "N" = cstr(SearchUsing) Then%> selected <%End if%>>N</option>
			</select>
			&nbsp;&nbsp;
			<b> 타이틀 : </b>
			<input type=text value ="<%= SearchTitle %>" name="evt_title" onKeydown="javascript:if(event.keyCode == 13) form.submit();">
			&nbsp;&nbsp;&nbsp;
			<input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >종료이전
		</td>
		<td lowsapn="2" width=100 bgcolor="<%=admincolor("gray")%>">
			<input type="button" class="button" value="검색" onclick="searchFrm('');">
		</td>
	</tr>
</table>
</form>
<!--검색끝----------------------------------------------------------------------------------------------->
<br>
<!--신규입력버튼---------------------------------------------------------------------------------------->
<table width="100%" align="center">
	<tr>
		<td align="right"><input type="button" name="newBT" class="button" value="신규입력" onclick="hicwrite('');"></td>
	</tr>
</table>
<!--신규입력버튼끝-------------------------------------------------------------------------------------->

<!--리스트----------------------------------------------------------------------------------------------->
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%=adminColor("tablebg")%>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="7">
			검색결과 : <b><%= iTotCnt %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%=adminColor("tabletop")%>" height="30">
		<td width="5%"><b>번호</b></td>
		<td width="10%"><b>구분</b></td>
		<td width="40%"><b>타이틀</b></td>
		<td width="5%"><b>사용</b></td>
		<td width="5%"><b>우선순위</b></td>
		<td width="15%"><b>시작일/종료일</b></td>
		<td width="23%"><b>등록일</b></td>
	</tr>
	
	<% if opart.FResultCount > 0 then %>
	
		<% for i = 0 to opart.FResultCount - 1 %>
		<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background='#FFFFFF'; height="30"> 
			<td style="cursor:hand"  onclick="hicwrite('<%= opart.FItemList(i).Fidx %>');"><%= opart.FItemList(i).Fidx %></td>
			
			<td><%= getHitchhikerGubun(opart.FItemList(i).FReqgubun) %></td><!--구분(발간,에디터모집,기타-->
			
			<td><%= opart.FItemList(i).FReqTitle %></td>
	
			<td><%= opart.FItemList(i).FReqIsusing %></td>
			
			<td><%= opart.FItemList(i).FReqSortnum %></td>
			
			<td align="left">
				<% 
					Response.Write "시작: "
					Response.Write replace(left(opart.FItemList(i).FReqSdate,10),"-",".") & " / " & Num2Str(hour(opart.FItemList(i).FReqSdate),2,"0","R") & ":" &Num2Str(minute(opart.FItemList(i).FReqSdate),2,"0","R")
					Response.Write "<br />종료: "
					Response.Write replace(left(opart.FItemList(i).FReqEdate,10),"-",".") & " / " & Num2Str(hour(opart.FItemList(i).FReqEdate),2,"0","R") & ":" & Num2Str(minute(opart.FItemList(i).FReqEdate),2,"0","R")
					Response.Write "<br />"
		
					if now() >=  opart.FItemList(i).FReqSdate AND NOW() <= opart.FItemList(i).FReqEdate then
						Response.write " <span style=""color:blue"">(오픈)</span>"
					elseif now() < opart.FItemList(i).FReqSdate then
						Response.write " <span style=""color:green"">(오픈예정)</span>"
					else
						Response.write " <span style=""color:red"">(종료)</span>"
					end if
					Response.Write "<br />"
				%>
			</td>
			<td><% Response.Write left(opart.FItemList(i).FReqregdate,22) %></td>
		</tr>
		<% next %>
		<!--페이징처리------------------------------------------>
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="15" align="center">
		       	<% if opart.HasPreScroll then %>
					<span class="list_link"><a href="javascript:searchFrm('<%= opart.StartScrollPage-1 %>')">[pre]</a></span> '&menupos=<%=menupos%>
				<% else %>
				[pre]
				<% end if %>
					<% for i = 0 + opart.StartScrollPage to opart.StartScrollPage + opart.FScrollCount - 1 %>
						<% if (i > opart.FTotalpage) then Exit for %>
						<% if CStr(i) = CStr(iCurrentpage) then %>
							<span class="page_link"><font color="red"><b><%= i %></b></font></span>
						<% else %>
							<a href="javascript:searchFrm('<%= i %>')" class="list_link"><font color="#000000"><%= i %></font></a>
						<% end if %>
					<% next %>
				<% if opart.HasNextScroll then %>
					<span class="list_link"><a href="javascript:searchFrm('<%= i %>')">[next]</a></span>
				<% else %>
				[next]
				<% end if %>
			</td>
		</tr>
		<!--페이징처리끝------------------------------------------>	
	<% else %>	
		<tr>
			<td colspan=7 align="center">
				검색결과가 없습니다.
			</td>
		</tr>
	<% end if %>
</table>
<!--리스트끝----------------------------------------------------------------------------------------------->
<%
set opart = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->