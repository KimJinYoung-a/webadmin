<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : 히치하이커 어드민 메인배너 리스트 페이지
'	History		: 2014.07.24 유태욱 생성
'#############################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/hitchhiker/hitchhiker_mainbannerCls.asp"-->

<%
Dim i
Dim FResultCount, iCurrentpage, iTotCnt
Dim Searchgubun, SearchUsing, validdate, research
	research= request("research")
	validdate= request("validdate")
	SearchUsing = request("SearchUsing")
	Searchgubun = request("Searchgubun")

	iCurrentpage = NullFillWith(requestCheckVar(Request("IC"),10),1)
if iCurrentpage="" then iCurrentpage=1

if ((research="") and (SearchUsing="")) then 
    SearchUsing = "Y"
    validdate = "on"
end if

Dim opart
set opart = new CAbouthitchhiker
	opart.FCurrPage = iCurrentpage
	opart.FPageSize = 15
	opart.FIsusing = SearchUsing
	opart.Fgubun = Searchgubun
	opart.FValiddate = validdate
	opart.fnGetHitchhikerList
iTotCnt = opart.FTotalCount
%>

<script type="text/javascript">
function conwrite(idx){
	var conwrite = window.open('/admin/hitchhiker/mainbanner/hitchhiker_mainbanner_write.asp?idx='+idx,'hitchhiker_mainbanner_write','width=800,height=768,scrollbars=yes,resizable=yes');
	conwrite.focus();
}
function searchFrm(p){
	frm.iC.value = p;
	frm.submit();
}
</script>
<!-- #include virtual="/admin/hitchhiker/inc_HichHead.asp"-->
<img src="/images/icon_arrow_link.gif"> <b>메인배너</b>
<% '검색---------------------------------------------------------------------------------------------------------- %>
<form name="frm" action="index.asp" method="get">
<input type="hidden" name="iC" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>" >
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%=admincolor("tablebg")%>">
	<tr align="center" bgcolor="#FFFFFF">
		<td lowsapn="2" width="100" bgcolor="<%=admincolor("gray")%>"> <b>검색조건</b> </td>
		<td align="left">
			<select name="SearchGubun">
				<option value ="" style="color:blue">구 분</option>
				<option value="1" <% If "1" = cstr(SearchGubun) Then%> selected <%End if%>>메인상단롤링배너_링크</option>
				<option value="2" <% If "2" = cstr(SearchGubun) Then%> selected <%End if%>>메인상단롤링배너_레이어팝업</option>
				<option value="3" <% If "3" = cstr(SearchGubun) Then%> selected <%End if%>>메인상단롤링배너_OnlyView</option>
				<option value="4" <% If "4" = cstr(SearchGubun) Then%> selected <%End if%>>메인상단롤링배너_모집&발간</option>
			</select>&nbsp;&nbsp;
			
			<b> 사 용 : </b>
			<select name="SearchUsing">
				<option value ="" style="color:blue">전 체</option>
				<option value="Y" <% If "Y" = cstr(SearchUsing) Then%> selected <%End if%>>Y</option>
				<option value="N" <% If "N" = cstr(SearchUsing) Then%> selected <%End if%>>N</option>
			</select>&nbsp;&nbsp;&nbsp;
			
			<input type="checkbox" name="validdate" <% if validdate="on" then response.write "checked" %> >종료이전
		</td>
		<td lowsapn="2" width=100 bgcolor="<%=admincolor("gray")%>">
			<input type="button" class="button" value="검색" onclick="searchFrm('');">
		</td>
	</tr>
</table>
</form>
<% '검색 끝------------------------------------------------------------------------------------------------------- %>
<br>
<% '신규입력버튼-------------------------------------------------------------------------------------------------- %>
<table width="100%" align="center">
	<tr>
		<td align="right"><input type="button" name="newBT" class="button" value="신규입력" onclick="conwrite('');"></td>
	</tr>
</table>
<% '신규입력버튼 끝----------------------------------------------------------------------------------------------- %>

<% '리스트-------------------------------------------------------------------------------------------------------- %>
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%=adminColor("tablebg")%>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="7" > <% '셀합병(colspan)7개 %>
			검색결과 : <b><%= iTotCnt %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%=adminColor("tabletop")%>" height="30">
		<td width="5%"><b>번호</b></td>
		<td width="5%"><b>구분</b></td>
		<td width="25%"><b>이미지</b></td>
		<td width="5%"><b>사용여부</b></td>
		<td width="5%"><b>우선순위</b></td>
		<td width="15%"><b>시작일/종료일</b></td>
		<td width="15%"><b>등록일</b></td>
	</tr>
	
	<% if opart.FResultCount > 0 then %>
	
		<% for i = 0 to opart.FResultCount - 1 %>
		<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background='#FFFFFF'; height="30"> 
			<td style="cursor:hand"  onclick="conwrite('<%= opart.FItemList(i).Fidx %>');"><%= opart.FItemList(i).Fidx %></td> <% '번호(idx) %>
			
			<td><%= getHitchhikerGubun(opart.FItemList(i).FReqgubun) %></td> <% '구분(PC,MOBILE,MOVIE %>
			
			<td><img src="<%= opart.FItemList(i).FReqcon_viewthumbimg %>" width="300" height="100"></td> <% '썸네일 %>
	
			<td><%= opart.FItemList(i).FReqIsusing %></td> <% '사용여부 %>
			
			<td><%= opart.FItemList(i).FReqsortnum %></td> <% '우선순위 %>
	
			<td align="left"> <% '시작일,종료일 %>
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
			<td><% Response.Write left(opart.FItemList(i).FreqRegdate,22) %></td> <% '등록일 %>
		</tr>
		<% next %>
		
		<% '페이징처리----------------------------------------- %>
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
		<% '페이징처리 끝------------------------------------------ %>
	<% else %>	
		<tr>
			<td colspan=7 align="center">
				검색결과가 없습니다.
			</td>
		</tr>
	<% end if %>
</table>
<% '리스트 끝----------------------------------------------------------------------------------------------- %>
<%
set opart = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->