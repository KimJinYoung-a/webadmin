<%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : 핑거스 모바일 상단 메인 배너
'	History		: 2016.07.29 유태욱 생성
'#############################################################
%>
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/academy/mobile/main_banner/academy_mobile_mainbannerCls.asp"-->

<%
Dim i
Dim FResultCount, iCurrentpage, iTotCnt
Dim Searchgubun, SearchUsing, validdate, research
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2,nowdate, datesearch


'==============================================================================
yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1 = RequestCheckvar(request("mm1"),2)
dd1 = RequestCheckvar(request("dd1"),2)
yyyy2 = RequestCheckvar(request("yyyy2"),2)
mm2 = RequestCheckvar(request("mm2"),2)
dd2 = RequestCheckvar(request("dd2"),2)
datesearch = RequestCheckvar(request("datesearch"),1)

	research= RequestCheckvar(request("research"),2)
	validdate= RequestCheckvar(request("validdate"),2)
	SearchUsing = RequestCheckvar(request("SearchUsing"),1)
	Searchgubun = RequestCheckvar(request("Searchgubun"),1)

	iCurrentpage = NullFillWith(requestCheckVar(Request("IC"),10),1)
if iCurrentpage="" then iCurrentpage=1

if ((research="") and (SearchUsing="")) then 
    SearchUsing = "Y"
    validdate = "on"
end if

Dim opart
set opart = new CAcademyMobileMainBanner
	opart.FCurrPage = iCurrentpage
	opart.FPageSize = 15
	opart.FIsusing = SearchUsing
	opart.Fgubun = Searchgubun
	opart.FValiddate = validdate
	If yyyy1 <> "" And datesearch="Y" Then
	opart.FRectSearchSDate = yyyy1 + "-" + mm1 + "-" + dd1
	End If
	If yyyy2 <> "" And datesearch="Y" Then
	opart.FRectSearchEDate = yyyy2 + "-" + mm2 + "-" + dd2
	End if
	opart.fnGetAcademyMobileMainBannerList
iTotCnt = opart.FTotalCount


if yyyy1="" Or yyyy2="" then
	nowdate = CStr(Now)
	nowdate = DateSerial(Left(nowdate,4), CLng(Mid(nowdate,6,2)),Mid(nowdate,9,2))
	yyyy1 = Left(nowdate,4)
	mm1 = Mid(nowdate,6,2)
	dd1 = Mid(nowdate,9,2)
	yyyy2 = Left(nowdate,4)
	mm2 = Mid(nowdate,6,2)
	dd2 = Mid(nowdate,9,2)
end If
%>

<script type="text/javascript">
function conwrite(idx){
//	var conwrite = window.open('/admin/hitchhiker/mainbanner/hitchhiker_mainbanner_write.asp?idx='+idx,'hitchhiker_mainbanner_write','width=800,height=768,scrollbars=yes,resizable=yes');
	var conwrite = window.open('/academy/mobile/main_banner/academy_mobile_mainbanner_write.asp?idx='+idx,'hitchhiker_mainbanner_write','width=800,height=768,scrollbars=yes,resizable=yes');
	conwrite.focus();
}
function searchFrm(p){
	frm.iC.value = p;
	frm.submit();
}

//이미지 확대화면 새창으로 보여주기
function jsImgView(sImgUrl){
	var wImgView;
	wImgView = window.open('/admin/itemmaster/colortrend_detailImg.asp?sUrl='+sImgUrl,'pImg','width=100,height=100');
	wImgView.focus();
}
</script>
<% '검색---------------------------------------------------------------------------------------------------------- %>
<form name="frm" action="index.asp" method="get">
<input type="hidden" name="iC" value="1">
<input type="hidden" name="research" value="on">
<input type="hidden" name="menupos" value="<%= menupos %>" >
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%=admincolor("tablebg")%>">
	<tr align="center" bgcolor="#FFFFFF">
		<td lowsapn="2" width="100" bgcolor="<%=admincolor("gray")%>"> <b>검색조건</b> </td>
		<td align="left">
			기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %> <input type="checkbox" name="datesearch" value="Y"<% If datesearch="Y" Then Response.write " checked"%>>기간적용&nbsp;
			<select name="SearchGubun">
				<option value ="" style="color:blue">구 분</option>
				<option value="1" <% If "1" = cstr(SearchGubun) Then%> selected <%End if%>>강좌링크</option>
				<option value="2" <% If "2" = cstr(SearchGubun) Then%> selected <%End if%>>상품링크</option>
				<option value="3" <% If "3" = cstr(SearchGubun) Then%> selected <%End if%>>매거진링크</option>
				<option value="4" <% If "4" = cstr(SearchGubun) Then%> selected <%End if%>>강사/작가 링크</option>
				<option value="4" <% If "5" = cstr(SearchGubun) Then%> selected <%End if%>>기타 링크</option>
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
		<td colspan="9" > <% '셀합병(colspan)7개 %>
			검색결과 : <b><%= iTotCnt %></b>
		</td>
	</tr>

	<tr align="center" bgcolor="<%=adminColor("tabletop")%>" height="30">
		<td width="5%"><b>번호</b></td>
		<td width="5%"><b>구분</b></td>
		<td width="10%"><b>이미지</b></td>
		<td width="5%"><b>사용여부</b></td>
		<td width="5%"><b>우선순위</b></td>
		<td width="5%"><b>상태</b></td>
		<td width="10%"><b>시작일</b></td>
		<td width="10%"><b>종료일</b></td>
		<td width="10%"><b>등록자</b></td>
	</tr>
	
	<% if opart.FResultCount > 0 then %>
	
		<% for i = 0 to opart.FResultCount - 1 %>
		<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background='#FFFFFF'; height="30"> 
			<td style="cursor:hand"  onclick="conwrite('<%= opart.FItemList(i).Fidx %>');"><%= opart.FItemList(i).Fidx %></td> <% '번호(idx) %>
			
			<td><%= getAcademyMobileMainBannerGubun(opart.FItemList(i).FReqgubun) %></td> <% '구분(강좌,상품,매거진,강사/작가,기타 %>
			
			<td><img src="<%= opart.FItemList(i).FReqcon_viewthumbimg %>" onclick="jsImgView('<%=opart.FItemList(i).FReqcon_viewthumbimg %>');" width="100" height="100"></td> <% '썸네일 %>
	
			<td><%= opart.FItemList(i).FReqIsusing %></td> <% '사용여부 %>
			
			<td><%= opart.FItemList(i).FReqsortnum %></td> <% '우선순위 %>
			<td>
				<% 
					if now() >=  opart.FItemList(i).FReqSdate AND NOW() <= opart.FItemList(i).FReqEdate then
						Response.write " <span style=""color:blue"">오픈</span>"
					elseif now() < opart.FItemList(i).FReqSdate then
						Response.write " <span style=""color:green"">오픈예정</span>"
					else
						Response.write " <span style=""color:red"">종료</span>"
					end if
					Response.Write "<br />"
				%>
			</td>
			<td> <% '시작일,종료일 %>
				<% 
					Response.Write replace(left(opart.FItemList(i).FReqSdate,10),"-",".") & " / " & Num2Str(hour(opart.FItemList(i).FReqSdate),2,"0","R") & ":" &Num2Str(minute(opart.FItemList(i).FReqSdate),2,"0","R")
				%>
			</td>
			<td><%= replace(left(opart.FItemList(i).FReqEdate,10),"-",".") & " / " & Num2Str(hour(opart.FItemList(i).FReqEdate),2,"0","R") & ":" & Num2Str(minute(opart.FItemList(i).FReqEdate),2,"0","R") %></td> <% '등록일 %>
			<td><%= opart.FItemList(i).FReqmakerid %></td> <% '등록일 %>
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
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
