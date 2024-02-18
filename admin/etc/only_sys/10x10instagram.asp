 <%@ language=vbscript %>
<% option explicit %>
<%
'#############################################################
'	Description : 인스타그램 이벤트용 수동 등록페이지
'	History		: 2016.06.23 유태욱 생성
'#############################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/only_sys/instagrameventCls.asp"-->

<%
Dim i
Dim FResultCount, iCurrentpage, iTotCnt , eventid
Dim Searchgubun, SearchTitle, SearchUsing, SearchEvtCode
	SearchUsing = request("SearchUsing")
	eventid = request("eventid")

	Response.write "현재 이벤트 코드 : "& eventid

	iCurrentpage = NullFillWith(requestCheckVar(Request("IC"),10),1)
if iCurrentpage="" then iCurrentpage=1
	
Dim oinsta
set oinsta = new CInstagramevent
	oinsta.FCurrPage = iCurrentpage
	oinsta.FPageSize = 15
	oinsta.FrectIsusing = SearchUsing
	oinsta.Feventid = eventid
	oinsta.fnGetInstagrameventList
iTotCnt = oinsta.FTotalCount
%>

<script type="text/javascript">
function conwrite(contentsidx,md){
	var conwrite = window.open('/admin/etc/only_sys/instagramevent_write.asp?mode='+md+'&contentsidx='+contentsidx,'instagramevent_write','width=800,height=768,scrollbars=yes,resizable=yes');
	conwrite.focus();
}

function searchFrm(p){
	frm.iC.value = p;
	frm.submit();
}

</script>


<form name="frm" action="10x10instagram.asp" method="get">
<input type="hidden" name="iC" value="1">
<input type="hidden" name="menupos" value="<%'= menupos %>" >
<!--검색-----------------------------------------------------------------------------------------------
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%'=admincolor("tablebg")%>">
	<tr align="center" bgcolor="#FFFFFF">
		<td lowsapn="2" with="50" bgcolor="<%'=admincolor("gray")%>"> <b>검색조건</b> </td>
		<td align="left">
			<b> 사 용 : </b>
			<select name="SearchUsing">
				<option value ="" style="color:blue">전 체</option>
				<option value="Y" <%' If "Y" = cstr(SearchUsing) Then%> selected <%'End if%>>Y</option>
				<option value="N" <%' If "N" = cstr(SearchUsing) Then%> selected <%'End if%>>N</option>
			</select>
			<input type="button" class="button" value="검색조건Reset" onClick="location.href='about_contents_list.asp?reload=on&menupos=<%'=menupos%>';">
		</td>
		<td lowsapn="2" with="50" bgcolor="<%'=admincolor("gray")%>">
			<input type="button" class="button" value="검색" onclick="searchFrm('');">&nbsp;
		</td>
	</tr>
</table>
검색끝----------------------------------------------------------------------------------------------->
</form>

<br>
<!--신규입력버튼---------------------------------------------------------------------------------------->
<table width="100%" align="center">
	<tr>
		<td align="right"><input type="button" name="newBT" class="button" value="신규입력" onclick="conwrite('<%=eventid%>','NEW');"></td>
	</tr>
</table>
<!--신규입력버튼끝------------------------------------------------------------------------------------->

<!--리스트----------------------------------------------------------------------------------------------->
<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%=adminColor("tablebg")%>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="7" > <!--셀합병(colspan)7개-->
			검색결과 : <b><%= iTotCnt %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%=adminColor("tabletop")%>" height="30">
		<td width="5%"><b>번호</b></td>
		<td width="5%"><b>이벤트코드</b></td>
		<td width="20%"><b>게시자ID</b></td>
		<td width="30%"><b>이미지URL</b></td>
		<td width="30%"><b>게시물링크</b></td>
		<td width="5%"><b>사용여부</b></td>
		<td width="10%"><b>등록일</b></td>
	</tr>

	<% if oinsta.FResultCount > 0 then %>
	
		<% for i = 0 to oinsta.FResultCount - 1 %>
		<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background='#FFFFFF'; height="30"> 
			<td onclick="conwrite('<%= oinsta.FItemList(i).Fcontentsidx %>','EDIT');"><%= oinsta.FItemList(i).Fcontentsidx %></td><!--번호(idx)-->
			
			<td><%= oinsta.FItemList(i).Fevt_code %></td><!--이벤트코드-->
	
			<td><%= oinsta.FItemList(i).Fuserid %></td><!--게시자ID-->
	
			<td><a href="<%=oinsta.FItemList(i).Fimgurl %>" target="_blank"><img src="<%= oinsta.FItemList(i).Fimgurl %>" width="50" height="50"  border=0></a></td><!--이미지URL-->
	
			<td><%= oinsta.FItemList(i).Flinkurl %></td><!--게시물링크-->
			
			<td onclick="conwrite('<%= oinsta.FItemList(i).Fcontentsidx %>','EDIT');"><%= oinsta.FItemList(i).Fisusing %></td><!--사용여부-->
			
			<td><% Response.Write left(oinsta.FItemList(i).FRegdate,22) %></td><!--등록일-->
		</tr>
		<% next %>
		<!--페이징처리------------------------------------------>
		<tr height="25" bgcolor="FFFFFF">
			<td colspan="15" align="center">
		       	<% if oinsta.HasPreScroll then %>
					<span class="list_link"><a href="javascript:searchFrm('<%= oinsta.StartScrollPage-1 %>')">[pre]</a></span> '&menupos=<%=menupos%>
				<% else %>
				[pre]
				<% end if %>
					<% for i = 0 + oinsta.StartScrollPage to oinsta.StartScrollPage + oinsta.FScrollCount - 1 %>
						<% if (i > oinsta.FTotalpage) then Exit for %>
						<% if CStr(i) = CStr(iCurrentpage) then %>
						<span class="page_link"><font color="red"><b><%= i %></b></font></span>
						<% else %>
						<a href="javascript:searchFrm('<%= i %>')" class="list_link"><font color="#000000"><%= i %></font></a>
						<% end if %>
					<% next %>
				<% if oinsta.HasNextScroll then %>
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
set oinsta = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->