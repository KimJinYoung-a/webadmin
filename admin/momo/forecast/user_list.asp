<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성예보
' Hieditor : 2010.11.19 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim oforecast,i,page , cardidx , isusing , yyyy , mm , owinner , winner0 , winner1 , winner2 , contents0, contents1, contents2
	menupos = request("menupos")
	yyyy = request("yyyy1")
	mm = request("mm1")
	
	if yyyy = "" then yyyy = year(date())
	if mm = "" then mm = month(date())	

'// 리스트
set oforecast = new cforecast_list
	oforecast.frectyyyymm = yyyy & "-" & mm
	oforecast.fuser_list()

'// 리스트
set owinner = new cforecast_list
	owinner.frectyyyymm = yyyy & "-" & mm
	owinner.frectgubun = "0"
	owinner.fuser_winner()

	if owinner.ftotalcount > 0 then	
		for i = 0 to owinner.ftotalcount - 1
			if owinner.FItemList(i).forderno = "0" then
				winner0 = owinner.FItemList(i).fuserid
				contents0 = owinner.FItemList(i).fcontents
			elseif owinner.FItemList(i).forderno = "1" then
				winner1 = owinner.FItemList(i).fuserid
				contents1 = owinner.FItemList(i).fcontents
			elseif owinner.FItemList(i).forderno = "2" then
				winner2 = owinner.FItemList(i).fuserid
				contents2 = owinner.FItemList(i).fcontents												
			end if
		next
	end if
%>

<script language="javascript">

	//당첨자 등록
	function winnerreg(){
		if (winnerfrm.winner0.value==''){
			alert('1위 아이디를 입력해주세요');
			winnerfrm.winner0.focus();
			return;
		}
		if (winnerfrm.contents0.value==''){
			alert('1위 참여율을 입력해주세요');
			winnerfrm.contents0.focus();
			return;
		}
		if (winnerfrm.winner1.value==''){
			alert('2위 아이디를 입력해주세요');
			winnerfrm.winner2.focus();
			return;
		}
		if (winnerfrm.contents1.value==''){
			alert('2위 참여율을 입력해주세요');
			winnerfrm.contents2.focus();
			return;
		}
		if (winnerfrm.winner2.value==''){
			alert('3위 아이디를 입력해주세요');
			winnerfrm.winner3.focus();
			return;
		}
		if (winnerfrm.contents2.value==''){
			alert('3위 참여율을 입력해주세요');
			winnerfrm.contents3.focus();
			return;
		}												
		winnerfrm.submit()
	}
	
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=get action="">
<input type="hidden" name="cardidx">	
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>		
	<td align="left">
		날짜 : <% DrawYMBox yyyy , mm %> 	
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
<form name="winnerfrm" action="user_process.asp" method="post">
<input type="hidden" name="mode" value="winneredit">
<input type="hidden" name="yyyy" value="<%=yyyy%>">
<input type="hidden" name="mm" value="<%=mm%>">
<tr>
	<td align="left">				
		※<%=yyyy%>년 <%=MM%>월 당첨자<Br>
		1위 : 아이디<input type="text" name="winner0" value="<%=winner0%>"> &nbsp;&nbsp;&nbsp;참여율<input type="text" name="contents0" value="<%=contents0%>"><br>
		2위 : 아이디<input type="text" name="winner1" value="<%=winner1%>"> &nbsp;&nbsp;&nbsp;참여율<input type="text" name="contents1" value="<%=contents1%>"><br>
		3위 : 아이디<input type="text" name="winner2" value="<%=winner2%>"> &nbsp;&nbsp;&nbsp;참여율<input type="text" name="contents2" value="<%=contents2%>">
		<input type="button" onclick="winnerreg(<%=yyyy%>-<%=mm%>);" class="button" value="저장"><br>		
	</td>
	<td align="right">
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->
<br>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oforecast.FTotalCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oforecast.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>	
	<td align="center">고객</td>
	<td align="center">참여일수</td>	
	<td align="center">비고</td>
</tr>
<% for i=0 to oforecast.ftotalcount -1 %>			

<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
	<td align="center">
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
	</td>
	<td align="center">
		<%= oforecast.FItemList(i).fuserid %>
	</td>
	<td align="center">
		<%= oforecast.FItemList(i).fusercount %> [<%=fix((oforecast.FItemList(i).fusercount / datediff("d",DateSerial(yyyy, mm,1),DateSerial(yyyy, mm+1,1))) * 100)%> %]		
	</td>
	<td align="center">
	</td>			
</tr>   

<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if oforecast.HasPreScroll then %>
			<span class="list_link"><a href="?page=<%= oforecast.StartScrollPage-1 %>&isusing=<%=isusing%>">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + oforecast.StartScrollPage to oforecast.StartScrollPage + oforecast.FScrollCount - 1 %>
			<% if (i > oforecast.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oforecast.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="?page=<%= i %>&isusing=<%=isusing%>>" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if oforecast.HasNextScroll then %>
			<span class="list_link"><a href="?page=<%= i %>&isusing=<%=isusing%>">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
</table>

<%
	set oforecast = nothing
	set owinner = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->