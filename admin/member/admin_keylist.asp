<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 어드민 USB 인증
' History : 2008.06.30 한용민 생성 
'           2008.07.24 허진원 수정- 분실여부→사용여부
'           2008.09.25 허진원 수정- Key Int→char 변경
'			2009.02.02 허진원 수정- 필드 정렬기능 추가
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/admin_keyclass.asp" -->
<%
Dim oip , i, page , key_idx , oip_edit , del_isusing , idx, strSort
Dim scTnm, scUnm, scUse
	del_isusing = request("del_isusing")
	menupos = request("menupos")
	page = request("page")
	if page = "" then page = 1
	key_idx = request("key_idx")
	idx = request("idx")
	strSort = request("sort")
	if strSort="" then strSort="no"
	scTnm = request("tnm")
	scUnm = request("unm")
	scUse = request("duse")
	
set oip_edit = new ckey_list
	oip_edit.frectkey_idx = key_idx
	oip_edit.frectidx = idx	
	if idx <> "" then
		oip_edit.getkey_edit()
	end if
	
set oip = new ckey_list
	oip.FPageSize	= 1000
	oip.FCurrPage	= page
	oip.FrectSort	= strSort
	oip.FrectTnm	= scTnm
	oip.FrectUnm	= scUnm
	oip.FrectUse	= scUse
	oip.getkey_list()
%>

<script language="javascript">

	function viewplay(idx){
		frm.idx.value = idx;
		frm.submit();
	}
	
	function getsubmit(){
		frm_edit.mode.value = 'edit';	
		frm_edit.submit();
	}
	
	function new_submit(){	
		var new_submit;
		new_submit = window.open("/admin/member/admin_keynew.asp", "new_submit","width=1024,height=200,scrollbars=yes,resizable=yes");
		new_submit.focus();
	}

	function sortList(srt) {
		frm.sort.value = srt;
		frm.submit();
	}
</script>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm_edit" action="/admin/member/admin_keyprocess.asp" method="get">
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="mode">
	<% if oip_edit.Ftotalcount>0 then %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td align="center">번호</td>		
			<td align="center">인증KEY</td>
			<td align="center">Team</td>	
			<td align="center">사용자</td>	
			<td align="center">상세사용자</td>		
			<td align="center">사용여부</td>
			<td align="center">비고</td>
	    </tr>
	    <tr align="center" bgcolor="#FFFFFF">
				<td align="center">
					<input type="hidden" size=30 name="idx" value="<%= oip_edit.FOneItem.fidx %>">
					<%= oip_edit.FOneItem.fidx %>
				</td>
				<td align="center">
					<input type="text" size=30 name="key_idx" value="<%= oip_edit.FOneItem.fkey_idx %>">
				</td>
				<td align="center">
					<select name="teamname" value="<%= oip_edit.FOneItem.fteamname %>">
						<option value="" <% if oip_edit.FOneItem.fteamname = "" then response.write " selected" %>>선택</option>
						<option value="CEO" <% if oip_edit.FOneItem.fteamname = "CEO" then response.write " selected" %>>CEO</option>
						<option value="SYSTEM" <% if oip_edit.FOneItem.fteamname = "SYSTEM" then response.write " selected" %>>SYSTEM</option>
						<option value="ONLINE" <% if oip_edit.FOneItem.fteamname = "ONLINE" then response.write " selected" %>>ONLINE</option>
						<option value="MARKETING" <% if oip_edit.FOneItem.fteamname = "MARKETING" then response.write " selected" %>>MARKETING</option>
						<option value="MD" <% if oip_edit.FOneItem.fteamname = "MD" then response.write " selected" %>>MD</option>
						<option value="WD" <% if oip_edit.FOneItem.fteamname = "WD" then response.write " selected" %>>WD</option>
						<option value="물류" <% if oip_edit.FOneItem.fteamname = "물류" then response.write " selected" %>>물류</option>
						<option value="OFFLINE" <% if oip_edit.FOneItem.fteamname = "OFFLINE" then response.write " selected" %>>OFFLINE</option>
						<option value="CS" <% if oip_edit.FOneItem.fteamname = "CS" then response.write " selected" %>>CS</option>
						<option value="ITHINKSO" <% if oip_edit.FOneItem.fteamname = "ITHINKSO" then response.write " selected" %>>ITHINKSO</option>														
						<option value="경영" <% if oip_edit.FOneItem.fteamname = "경영" then response.write " selected" %>>경영</option>
						<option value="FINGERS" <% if oip_edit.FOneItem.fteamname = "FINGERS" then response.write " selected" %>>FINGERS</option>
						<option value="패션" <% if oip_edit.FOneItem.fteamname = "패션" then response.write " selected" %>>패션사업팀</option>
					</select> 					
				</td>
				<td align="center"><input type="text" size=10 name="username" value="<%= oip_edit.FOneItem.fusername %>"></td>		
				<td align="center"><input type="text" size=10 name="username_detail" value="<%= oip_edit.FOneItem.fusername_detail %>"></td>
				<td align="center">
					<select name="del_isusing" value="<%= oip_edit.FOneItem.fdel_isusing %>">
						<option value="Y" <% if oip_edit.FOneItem.fdel_isusing = "Y" then response.write " selected" %>>사용</option>
						<option value="N" <% if oip_edit.FOneItem.fdel_isusing = "N" then response.write " selected" %>>삭제</option>
					</select>
				</td>	 
				<td align="center"><input type="button" class="button" value="수정" onclick="getsubmit();"></td>
	    </tr>   
	<% else %>
	    <tr align="center" bgcolor="#FFFFFF">
				<td align="center"><font color="red"><b>하단에 수정하실 인증키를 선택해주세요</b></font></td>
	    </tr>   		    
	<% end if %>
</form>
</table>
<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<form name="frmSearch" method="GET">
<input type="hidden" name="menupos" value="<%=menupos%>">
	<tr>
		<td align="left">
		<input type="button" value="신규등록" class="button" onclick="new_submit();">
		</td>
		<td align="right">
			팀 :
				<select name="tnm" class="select">
					<option value="">전체</option>
					<option value="CEO">CEO</option>
					<option value="SYSTEM">SYSTEM</option>
					<option value="ONLINE">ONLINE</option>
					<option value="MARKETING">MARKETING</option>
					<option value="MD">MD</option>
					<option value="WD">WD</option>
					<option value="물류">물류</option>
					<option value="OFFLINE">OFFLINE</option>
					<option value="CS">CS</option>
					<option value="ITHINKSO">ITHINKSO</option>
					<option value="경영">경영</option>
					<option value="FINGERS">FINGERS</option>
					<option value="패션">패션사업팀</option>
				</select> /
			사용자 : <input type="text" name="unm" value="<%=scUnm%>" class="input" size="10"> /
			사용여부 :
				<select name="duse" class="select">
					<option value="">전체</option>
					<option value="Y">사용</option>
					<option value="N">삭제</option>
				</select> &nbsp;
			<input type="submit" value="검색" class="button">
			<script language="javascript">
			document.frmSearch.tnm.value="<%=scTnm%>";
			document.frmSearch.duse.value="<%=scUse%>";
			</script>
		</td>
	</tr>
</form>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" action="" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="key_idx" value="<%=key_idx%>">	
	<input type="hidden" name="idx" value="<%=idx%>">
	<input type="hidden" name="sort" value="<%=strSort%>">
	<% if oip.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oip.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= oip.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center"><% if strSort="no" then %><b>번호</b><% else %><a href="javascript:sortList('no')">번호</a><% end if %></td>
		<td align="center"><% if strSort="key" then %><b>인증KEY</b><% else %><a href="javascript:sortList('key')">인증KEY</a><% end if %></td>
		<td align="center"><% if strSort="team" then %><b>Team</b><% else %><a href="javascript:sortList('team')">Team</a><% end if %></td>	
		<td align="center"><% if strSort="name" then %><b>사용자</b><% else %><a href="javascript:sortList('name')">사용자</a><% end if %></td>	
		<td align="center">상세사용자</td>		
		<td align="center">사용여부</td>		
    </tr>
    
	<% for i=0 to oip.FresultCount-1 %>
	    <tr align="center" bgcolor="<% if oip.FItemList(i).fdel_isusing="Y" then Response.WRite "#FFFFFF": else Response.Write "#E0E0E0": end if %>" onmouseover=this.style.background="<% if oip.FItemList(i).fdel_isusing="Y" then Response.WRite "#F8C880": else Response.Write "#E8B870": end if %>"; onmouseout=this.style.background='<% if oip.FItemList(i).fdel_isusing="Y" then Response.WRite "#FFFFFF": else Response.Write "#E0E0E0": end if %>';>
				<td align="center"><a href="javascript:viewplay('<%= oip.FItemList(i).fidx %>');"><%= oip.FItemList(i).fidx %></a></td>
				<td align="center"><a href="javascript:viewplay('<%= oip.FItemList(i).fidx %>');"><%= oip.FItemList(i).fkey_idx %></a></td>
				<td align="center"><a href="javascript:viewplay('<%= oip.FItemList(i).fidx %>');"><%= oip.FItemList(i).fteamname %></a></td>
				<td align="center"><a href="javascript:viewplay('<%= oip.FItemList(i).fidx %>');"><%= oip.FItemList(i).fusername %></a></td>		
				<td align="center"><a href="javascript:viewplay('<%= oip.FItemList(i).fidx %>');"><%= oip.FItemList(i).fusername_detail %></a></td>
				<td align="center"><a href="javascript:viewplay('<%= oip.FItemList(i).fidx %>');"><%= oip.FItemList(i).fdel_isusing %></a></td>
	    </tr>   
	<% next %>
	
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if oip.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= oip.StartScrollPage-1 %>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oip.StartScrollPage to oip.StartScrollPage + oip.FScrollCount - 1 %>
				<% if (i > oip.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oip.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oip.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
