<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성놀이
' Hieditor : 2010.12.13 허진원 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->
<%
Dim oplay,i,page , playSn , isusing
	menupos = request("menupos")
	page = request("page")
	playSn = request("playSnSearch")
	isusing = request("isusing")
	if page = "" then page = 1

'// 리스트
set oplay = new cPlayList
	oplay.FPageSize = 20
	oplay.FCurrPage = page
	oplay.frectplaySn = playSn
	oplay.frectisusing = isusing
	oplay.fplay_list()
%>

<script language="javascript">

	//신규등록 & 수정
	function reg(playSn){
		var reg = window.open('/admin/momo/play/play_reg.asp?playSn='+playSn,'reg','width=600,height=400,scrollbars=yes,resizable=yes');
		reg.focus();
	}
	
	//상품등록
	function item_reg(playSn){
		var item_reg = window.open('/admin/momo/play/play_item.asp?playSn='+playSn,'item_reg','width=800,height=768,scrollbars=yes,resizable=yes');
		item_reg.focus();
	}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=get action="">
<input type="hidden" name="playSn">	
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		playSn : <input type="text" name="playSnSearch" value="<%=playSn%>" size=10>
		&nbsp; 사용여부 : 
		<select name="isusing" value="<%=isusing%>">
			<option value="" <% if isusing = "" then response.write " selected" %>>사용여부</option>
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
</form>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
		</td>
		<td align="right">
			<input type="button" onclick="reg('');" value="신규등록" class="button">
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oplay.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oplay.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oplay.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td align="center">번호</td>
	<td align="center">상태</td>
	<td align="center">기간</td>
	<td align="center">등록일</td>
	<td align="center">사용여부</td>
	<td align="center">비고</td>
</tr>
<% for i=0 to oplay.FresultCount-1 %>

<% if oplay.FItemList(i).fisusing = "Y" then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
<% else %>    
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="orange"; onmouseout=this.style.background='FFFFaa';>
<% end if %>
	<td align="center">
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
	</td>
	<td align="center">
		<%= oplay.FItemList(i).fplaySn %><input type="hidden" name="playSn" value="<%= oplay.FItemList(i).fplaySn %>">
	</td>
	<td align="center">
		<%= statsgubun(oplay.FItemList(i).fstats) %>
	</td>	
	<td align="center">
		<%= formatdate(oplay.FItemList(i).fstartdate,"0000.00.00") %> ~ <%=formatdate(oplay.FItemList(i).fenddate,"0000.00.00")%>
	</td>
	<td align="center">
		<%= formatdate(oplay.FItemList(i).fregdate,"0000.00.00") %>
	</td>		
	<td align="center">
		<%= oplay.FItemList(i).fisusing %>
	</td>			
	<td align="center">
		<input type="button" onclick="reg(<%= oplay.FItemList(i).fplaySn %>);" class="button" value="수정">
		<input type="button" onclick="item_reg(<%= oplay.FItemList(i).fplaySn %>);" class="button" value="상품등록(<%= oplay.FItemList(i).fitemcount %>)">
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
       	<% if oplay.HasPreScroll then %>
			<span class="list_link"><a href="?page=<%= oplay.StartScrollPage-1 %>&isusing=<%=isusing%>">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + oplay.StartScrollPage to oplay.StartScrollPage + oplay.FScrollCount - 1 %>
			<% if (i > oplay.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oplay.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="?page=<%= i %>&isusing=<%=isusing%>>" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if oplay.HasNextScroll then %>
			<span class="list_link"><a href="?page=<%= i %>&isusing=<%=isusing%>">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
</table>

<%
	set oplay = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->