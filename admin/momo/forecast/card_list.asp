<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성예보
' Hieditor : 2010.11.15 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim oforecast,i,page , cardidx , isusing
	menupos = request("menupos")
	page = request("page")
	cardidx = request("cardidxsearch")	
	isusing = request("isusing")			
	if page = "" then page = 1

'// 리스트
set oforecast = new cforecast_list
	oforecast.FPageSize = 20
	oforecast.FCurrPage = page
	oforecast.frectcardidx = cardidx	
	oforecast.frectisusing = isusing			
	oforecast.fcard_list()
%>

<script language="javascript">

	//신규등록 & 수정
	function reg(cardidx){
		var reg = window.open('/admin/momo/forecast/card_reg.asp?cardidx='+cardidx,'reg','width=600,height=400,scrollbars=yes,resizable=yes');
		reg.focus();
	}
	
	//투표등록
	function card_reg(cardidx){
		var card_reg = window.open('/admin/momo/forecast/card_detail.asp?cardidx='+cardidx,'card_reg','width=600,height=768,scrollbars=yes,resizable=yes');
		card_reg.focus();
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
		cardidx : <input type="text" name="cardidxsearch" value="<%=cardidx%>" size=10>		
		&nbsp; 사용여부 : 
		<select name="isusing">
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
<% if oforecast.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oforecast.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oforecast.FTotalPage %></b>
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
<% for i=0 to oforecast.FresultCount-1 %>			

<% if oforecast.FItemList(i).fisusing = "Y" then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
<% else %>    
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="orange"; onmouseout=this.style.background='FFFFaa';>
<% end if %>
	<td align="center">
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
	</td>
	<td align="center">
		<%= oforecast.FItemList(i).fcardidx %><input type="hidden" name="cardidx" value="<%= oforecast.FItemList(i).fcardidx %>">
	</td>
	<td align="center">
		<%= statsgubun(oforecast.FItemList(i).fstats) %>
	</td>	
	<td align="center">
		<%= formatdate(oforecast.FItemList(i).fstartdate,"0000.00.00") %> ~ <%=formatdate(oforecast.FItemList(i).fenddate,"0000.00.00")%>
	</td>			
	<td align="center">
		<%= formatdate(oforecast.FItemList(i).fregdate,"0000.00.00") %>
	</td>		
	<td align="center">
		<%= oforecast.FItemList(i).fisusing %>
	</td>			
	<td align="center">
		<input type="button" onclick="reg(<%= oforecast.FItemList(i).fcardidx %>);" class="button" value="수정">
		<input type="button" onclick="card_reg(<%= oforecast.FItemList(i).fcardidx %>);" class="button" value="카드등록(<%= oforecast.FItemList(i).fcardcount %>)">
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
			<a href="?page=<%= i %>&isusing=<%=isusing%>" class="list_link"><font color="#000000"><%= i %></font></a>
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
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->