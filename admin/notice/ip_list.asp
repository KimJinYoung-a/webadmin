<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 사내 ip 관리
' History : 2008.07.01 한용민 생성 
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/member/10x10staffcls.asp" -->
<%
Dim oip , i, page , company_ip , oip_edit , part_sn , gubuncd , ipidx
	menupos = request("menupos")
	page = request("page")
	if page = "" then page = 1
	company_ip = request("company_ip")
	part_sn = request("part_sn_box")
	gubuncd = request("gubuncd")
	ipidx = request("ipidx")

if gubuncd = "" then gubuncd = "01"
	
set oip_edit = new cip_list
	oip_edit.frectipidx = ipidx

	if ipidx <> "" then
		oip_edit.getip_edit()
		
		if oip_edit.FOneItem.fpart_sn = "" or isnull(oip_edit.FOneItem.fpart_sn) then
			part_sn = "0000"
		else 
			part_sn = oip_edit.FOneItem.fpart_sn
		end if
	end if
	
set oip = new cip_list
	oip.FPageSize = 100
	oip.FCurrPage = page
	oip.frectgubuncd = gubuncd	
	oip.getip_list()
%>

<script language="javascript">

	function viewplay(ipidx){
		frmedit.ipidx.value = ipidx;
		frmedit.submit();
	}

	function pagesubmit(page){
		frmsearch.page.value = page;
		frmsearch.submit();
	}

	function newreg(){
		frmsearch.submit();
	}
	
	function edit(){
		if (frmedit.gubuncd.value == ''){
			alert('지역구분이 없습니다');
			return;
		}

		if (frmedit.company_ip.value == ''){
			alert('ip를 입력하세요');
			return;
		}
				
		frmedit.action = '/admin/notice/ip_process.asp';
		frmedit.submit();
	}
		
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmsearch" method="post" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		지역구분 : <% DrawEquipMentGubun "30" ,"gubuncd" , gubuncd ," onchange=""pagesubmit('');""" %>
	</td>

	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="pagesubmit('');">
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
<% if gubuncd = "01" then %>
-공인ip 대역<br>
61.252.133.02 - 61.252.133.127<br><br>

-사설ip 대역 <br>
192.168.1.1 - 192.168.1.254<br>
192.168.0.1 - 192.168.0.254<br>
<% else %>

<% end if %>

<!-- 리스트 시작 -->
<input type="button" class="button" value="신규등록" onClick="newreg();">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmedit" method="post" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="edit">	
<%
'/수정
if oip_edit.Ftotalcount>0 then
%>
<input type="hidden" size=10 name="gubuncd" value="<%= oip_edit.FOneItem.fgubuncd %>">
<input type="hidden" size=10 name="ipidx" value="<%= oip_edit.FOneItem.fipidx %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">IP</td>
	<td align="center">ID</td>	
	<td align="center">이름</td>	
	<td align="center">파트</td>
	<td align="center">사용여부</td>	
	<td align="center">비고</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center">
		<input type="hidden" size=10 name="company_ip" value="<%= oip_edit.FOneItem.fcompany_ip %>">
		<%= oip_edit.FOneItem.fcompany_ip %>
	</td>
	<td align="center">
		<input type="text" size=10 name="id" value="<%= oip_edit.FOneItem.fid %>">		
	</td>
	<td align="center"><input type="text" size=10 name="company_name" value="<%= oip_edit.FOneItem.fcompany_name %>"></td>		
	<td align="center"><%= printPartOption ("part_sn_box" , part_sn) %></td>
	<td align="center"><% drawSelectBoxUsingYN "isusing" , oip_edit.FOneItem.fisusing %></td>
	<td align="center"><input type="button" class="button" value="수정" onclick="edit();"></td>
</tr>   
<%
'/신규등록
else
%>
<input type="hidden" size=10 name="ipidx">
<input type="hidden" size=10 name="gubuncd" value="<%= gubuncd %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">IP</td>
	<td align="center">ID</td>	
	<td align="center">이름</td>	
	<td align="center">파트</td>	
	<td align="center">비고</td>	
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center">
		<input type="text" size=15 name="company_ip">
	</td>
	<td align="center">
		<input type="text" size=15 name="id">		
	</td>
	<td align="center"><input type="text" size=15 name="company_name"></td>		
	<td align="center"><%= printPartOption ("part_sn_box" , part_sn) %></td>
	<td align="center"><input type="button" class="button" value="신규저장" onclick="edit();"></td>
</tr>
<% end if %>
</form>
</table>
<br>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oip.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= oip.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= oip.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td align="center">IP</td>
	<td align="center">ID</td>	
	<td align="center">이름</td>	
	<td align="center">파트</td>
	<td align="center">사용여부</td>
	<td align="center">비고</td>		
</tr>

<% for i=0 to oip.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background="#ffffff";>
	<td align="center"><%= oip.FItemList(i).fcompany_ip %></td>
	<td align="center"><%= oip.FItemList(i).fid %></td>
	<td align="center"><%= oip.FItemList(i).fcompany_name %></td>		
	<td align="center"><%= oip.FItemList(i).fpart_name %></td>
	<td align="center"><%= oip.FItemList(i).fisusing %></td>
	<td align="center"><input type="button" class="button" value="수정" onclick="viewplay('<%= oip.FItemList(i).fipidx %>');"></td>
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
			<span class="list_link"><a href="javascript:pagesubmit(<%= oip.StartScrollPage-1 %>);">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + oip.StartScrollPage to oip.StartScrollPage + oip.FScrollCount - 1 %>
			<% if (i > oip.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(oip.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:pagesubmit(<%= i %>);" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if oip.HasNextScroll then %>
			<span class="list_link"><a href="javascript:pagesubmit(<%= i %>);">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
</table>

<%
	set oip_edit = nothing
	set oip = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
