<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 삽별구역설정
' Hieditor : 2010.01.11 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/zone/zone_cls.asp"-->
<%
Dim ozone_list, ozone_detail, i, page, zonegroup ,zonegroup_name ,zonegroup_type
dim isusing, regdate ,menupos
	zonegroup = requestCheckVar(request("zonegroup"),10)
	menupos = requestCheckVar(request("menupos"),10)
	page = requestCheckVar(request("page"),10)

if page = "" then page = 1

set ozone_list = new czone_list
	ozone_list.FPageSize = 20
	ozone_list.FCurrPage = page
	ozone_list.Getoffshopzonecommon_list()

set ozone_detail = new czone_list
	ozone_detail.frectzonegroup = zonegroup
	
	if zonegroup <> "" then		
		ozone_detail.Getoffshopzonecommon_detail()
		
		if ozone_detail.ftotalcount > 0 then
								
			zonegroup_name = ozone_detail.FOneItem.fzonegroup_name
			zonegroup_type = ozone_detail.FOneItem.fzonegroup_type
			isusing = ozone_detail.FOneItem.fisusing
			regdate = ozone_detail.FOneItem.fregdate
		end if
		
	end if
%>

<script language="javascript">
	
	function groupedit(zonegroup){
		location.href="/admin/offshop/zone/zone_common.asp?menupos=<%=menupos%>&zonegroup="+zonegroup;
	}

	function newreg(){
		location.href="/admin/offshop/zone/zone_common.asp?menupos=<%=menupos%>";
	}

	function reg(){
		
		if (frm.zonegroup_name.value=='') {
			alert('그룹명을 입력해 주세요');
			frm.zonegroup_name.focus();
			return;
		}

		if (frm.isusing.value=='') {
			alert('사용여부를 선택해 주세요');
			frm.isusing.focus();			
			return;
		}
		
		frm.action='zone_process.asp';
		frm.mode.value = "zonecommonedit";
		frm.submit();
	}
	
</script>

<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post">
<input type="hidden" name="mode">
<input type="hidden" name="isusing" value="Y">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr bgcolor="#FFFFFF">
	<td align="center">그룹번호<br></td>
	<td>
		<%=zonegroup%><input type="hidden" name="zonegroup" value="<%=zonegroup%>">
	</td>
</tr>	
<tr bgcolor="#FFFFFF">
	<td align="center">그룹명</td>
	<td>
		<input type="text" name="zonegroup_name" value="<%=zonegroup_name%>">
	</td>
</tr>
<!--<tr bgcolor="#FFFFFF">
	<td align="center">사용여부<br></td>
	<td>
		<select name="isusing">
			<option value="Y" <%' if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <%' if isusing = "N" then response.write " selected" %>>N</option>
		</select>
	</td>
</tr>-->
<tr bgcolor="#FFFFFF">
	<td align="center" colspan=2>
		<input type="button" value="저장" class="button" onclick="reg();">
	</td>
</tr>
</form>
</table>
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		※ <font color="red">[중요] </font>매장내 그룹이 변경되면 기존 그룹명을 수정하지 마시고, 새로 등록하세요.
		<br>기존 그룹명을 현재 변경될 그룹명으로 수정후 사용 하실경우,
		<br>기존 그룹으로 등록되어진 상품들이 모두 현재 그룹명으로 변경되는 문제가 발생됩니다
	</td>
	<td align="right">	
		<input type="button" class="button" value="신규등록" onclick="newreg('');">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		검색결과 : <b><%= ozone_list.FTotalCount %></b>
		&nbsp;
		페이지 : <b><%= page %>/ <%= ozone_list.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>그룹번호</td>
	<td>그룹명</td>
	<!--<td>사용여부</td>-->
	<td>비고</td>	
</tr>
<% if ozone_list.FresultCount>0 then %>
<% for i=0 to ozone_list.FresultCount-1 %>
<% if ozone_list.FItemList(i).fisusing = "Y" then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>
<% else %>    
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFaa';>
<% end if %>
	<td align="center">
		<%= ozone_list.FItemList(i).fzonegroup %>
	</td>		
	<td align="center">
		<%= ozone_list.FItemList(i).fzonegroup_name %>
	</td>
	<!--<td align="center">
		<%'= ozone_list.FItemList(i).fisusing %>
	</td>-->
	<td align="center">
		<input type="button" value="수정" class="button" onclick="groupedit('<%= ozone_list.FItemList(i).fzonegroup %>');">
	</td>	
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if ozone_list.HasPreScroll then %>
			<span class="list_link"><a href="?page=<%= ozone_list.StartScrollPage-1 %>">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + ozone_list.StartScrollPage to ozone_list.StartScrollPage + ozone_list.FScrollCount - 1 %>
			<% if (i > ozone_list.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(ozone_list.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if ozone_list.HasNextScroll then %>
			<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
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
set ozone_list = nothing
set ozone_detail = nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->