<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2008.10.28 한용민 생성
'	Description : 오거나이저
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/organizer/organizer_cls.asp"-->

<%
Dim oip , i, page , key_idx , oip_edit , isusing , idx
	isusing = request("isusing")
	menupos = request("menupos")
	page = request("page")
	if page = "" then page = 1
	idx = request("idx")
	
set oip_edit = new organizerCls
	oip_edit.frectidx = idx
	if idx <> "" then
		oip_edit.falpha_value()
	end if
	
set oip = new organizerCls
	oip.FPageSize = 20
	oip.FCurrPage = page
	oip.falpha()
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
		new_submit = window.open("/admin/organizer/alpha_new.asp", "new_submit","width=1024,height=200,scrollbars=yes,resizable=yes");
		new_submit.focus();
	}
	
</script>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm_edit" action="/admin/organizer/alpha_process.asp" method="get">
	<input type="hidden" name="mode">
	<% if oip_edit.Ftotalcount>0 then %>
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center">번호</td>		
		<td align="center">상품코드</td>
		<td align="center">사용여부</td>
		<td align="center">비고</td>
	    </tr>
	    <tr align="center" bgcolor="#FFFFFF">
				<td align="center">
					<input type="hidden" size=30 name="idx" value="<%= oip_edit.FOneItem.fidx %>">
					<%= oip_edit.FOneItem.fidx %>
				</td>
				<td align="center">
					<input type="text" size=30 name="itemid" value="<%= oip_edit.FOneItem.fitemid %>">
				</td>	
				<td align="center">
					<select name="isusing" value="<%= oip_edit.FOneItem.fisusing %>">
						<option value="" <% if oip_edit.FOneItem.fisusing = "" then response.write " selected" %>>선택</option>
						<option value="Y" <% if oip_edit.FOneItem.fisusing = "Y" then response.write " selected" %>>Y</option>
						<option value="N" <% if oip_edit.FOneItem.fisusing = "N" then response.write " selected" %>>N</option>
					</select>
				</td>	 
				<td align="center"><input type="button" class="button" value="수정" onclick="getsubmit();"></td>
	    </tr>   
	<% else %>
	    <tr align="center" bgcolor="#FFFFFF">
				<td align="center"><font color="red"><b>하단에 수정하실 값을 선택해주세요</b></font></td>
	    </tr>   		    
	<% end if %>
</form>
</table>
<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
		<input type="button" value="신규등록" class="button" onclick="new_submit();">
		</td>
		<td align="right">	
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
※ 상단에 있는 3개만 오거나이저 프론트페이지에 노출됩니다.
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" action="" method="get">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="key_idx" value="<%=key_idx%>">	
	<input type="hidden" name="idx" value="<%=idx%>">
	<% if oip.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oip.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= oip.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td align="center">상품코드</td>
		<td align="center">사용여부</td>			
    </tr>
    
	<% for i=0 to oip.FresultCount-1 %>
	    <tr align="center" bgcolor="<% if oip.FItemList(i).fisusing="Y" then Response.WRite "#FFFFFF": else Response.Write "#E0E0E0": end if %>">
				<td align="center"><a href="javascript:viewplay('<%= oip.FItemList(i).fidx %>');"><%= oip.FItemList(i).fitemid %></a></td>		
				<td align="center"><a href="javascript:viewplay('<%= oip.FItemList(i).fidx %>');"><%= oip.FItemList(i).fisusing %></a></td>
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