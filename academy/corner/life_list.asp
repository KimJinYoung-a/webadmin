<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 생활레시피
' History : 2009.09.15 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/corner/corner_cls.asp"-->

<%
Dim oip,i,page , life_id , isusing
	menupos = RequestCheckvar(request("menupos"),10)
	page = RequestCheckvar(request("page"),10)
	life_id = requestcheckvar(request("life_id"),4)
	isusing = requestcheckvar(request("isusing"),1)
		
	if page = "" then page = 1
				
'//리스트
set oip = new clife_list
	oip.FPageSize = 20
	oip.FCurrPage = page
	oip.frectisusing = isusing
	oip.frectlife_id = life_id
	oip.flife_list()
%>

<script language="javascript">

// 강사등록&수정
function reg_life(life_id){
	var reg_life = window.open('/academy/corner/life_reg.asp?life_id='+life_id,'reg_life','width=800,height=768,scrollbars=yes,resizable=yes');
	reg_life.focus();
}

function AnSelectAllFrame(bool){
	var frm;
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.disabled!=true){
				frm.cksel.checked = bool;
				AnCheckClick(frm.cksel);
			}
		}
	}
}	

function AnCheckClick(e){
	if (e.checked)
		hL(e);
	else
		dL(e);
}	

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}

function CheckSelected(){
	var pass=false;
	var frm;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		return false;
	}
	return true;
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<select name="isusing">
				<option value="">사용여부</option>
				<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
			</select>
			&nbsp;ID: <input type="text" name="life_id" value="<%=life_id%>">
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
			<input type="button" class="button" value="신규등록" onclick="reg_life('');">				
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oip.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= oip.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
   		<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td align="center" >이미지</td>
		<td align="center" >ID</td>
		<td align="center">제목</td>				
		<td align="center">comment<br>사용여부</td>
		<td align="center">사용여부</td>			
		<td align="center">비고</td>
    </tr>
	<% 
	if oip.FresultCount>0 then    
	
	for i=0 to oip.FresultCount-1 
	%>
	<form action="" name="frmBuyPrc<%=i%>" method="get">		   
    <% if oip.FItemList(i).fisusing = "Y" then %>
    <tr align="center" bgcolor="#FFFFFF">
    <% else %>    
    <tr align="center" bgcolor="#FFFFaa">
	<% end if %>
		<td align="center">
			<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
		</td>
		<td align="center"><img src="<%= oip.FItemList(i).flist_image %>" width=40 height=40></td>
		<td align="center"><%= oip.FItemList(i).flife_id %></td>
		<td align="center"><%= oip.FItemList(i).flife_title %></td>	
		<td align="center"><%= oip.FItemList(i).fcommentyn %></td>
		<td align="center"><%= oip.FItemList(i).fisusing %></td>		
		<td align="center">
			<input type="button" class="button" value="수정" onclick="reg_life('<%= oip.FItemList(i).flife_id %>');">			
		</td>			
    </tr>   
	</form>
	<% next %>
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
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="10" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
</table>

<%
	set oip = nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
