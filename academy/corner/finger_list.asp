<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 소문난전시
' History : 2009.09.14 한용민 생성
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
Dim oip,i,page , rumour_id , isusing, vGubun, vParam
	menupos = RequestCheckvar(request("menupos"),10)
	page = RequestCheckvar(request("page"),10)
	rumour_id = requestcheckvar(request("rumour_id"),4)
	isusing = requestcheckvar(request("isusing"),1)
	vGubun = requestcheckvar(request("gubun"),1)
		
	if page = "" then page = 1
		
	vParam = "&menupos="&menupos&"&gubun="&vGubun&"&isusing="&isusing&"&rumour_id="&rumour_id&""
				
'//리스트
set oip = new crumour_one_list
	oip.FPageSize = 20
	oip.FCurrPage = page
	oip.frectisusing = isusing
	oip.frectgubun = vGubun
	oip.frectrumour_id = rumour_id
	oip.frumour_list()
%>

<script language="javascript">

document.domain = "10x10.co.kr";

// 강사등록&수정
function reg_rumour(rumour_id){
	var reg_rumour = window.open('/academy/corner/finger_reg.asp?rumour_id='+rumour_id,'reg_rumour','width=800,height=768,scrollbars=yes,resizable=yes');
	reg_rumour.focus();
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
		<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<select name="gubun">
				<option value="">구분</option>
				<option value="r" <% if vGubun = "r" then response.write " selected" %>>소문난전시</option>
				<option value="l" <% if vGubun = "l" then response.write " selected" %>>생활레시피</option>
				<option value="f" <% if vGubun = "f" then response.write " selected" %>>핑거스토리</option>
			</select>
			&nbsp;&nbsp;&nbsp;
			<select name="isusing">
				<option value="">사용여부</option>
				<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
			</select>
			&nbsp;&nbsp;&nbsp;
			ID: <input type="text" name="rumour_id" value="<%=rumour_id%>">
		</td>	
		<td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->
		
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			※구분 : <b>소(소문난 전시)</b>, <b>생(생활레시피)</b>, <b>핑(핑거스토리)</b>
		</td>
		<td align="right">				
			<input type="button" class="button" value="전시등록" onclick="reg_rumour('');">				
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
		<td align="center" >구분</td>
		<td align="center">제목</td>	
		<td align="center">전시자</td>	
		<td align="center">기간</td>		
		<td align="center">코맨트</td>
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
		<td align="center">
			<a href="http://<% If application("Svr_Info") = "Dev" THEN %>test.<% Else %>www.<% End If %>thefingers.co.kr/corner/finger_story.asp?idx=<%= oip.FItemList(i).fidx %>" target="_blank"><img src="<%= oip.FItemList(i).flist_image %>" width=40 height=40 border="0"></a>
		</td>
		<td align="center"><%= oip.FItemList(i).fidx %></td>
		<td align="center">
			<%
				If oip.FItemList(i).fgubun = "r" Then
					Response.Write "소"
				ElseIf oip.FItemList(i).fgubun = "l" Then
					Response.Write "생"
				ElseIf oip.FItemList(i).fgubun = "f" Then
					Response.Write "핑"
				End If
			%>
		</td>
		<td align="center"><%= oip.FItemList(i).ftitle %></td>
		<td align="center"><%= oip.FItemList(i).fuserid %></td>
		
		<td align="center">
			<%
				If oip.FItemList(i).fgubun = "r" Then
					Response.Write FormatDate(oip.FItemList(i).fstartdate,"0000.00.00") & "-" & FormatDate(oip.FItemList(i).fenddate,"0000.00.00")
				End IF
			%>
		</td>
		<td align="center"><%= oip.FItemList(i).fcommentyn %></td>
		<td align="center"><%= oip.FItemList(i).fisusing %></td>
		<td align="center">
			<input type="button" class="button" value="수정" onclick="reg_rumour('<%= oip.FItemList(i).fidx %>');">			
		</td>			
    </tr>   
	</form>
	<% next %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if oip.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= oip.StartScrollPage-1 %><%=vParam%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oip.StartScrollPage to oip.StartScrollPage + oip.FScrollCount - 1 %>
				<% if (i > oip.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oip.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %><%=vParam%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oip.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %><%=vParam%>">[next]</a></span>
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
