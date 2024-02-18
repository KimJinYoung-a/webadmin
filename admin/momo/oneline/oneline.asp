<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 한줄낙서
' Hieditor : 2010.11.23 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim ooneline,i,page , onelineid , isusing
	menupos = request("menupos")
	page = request("page")	
	onelineid = request("onelineidsearch")	
	isusing = request("isusing")			
	if page = "" then page = 1

'// 리스트
set ooneline = new coneline_list
	ooneline.FPageSize = 20
	ooneline.FCurrPage = page
	ooneline.frectonelineid = onelineid	
	ooneline.frectisusing = isusing			
	ooneline.foneline_list()
%>

<script language="javascript">

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

//신규등록 & 수정
function reg(onelineid){
	var reg = window.open('oneline_reg.asp?onelineid='+onelineid,'reg','width=800,height=768,scrollbars=yes,resizable=yes');
	reg.focus();
}

//참여자 보기
function onelinecomment(onelineid){
	var onelinecomment = window.open('oneline_comment_list.asp?onelineid='+onelineid,'onelinecomment','width=1024,height=768,scrollbars=yes,resizable=yes');
	onelinecomment.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get action="">
	<input type="hidden" name="onelineid">	
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>		
		<td align="left">
			번호 : <input type="text" name="onelineidsearch" value="<%=onelineid%>" size=10>			
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
	<% if ooneline.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= ooneline.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= ooneline.FTotalPage %></b>
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
	<% for i=0 to ooneline.FresultCount-1 %>
	<form action="" name="frmBuyPrc<%=i%>" method="get">			
	
    <% if ooneline.FItemList(i).fisusing = "Y" then %>
    <tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
    <% else %>    
    <tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="orange"; onmouseout=this.style.background='FFFFaa';>
	<% end if %>
		<td align="center">
			<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
		</td>
		<td align="center">
			<%= ooneline.FItemList(i).fonelineid %><input type="hidden" name="onelineid" value="<%= ooneline.FItemList(i).fonelineid %>">
		</td>
		<td align="center">
			<%= statsgubun(ooneline.FItemList(i).fstats) %>
		</td>
		<td align="center">
			<%= formatdate(ooneline.FItemList(i).fstartdate,"0000.00.00") %> ~ <%=formatdate(ooneline.FItemList(i).fenddate,"0000.00.00")%>
		</td>
		<td align="center">
			<%= formatdate(ooneline.FItemList(i).fregdate,"0000.00.00") %>
		</td>	
		<td align="center">
			<%= ooneline.FItemList(i).fisusing %>
		</td>				
		<td align="center">
			<input type="button" onclick="reg(<%= ooneline.FItemList(i).fonelineid %>);" class="button" value="수정">
			<input type="button" onclick="onelinecomment(<%= ooneline.FItemList(i).fonelineid %>);" class="button" value="참여자보기[<%= ooneline.FItemList(i).fcommentcount %>]">
		</td>			
    </tr>   
	</form>
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="20" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
	       	<% if ooneline.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= ooneline.StartScrollPage-1 %>&isusing=<%=isusing%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + ooneline.StartScrollPage to ooneline.StartScrollPage + ooneline.FScrollCount - 1 %>
				<% if (i > ooneline.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(ooneline.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&isusing=<%=isusing%>>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if ooneline.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&isusing=<%=isusing%>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<%
	set ooneline = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->