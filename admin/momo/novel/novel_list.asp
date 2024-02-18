<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 한줄소설 리스트
' Hieditor : 2009.11.17 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim ocontents,i,page , novelid , isusing , title
	menupos = request("menupos")
	page = request("page")
	novelid = request("novelidsearch")
	title = request("title")
	isusing = request("isusing")			
	if page = "" then page = 1

'// 리스트
set ocontents = new cnovel_list
	ocontents.FPageSize = 20
	ocontents.FCurrPage = page
	ocontents.frectnovelid = novelid
	ocontents.frecttitle = title
	ocontents.frectisusing = isusing			
	ocontents.fnovel_list()
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
function reg(novelid){
	var reg = window.open('/admin/momo/novel/novel_reg.asp?novelid='+novelid,'reg','width=600,height=400,scrollbars=yes,resizable=yes');
	reg.focus();
}

//제안글보기
function poposallist(){
	var poposallist = window.open('proposal_list.asp','poposallist','width=1024,height=768,scrollbars=yes,resizable=yes');
	poposallist.focus();
}

//코맨트보기
function regcomment(novelid){
	var regcomment = window.open('/admin/momo/novel/novel_comment_list.asp?novelid='+novelid,'regcomment','width=1024,height=768,scrollbars=yes,resizable=yes');
	regcomment.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get action="">
	<input type="hidden" name="novelid">	
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>		
		<td align="left">
			novelid:<input type="text" name="novelidsearch" value="<%=novelid%>" size=10>
			&nbsp; 주제:<input type="text" name="title" value="<%=title%>" size=20>
			&nbsp; 사용여부:
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
			<input type="button" onclick="poposallist();" value="제안글보기" class="button">
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if ocontents.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= ocontents.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= ocontents.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
   		<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td align="center">novelid</td>
		<td align="center">상태</td>
		<td align="center">소설시작<br>이미지</td>
		<td align="center">기간</td>
		<td align="center">등록일</td>
		<td align="center">주제</td>				
		<td align="center">사용여부</td>
		<td align="center">등록된<br>상품수</td>
		<td align="center">비고</td>
    </tr>
	<% for i=0 to ocontents.FresultCount-1 %>
	<form action="" name="frmBuyPrc<%=i%>" method="get">			
	
    <% if ocontents.FItemList(i).fisusing = "Y" then %>
    <tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
    <% else %>    
    <tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="orange"; onmouseout=this.style.background='ffffff';>
	<% end if %>
		<td align="center">
			<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
		</td>
		<td align="center">
			<%= ocontents.FItemList(i).fnovelid %><input type="hidden" name="novelid" value="<%= ocontents.FItemList(i).fnovelid %>">
		</td>
		<td align="center">
			<%= statsgubun(ocontents.FItemList(i).fstats) %>
		</td>		
		<td align="center">
			<img src="<%=webImgUrl%>/momo/novel/word/<%= ocontents.FItemList(i).fwordimage %>" width=60 height=40>
		</td>		
		<td align="center">
			<%= formatdate(ocontents.FItemList(i).fstartdate,"0000.00.00") %> ~ <%=formatdate(ocontents.FItemList(i).fenddate,"0000.00.00")%>
		</td>			
		<td align="center">
			<%= formatdate(ocontents.FItemList(i).fregdate,"0000.00.00") %>
		</td>	
		<td align="center">
			<%= chrbyte(ocontents.FItemList(i).ftitle,20,"Y") %>
		</td>	
		<td align="center">
			<%= ocontents.FItemList(i).fisusing %>
		</td>	
		<td align="center">
			<a href="javascript:regcomment(<%= ocontents.FItemList(i).fnovelid %>)"><%= ocontents.FItemList(i).fcommentcount %></a>
		</td>					
		<td align="center">
			<input type="button" onclick="reg(<%= ocontents.FItemList(i).fnovelid %>);" class="button" value="수정">
		</td>			
    </tr>   
	</form>
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="10" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if ocontents.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= ocontents.StartScrollPage-1 %>&isusing=<%=isusing%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + ocontents.StartScrollPage to ocontents.StartScrollPage + ocontents.FScrollCount - 1 %>
				<% if (i > ocontents.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(ocontents.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&isusing=<%=isusing%>>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if ocontents.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&isusing=<%=isusing%>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<%
	set ocontents = nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->