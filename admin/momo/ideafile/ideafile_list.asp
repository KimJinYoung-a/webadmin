<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 아이디어파일
' Hieditor : 2009.11.19 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim ocontents,i,page , ideafileid , isusing , cate_large , bestyn
	menupos = request("menupos")
	page = request("page")
	ideafileid = request("ideafileidsearch")
	bestyn = request("bestyn")
	cate_large = requestcheckvar(request("cate_large"),3)	
	isusing = request("isusing")			
	if page = "" then page = 1

'// 리스트
set ocontents = new cideafile_list
	ocontents.FPageSize = 20
	ocontents.FCurrPage = page
	ocontents.frectcate_large = cate_large
	ocontents.frectbestyn = bestyn
	ocontents.frectideafileid = ideafileid	
	ocontents.frectisusing = isusing			
	ocontents.fideafile_list()
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

// 진행중으로 변경 
function changestats(upfrm){
if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}	
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.ideafileid.value = upfrm.ideafileid.value + frm.ideafileid.value + "," ;
						
				}
			}
		}

			var tot;
			tot = upfrm.ideafileid.value;
			upfrm.ideafileid.value = ""
		var changestats;

		changestats = window.open("/admin/momo/ideafile/ideafile_process.asp?ideafileid=" +tot + "&mode=ing" , "changestats","width=400,height=300,scrollbars=yes,resizable=yes");
		changestats.focus();
}

// 삭제 
function delete_ideafileid(upfrm){
if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}	
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.ideafileid.value = upfrm.ideafileid.value + frm.ideafileid.value + "," ;
						
				}
			}
		}

			var tot;
			tot = upfrm.ideafileid.value;
			upfrm.ideafileid.value = ""
		var changestats;

		changestats = window.open("/admin/momo/ideafile/ideafile_process.asp?ideafileid=" +tot + "&mode=delete" , "changestats","width=400,height=300,scrollbars=yes,resizable=yes");
		changestats.focus();
}

	//카테고리 변경
	function change_id(tmp){
		
		frm_search.submit();	
	}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm_search" method=get action="">
	<input type="hidden" name="ideafileid">	
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>		
		<td align="left">
			ideafileid:<input type="text" name="ideafileidsearch" value="<%=ideafileid%>" size=10>			
			&nbsp; 카테고리:<% Drawcate "cate_large" , cate_large %>
			&nbsp; 사용여부:
			<select name="isusing" value="<%=isusing%>">
				<option value="" <% if isusing = "" then response.write " selected" %>>사용여부</option>
				<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
			</select>
			&nbsp; 베스트여부:
			<select name="bestyn" value="<%=bestyn%>">
				<option value="" <% if bestyn = "" then response.write " selected" %>>베스트여부</option>
				<option value="Y" <% if bestyn = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if bestyn = "N" then response.write " selected" %>>N</option>
			</select>						
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="frm_search.submit();">
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
			<input type="button" onclick="changestats(frm_search);" value="베스트선정" class="button">
			<input type="button" onclick="delete_ideafileid(frm_search);" value="노출안함" class="button">
		</td>
		<td align="right">			
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
		<td align="center">ideafileid</td>
		<td align="center">이미지</td>		
		<td align="center">상품코드</td>
		<td align="center">등록일</td>
		<td align="center" >코맨트</td>			
		<td align="center">추천수</td>	
		<td align="center">베스트여부</td>
		<td align="center">사용여부</td>		
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
			<%= ocontents.FItemList(i).fideafileid %><input type="hidden" name="ideafileid" value="<%= ocontents.FItemList(i).fideafileid %>">
		</td>
		<td align="center">
			<img src="<%= ocontents.FItemList(i).fImageList %>" width=50 height=50>
		</td>		
		<td align="center">
			<%= ocontents.FItemList(i).fitemid %>
		</td>				
		<td align="center">
			<%= formatdate(ocontents.FItemList(i).fregdate,"0000.00.00") %>
		</td>	
		<td align="center">
			<%= chrbyte(ocontents.FItemList(i).fcomment,50,"Y") %>
		</td>
		<td align="center">
			<%= ocontents.FItemList(i).fbest %>
		</td>			
		<td align="center">
		<% if ocontents.FItemList(i).fbestyn = "Y" then %>
			<font color="red"><b><%= ocontents.FItemList(i).fbestyn %></b></font>
		<% else %>
			<%= ocontents.FItemList(i).fbestyn %>
		<% end if %>	
		</td>		
		<td align="center">
			<%= ocontents.FItemList(i).fisusing %>
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
				<span class="list_link"><a href="?page=<%= ocontents.StartScrollPage-1 %>&isusing=<%=isusing%>&ideafileid=<%=ideafileid%>&cate_large=<%=cate_large%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + ocontents.StartScrollPage to ocontents.StartScrollPage + ocontents.FScrollCount - 1 %>
				<% if (i > ocontents.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(ocontents.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&isusing=<%=isusing%>&ideafileid=<%=ideafileid%>&cate_large=<%=cate_large%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if ocontents.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&isusing=<%=isusing%>&ideafileid=<%=ideafileid%>&cate_large=<%=cate_large%>">[next]</a></span>
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