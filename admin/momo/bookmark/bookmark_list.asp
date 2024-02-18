<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모 북마크
' Hieditor : 2009.11.20 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
Dim ocontents,i,page , bookmarkid , isusing ,coinyn
	menupos = request("menupos")
	page = request("page")
	bookmarkid = request("bookmarkidsearch")	
	isusing = request("isusing")			
	coinyn = request("coinyn")
	if page = "" then page = 1

'//북마크 리스트	
set ocontents = new cbookmark_list	
	ocontents.FPageSize = 5
	ocontents.FCurrPage = page		
	ocontents.frectbookmarkid = bookmarkid
	ocontents.frectisusing = isusing
	ocontents.frectcoinyn = coinyn
	ocontents.fbookmark_list()
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

// 삭제 
function delete_bookmarkid(upfrm){
if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}	
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.bookmarkid.value = upfrm.bookmarkid.value + frm.bookmarkid.value + "," ;
						
				}
			}
		}

			var tot;
			tot = upfrm.bookmarkid.value;
			upfrm.bookmarkid.value = ""
		var changestats;

		changestats = window.open("/admin/momo/bookmark/bookmark_process.asp?bookmarkid=" +tot + "&mode=delete" , "changestats","width=400,height=300,scrollbars=yes,resizable=yes");
		changestats.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm_search" method=get action="">
	<input type="hidden" name="bookmarkid">	
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>		
		<td align="left">
			bookmarkid:<input type="text" name="bookmarkidsearch" value="<%=bookmarkid%>" size=10>			
			&nbsp; 사용여부:
			<select name="isusing" value="<%=isusing%>">
				<option value="" <% if isusing = "" then response.write " selected" %>>사용여부</option>
				<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
			</select>
			&nbsp; 코인발급여부여부:
			<select name="coinyn" value="<%=coinyn%>">
				<option value="" <% if coinyn = "" then response.write " selected" %>>사용여부</option>
				<option value="Y" <% if coinyn = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if coinyn = "N" then response.write " selected" %>>N</option>
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
			<input type="button" onclick="delete_bookmarkid(frm_search);" value="노출안함" class="button">
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
		<td align="center">bookmarkid</td>
		<td align="center">구분</td>		
		<td align="center">사이트이름</td>
		<td align="center">사이트주소</td>
		<td align="center">사이트설명</td>			
		<td align="center">코인발급</td>	
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
			<%= ocontents.FItemList(i).fbookmarkid %><input type="hidden" name="bookmarkid" value="<%= ocontents.FItemList(i).fbookmarkid %>">
		</td>
		<td align="center">
			<%= imggubun(ocontents.FItemList(i).fgubun) %>
		</td>		
		<td align="center">
			<%= chrbyte(ocontents.FItemList(i).fsitename,30,"Y") %>
		</td>				
		<td align="center">
			<a href="http://<%=replace(ocontents.FItemList(i).fsiteaddress,"http://","")%>" target="_blank" onfocus="this.blur();">
			<%= chrbyte(ocontents.FItemList(i).fsiteaddress,50,"Y") %></a>		
		</td>	
		<td align="center">
			<%= chrbyte(ocontents.FItemList(i).fsiteinfo,100,"Y") %>
		</td>
		<td align="center">
			<%= ocontents.FItemList(i).fcoinyn %>
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
				<span class="list_link"><a href="?page=<%= ocontents.StartScrollPage-1 %>&isusing=<%=isusing%>&bookmarkid=<%=bookmarkid%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + ocontents.StartScrollPage to ocontents.StartScrollPage + ocontents.FScrollCount - 1 %>
				<% if (i > ocontents.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(ocontents.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&isusing=<%=isusing%>&bookmarkid=<%=bookmarkid%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if ocontents.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&isusing=<%=isusing%>&bookmarkid=<%=bookmarkid%>">[next]</a></span>
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