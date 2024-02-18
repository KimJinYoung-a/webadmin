<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : Culture Station 에디터
' Hieditor : 2009.04.06 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/culturestation/culturestation_class.asp"-->

<%
Dim oip,i,page,isusing,editor_no_search , editor_no_count
	editor_no_search = requestCheckVar(request("editor_no_search"),10)
	isusing = requestCheckVar(request("isusing"),1)
	editor_no_count = requestCheckVar(request("editor_no_count"),10)
	menupos = requestCheckVar(getNumeric(request("menupos")),10)
	page = requestCheckVar(getNumeric(request("page")),10)
	if page = "" then page = 1

'// 이벤트 리스트
set oip = new ceditor_list
	oip.FPageSize = 20
	oip.FCurrPage = page
	oip.frectisusing = isusing
	oip.frecteditor_no = editor_no_search
	oip.frecteditor_no_count = editor_no_count
	oip.feditor_list()
%>

<script type='text/javascript'>

function editor_edit(editor_no){
	var editor_edit = window.open('/admin/culturestation/editor_edit.asp?editor_no='+editor_no,'addreg','width=1200,height=700,scrollbars=yes,resizable=yes');
	editor_edit.focus();
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


function comment_list(editor_no){

	 var comment_list = window.open('/admin/culturestation/editor_comment_list.asp?editor_no='+editor_no,'comment_list','width=800,height=600,scrollbars=yes,resizable=yes');
	 comment_list.focus();

}

function AssignbestReal(upfrm){
if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}	
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.editor_no.value = upfrm.editor_no.value + frm.editor_no.value + "," ;
						
				}
			}
		}

			var tot;
			tot = upfrm.editor_no.value;
			upfrm.editor_no.value = ""
		var AssignbestReal;

		AssignbestReal = window.open("<%=wwwUrl%>/chtml/culturestation_editorbestmake.asp?editor_no=" +tot, "AssignbestReal","width=400,height=300,scrollbars=yes,resizable=yes");
		AssignbestReal.focus();
}
function RefreshMainCorItemRec(upfrm){
	if (!CheckSelected()){
		alert('선택아이템이 없습니다.');
		return;
	}
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.editor_no.value = upfrm.editor_no.value + frm.editor_no.value + "," ;
						
				}
			}
		}

	var tot;
	tot = upfrm.editor_no.value;
	upfrm.editor_no.value = ""
	var AssignbestReal;

	AssignReal = window.open("<%=wwwUrl%>/chtml/make_curture_editor.asp?editor_no=" +tot, "AssignReal","width=400,height=300,scrollbars=yes,resizable=yes");
	AssignReal.focus();
}
</script>

<!-- 검색 시작 -->
<form name="frm" method=get action="" style="margin:0px;">
<input type="hidden" name="editor_no">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			&nbsp;스토리번호: <input type="text" name="editor_no_search" value="<%= editor_no_search%>" size="10"> 			
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<select name="isusing" value="<%=isusing%>">
				<option value="" <% if isusing = "" then response.write " selected" %>>사용여부</option>
				<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
				<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
			</select>
			<select name="editor_no_count" value="<%=editor_no_count%>">
				<option value="" <% if editor_no_count = "" then response.write " selected" %>>comment우선정렬</option>
				<option value="Y" <% if editor_no_count = "Y" then response.write " selected" %>>Y</option>
			</select>			
		</td>
	</tr>
</table>
</form>
<!-- 검색 끝 -->
		
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<a href="javascript:AssignbestReal(frm);"><img src="/images/refreshcpage.gif" border="0">왼쪽메뉴베스트생성</a>
		</td>
		<td align="right">	
			<input type="button" class="button" value="editor 신청리스트" onclick="window.open('/admin/culturestation/apply/','apply','width=1200,height=700,scrollbars=yes,resizable=yes');">&nbsp;&nbsp;&nbsp;
			<input type="button" class="button" value="editor등록" onclick="editor_edit('');">
		</td>
	</tr>
	<tr>
		<td align="left">
			<img src="/images/icon_reload.gif" onClick="javascript:RefreshMainCorItemRec(frm);" style="cursor:pointer" align="absmiddle" alt="XML만들기">2015프론트에 적용
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if oip.FresultCount>0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oip.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= oip.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
   		<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td align="center">Eventcode</td>
		<td align="center">기본Image</td>	
		<td align="center">배너Image</td>	
		<td align="center">이벤트명</td>
		<td align="center">등록일</td>
		<td align="center">사용여부</td>	
		<td align="center">코맨트수</td>
    </tr>
	<% for i=0 to oip.FresultCount-1 %>
	<form action="" name="frmBuyPrc<%=i%>" method="get" style="margin:0px;">			
	
    <% if oip.FItemList(i).fisusing = "Y" then %>
    <tr align="center" bgcolor="#FFFFFF">
    <% else %>    
    <tr align="center" bgcolor="#FFFFaa">
	<% end if %>
		<td align="center">
			<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
		</td>
		<td align="center">
			<a href="javascript:editor_edit('<%= oip.FItemList(i).feditor_no %>');"><%= oip.FItemList(i).feditor_no %></a>
			<input type="hidden" name="editor_no" value="<%= oip.FItemList(i).feditor_no %>">
		</td>		
		<td align="center">
			<a href="javascript:editor_edit('<%= oip.FItemList(i).feditor_no %>');">
			<image src="<%=webImgUrl%>/culturestation/editor/2009/list/<%= oip.FItemList(i).fimage_list %>" width="40" height="40" border=0></a>
		</td>	
		<td align="center">
			<a href="javascript:editor_edit('<%= oip.FItemList(i).feditor_no %>');">
			<image src="<%=webImgUrl%>/culturestation/editor/2009/barner/<%= oip.FItemList(i).fimage_barner %>" width="40" height="40" border=0></a>
		</td>
		<td align="center">
			<a href="javascript:editor_edit('<%= oip.FItemList(i).feditor_no %>');"><%= ReplaceBracket(oip.FItemList(i).feditor_name) %></a>
		</td>
		<td align="center"><%= left(oip.FItemList(i).fregdate,10) %></td>
		<td align="center"><%= oip.FItemList(i).fisusing %></td>
		<td align="center">
		<a href="javascript:comment_list(<%= oip.FItemList(i).feditor_no %>);"><%= oip.FItemList(i).feditor_no_count %></a>
		</td>
    </tr>   
	</form>
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if oip.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= oip.StartScrollPage-1 %>&isusing=<%=isusing%>&editor_no_count=<%=editor_no_count%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oip.StartScrollPage to oip.StartScrollPage + oip.FScrollCount - 1 %>
				<% if (i > oip.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oip.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&isusing=<%=isusing%>&editor_no_count=<%=editor_no_count%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oip.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&isusing=<%=isusing%>&editor_no_count=<%=editor_no_count%>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

