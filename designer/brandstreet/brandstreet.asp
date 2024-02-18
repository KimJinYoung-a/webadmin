<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 업체 브랜드페이지 관리 
' History : 2009.03.26 한용민 생성
'###########################################################
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/brandstreet/brandstreet_upche_cls.asp"-->

<%
dim page , isusing , types
	types = requestCheckVar(request("types"),50)
	isusing = requestCheckVar(request("isusing"),30)
	page    = requestCheckVar(request("page"),10)
	
if page="" then page=1
if types = "" then types = 1
if isusing = "" then isusing = "Y"
dim oMainContents
set oMainContents = new cbrandstreet_list
	oMainContents.FPageSize = 6
	oMainContents.FCurrPage = page
	oMainContents.frectisusing = isusing
	oMainContents.frecttype = types	
	oMainContents.frectmakerid = session("ssBctId")		
	oMainContents.fcontents_list

dim i
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

// 일괄 사용안함
function display_no(upfrm){
	if (!CheckSelected()){
			alert('선택아이템이 없습니다.');
			return;
		}
	var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.fidx.value = upfrm.fidx.value + frm.idx.value + "," ;
					
				}
			}
		}
	var tot;
	tot = upfrm.fidx.value;
	upfrm.fidx.value = ""
	var display_no;
	display_no = window.open("/designer/brandstreet/brandstreet_process.asp?itemid=" +tot + '&mode=isusing_no', "display_no","width=400,height=300,scrollbars=yes,resizable=yes");
	display_no.focus();
}


//신규등록 
function AddNewMainContents(){
    var AddNewMainContents = window.open('/designer/brandstreet/brandstreet_upcheitem.asp','AddNewMainContents','width=600,height=768,scrollbars=yes,resizable=yes');
    AddNewMainContents.focus();
}


</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="fidx">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			적용구분
			<select name="types">
			<option value="">전체</option>
			<option value="1" <% if types="1" then response.write "selected" %>>중단배너</option>
			</select>
		    노출여부
			<select name="isusing">
			<option value="">전체</option>
			<option value="Y" <% if isusing="Y" then response.write "selected" %> >사용함</option>
			<option value="N" <% if isusing="N" then response.write "selected" %> >사용안함</option>
			</select>

		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
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
			<input type="button" onclick="display_no(frm);" value="일괄노출안함" class="button" >	
		</td>
		<td align="right">		
			<input type="button" onclick="AddNewMainContents();" value="신규등록" class="button" >		
		</td>
	</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oMainContents.FResultCount > 0 then %> 
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oMainContents.FTotalCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
 		<td align="center"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	    <td align="center">Image</td>
	    <td align="center">구분명</td>
	    <td align="center">상품코드</td>
	    <td align="center">상품명</td>
	    <td align="center">노출여부</td>
	    
    </tr>
    <tr align="center" bgcolor="#FFFFFF">
	<% for i=0 to oMainContents.FResultCount - 1 %>
	<form action="" name="frmBuyPrc<%=i%>" method="get">			<!--for문 안에서 i 값을 가지고 루프-->	 		
		<% if oMainContents.FItemList(i).FIsusing="N" then %>
			<tr bgcolor="#DDDDDD">
		<% else %>
			<tr bgcolor="#FFFFFF">
		<% end if %>	
		<td align="center">
			<input type="checkbox" name="cksel" onClick="AnCheckClick(this);">
			<input type="hidden" name="idx" value="<%= oMainContents.FItemList(i).Fidx %>">
		</td>		
	    <td align="center">
	    	<img width=40 height=40 src="<%= oMainContents.FItemList(i).fsmallimage %>" border="0">
	    </td>
	    <td align="center">
	    	<%
	    	if oMainContents.FItemList(i).ftype = "1" then
	    		response.write "중단배너"
	    	end if 
	    	%>
	    </td>
	    <td align="center"><%= oMainContents.FItemList(i).fitemid %></td>
	    <td align="center"><%= oMainContents.FItemList(i).fitemname %></td>
	    <td align="center"><%= oMainContents.FItemList(i).fisusing %></td>
	    
	</tr>
	</form>	
	<% next %>
    </tr>   
    
<% else %>

	<tr bgcolor="#FFFFFF">
		<td colspan="11" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
	       	<% if oMainContents.HasPreScroll then %>
				<span class="list_link"><a href="?page=<%= oMainContents.StartScrollPage-1 %>&isusing=<%=isusing%>&types=<%=types%>">[pre]</a></span>
			<% else %>
			[pre]
			<% end if %>
			<% for i = 0 + oMainContents.StartScrollPage to oMainContents.StartScrollPage + oMainContents.FScrollCount - 1 %>
				<% if (i > oMainContents.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oMainContents.FCurrPage) then %>
				<span class="page_link"><font color="red"><b><%= i %></b></font></span>
				<% else %>
				<a href="?page=<%= i %>&isusing=<%=isusing%>&types=<%=types%>" class="list_link"><font color="#000000"><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oMainContents.HasNextScroll then %>
				<span class="list_link"><a href="?page=<%= i %>&isusing=<%=isusing%>&types=<%=types%>">[next]</a></span>
			<% else %>
			[next]
			<% end if %>
		</td>
	</tr>
</table>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

