<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : 지문인식 근태관리
' Hieditor : 2011.03.22 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/member/fingerprints/fingerprints_cls.asp" -->

<%
dim linktype, fixtype ,placeid, page ,oposcode,oposcodeList ,i , placeiname ,validpart
dim mode , isusing
	placeid = requestCheckVar(request("placeid"),10)
	page = requestCheckVar(request("page"),10)
	mode = requestCheckVar(request("mode"),32)
	placeiname = requestCheckVar(request("placeiname"),32)
	validpart = requestCheckVar(request("validpart"),10)
	isusing = requestCheckVar(request("isusing"),1)

	if page="" then page=1
	if mode = "" then mode = "ADD"
		
set oposcode = new cfingerprints_list
	oposcode.FRectplaceid = placeid
	
	if placeid <> "" and mode = "EDIT" then
		oposcode.fposcode_oneitem
		
		if oposcode.ftotalcount > 0 then
			placeid = oposcode.FOneItem.fplaceid
			placeiname = oposcode.FOneItem.fplaceiname
			validpart = oposcode.FOneItem.fvalidpart
			isusing = oposcode.FOneItem.fisusing
		end if
	end if

set oposcodeList = new cfingerprints_list
	oposcodeList.FPageSize=10
	oposcodeList.FCurrPage= page
	oposcodeList.fposcode_list
%>

<script type='text/javascript'>

function SavePosCode(frm){
    if (frm.placeid.value.length<1){
        alert('번호를 입력하세요.');
        frm.placeid.focus();
        return;
    }

    if (frm.placeiname.value.length<1){
        alert('구분명을 입력하세요.');
        frm.placeiname.focus();
        return;
    }
        
    if (frm.validpart.value.length<1){
        alert('파트번호를 입력하세요.');
        frm.validpart.focus();
        return;
    }

    if (frm.isusing.value.length<1){
        alert('사용여부를 선택하세요.');
        frm.isusing.focus();
        return;
    }            
    
    if (confirm('저장 하시겠습니까?')){
        frm.submit();
    }
    
}

function ChangeLinktype(){
    // Do nothing
}

</script>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<form name="frmposcode" method="post" action="/common/member/fingerprints/fingerprints_poscode_process.asp">
<input type="hidden" name="mode" value="<%= mode %>" >
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">번호</td>
    <td>
		<% if mode = "EDIT" then %>
			<%= placeid %>
			<input type="hidden" name="placeid" value="<%= placeid %>" >
		<% else %>
			<input type="text" name="placeid" value="<%= placeid %>" maxlength="10" size="10">(숫자)			
		<% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">구분명</td>
    <td>
        <input type="text" name="placeiname" value="<%= placeiname %>" maxlength="32" size="64">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">파트번호</td>
    <td>
        <input type="text" name="validpart" value="<%= validpart %>" maxlength="10" size="10">(숫자)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">사용여부</td>
    <td>
        <select name="isusing">
        	<option value="" <% if isusing = "" then response.write " selected" %>>선택</option>
        	<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
        	<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
        </select>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center">
    	<input type="button" value=" 저 장 " onClick="SavePosCode(frmposcode);" class="button">
    </td>
</tr>
</form>
</table>
<br>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF">
    <td colspan="6" align="right"><a href="?mode=ADD"><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
<% if oposcodeList.FResultCount > 0 then %>
<tr bgcolor="#DDDDFF" align="center">
    <td>번호</td>
    <td>구분명</td>
    <td>파트번호</td>
    <td>사용여부</td>
</tr>
<% for i=0 to oposcodeList.FResultCount-1 %>

<% if oposcodeList.FItemList(i).fplaceid=placeid then %>
	<tr bgcolor="#9999CC" align="center">
<% else %>
	<tr bgcolor="#FFFFFF" align="center">
<% end if %>

    <td ><%= oposcodeList.FItemList(i).fplaceid %></td>
    <td ><a href="?placeid=<%= oposcodeList.FItemList(i).fplaceid %>&mode=EDIT"><%= oposcodeList.FItemList(i).fplaceiname %></a></td>
    <td ><%= oposcodeList.FItemList(i).fvalidpart %></td>
    <td ><%= oposcodeList.FItemList(i).fisusing %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
   <td valign="bottom" align="center" colspan=20>
	    <% if oposcodeList.HasPreScroll then %>
			<a href="?page=<%= oposcodeList.StartScrollPage-1 %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
	
		<% for i=0 + oposcodeList.StartScrollPage to oposcodeList.FScrollCount + oposcodeList.StartScrollPage - 1 %>
			<% if i>oposcodeList.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>">[<%= i %>]</a>
			<% end if %>
		<% next %>
	
		<% if oposcodeList.HasNextScroll then %>
			<a href="?page=<%= i %>">[next]</a>
		<% else %>
			[next]
		<% end if %>
    </td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td align="center">내용이 없습니다.</td>
	</tr>	
<% end if %>
</table>

<%
set oposcode = Nothing
set oposcodeList = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
