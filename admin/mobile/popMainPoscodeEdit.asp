<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/mobile/main_ContentsManageCls.asp" -->
<%
'###############################################
' PageName : popMainPoscodeEdit.asp
' Discription : 모바일 사이트 메인 코드관리
' History : 2010.02.23 허진원
'###############################################

dim linktype, fixtype
dim poscode, page

poscode = request("poscode")
page = request("page")

if poscode="" then poscode=0
if page="" then page=1

dim oposcode,oposcodeList

set oposcode = new CMainContentsCode
oposcode.FRectPosCode = poscode
oposcode.GetOneContentsCode

set oposcodeList = new CMainContentsCode
oposcodeList.FPageSize=20
oposcodeList.FCurrPage= page
oposcodeList.GetposcodeList

dim i
%>
<script language='javascript'>
function SavePosCode(frm){
    if (frm.poscode.value.length<1){
        alert('구분 코드 값을 입력하세요.');
        frm.poscode.focus();
        return;
    }
    
    if (frm.poscode.value*1<1){
        alert('구분 코드 값은 1 이상입니다.');
        frm.poscode.focus();
        return;
    }
    
    if (frm.posname.value.length<1){
        alert('구분명을 입력하세요.');
        frm.posname.focus();
        return;
    }
    
    if (frm.posVarname.value.length<1){
        alert('변수명을  입력하세요.');
        frm.posVarname.focus();
        return;
    }
    
    if (frm.linktype.value.length<1){
        alert('링크구분을 선택하세요.');
        frm.linktype.focus();
        return;
    }
    
    if (frm.imagewidth.value.length<1){
        alert('이미지 사이즈W를 입력하세요.');
        frm.imagewidth.focus();
        return;
    }
    
    if (frm.imageheight.value.length<1){
        alert('이미지 사이즈H를 입력하세요.');
        frm.imageheight.focus();
        return;
    }

    if (frm.useSet.value.length<1){
        alert('이미지 사용개수를 입력하세요.');
        frm.useSet.focus();
        return;
    }
    
    if (frm.fixtype.value.length<1){
        alert('반영주기를 선택하세요.');
        frm.fixtype.focus();
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

<table width="660" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmposcode" method="post" action="do_mainPosCode.asp" >
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">구분코드</td>
    <td>
        <% if oposcode.FOneItem.Fposcode<>"" then %>
        <%= oposcode.FOneItem.Fposcode %>
        <input type="hidden" name="poscode" value="<%= oposcode.FOneItem.Fposcode %>" >
        <% else %>
        <input type="text" name="poscode" value="<%= oposcode.FOneItem.Fposcode %>" maxlength="7" size="5">
        (숫자)
        <% end if %>
            
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">구분명</td>
    <td>
        <input type="text" name="posname" value="<%= oposcode.FOneItem.Fposname %>" maxlength="32" size="64">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">변수명</td>
    <td>
        <input type="text" name="posVarname" value="<%= oposcode.FOneItem.FposVarname %>" maxlength="32" size="20">
        <br>
        (영문/ 변수명으로 사용 : 띠어쓰기 금지, 특수문자 금지)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">이미지 너비</td>
    <td>
        <input type="text" name="imagewidth" value="<%= oposcode.FOneItem.Fimagewidth %>" maxlength="16" size="8">
        (이미지 Width Size 숫자)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">이미지 높이</td>
    <td>
        <input type="text" name="imageheight" value="<%= oposcode.FOneItem.Fimageheight %>" maxlength="16" size="8">
        (이미지 Height Size 숫자 : 0 인경우 height 지정 안함)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">이미지 사용개수</td>
    <td>
        <input type="text" name="useSet" value="<% if oposcode.FOneItem.FuseSet="" then Response.Write "1":Else Response.Write oposcode.FOneItem.FuseSet:End if %>" size="5">
        (XML에 사용될 개수, 일반은 1로 지정)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">링크구분</td>
    <td>
        
        <% call DrawLinktypeCombo ("linktype", oposcode.FOneItem.Flinktype, "") %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">적용구분(반영주기)</td>
    <td>
        <% call DrawFixTypeCombo ("fixtype", oposcode.FOneItem.Ffixtype, "") %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">사용여부</td>
    <td>
        <% if oposcode.FOneItem.Fisusing="N" then %>
        <input type="radio" name="isusing" value="Y">사용함
        <input type="radio" name="isusing" value="N" checked >사용안함
        <% else %>
        <input type="radio" name="isusing" value="Y" checked >사용함
        <input type="radio" name="isusing" value="N">사용안함
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" value=" 저 장 " onClick="SavePosCode(frmposcode);"></td>
</tr>
</form>
</table>
<%
set oposcode = Nothing
%>
<br>

<table width="660" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<tr bgcolor="#FFFFFF">
    <td colspan="6" align="right"><a href="?poscode="><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
<tr bgcolor="#DDDDFF">
    <td width="100">code</td>
    <td width="100">구분명</td>
    <td width="100">변수명</td>
    <td width="100">링크구분</td>
    <td width="100">반영주기</td>
    <td width="60">사용여부</td>
</tr>
<% for i=0 to oposcodeList.FResultCount-1 %>
<% if (CStr(oposcodeList.FItemList(i).FposCode)=poscode) then %>
<tr bgcolor="#9999CC">
<% else %>
<tr bgcolor="#FFFFFF">
<% end if %>
    <td ><%= oposcodeList.FItemList(i).FposCode %></td>
    <td ><a href="?poscode=<%= oposcodeList.FItemList(i).FposCode %>&page=<%= page %>"><%= oposcodeList.FItemList(i).FposName %></a></td>
    <td ><%= oposcodeList.FItemList(i).FposVarName %></td>
    <td ><%= oposcodeList.FItemList(i).getlinktypeName %></td>
    <td ><%= oposcodeList.FItemList(i).getfixtypeName %></td>
    <td ><%= oposcodeList.FItemList(i).Fisusing %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="6" align="center">
    <% if oposcodeList.HasPreScroll then %>
		<a href="?page=<%= oposcodeList.StarScrollPage-1 %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oposcodeList.StarScrollPage to oposcodeList.FScrollCount + oposcodeList.StarScrollPage - 1 %>
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
</table>

<%
set oposcodeList = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->