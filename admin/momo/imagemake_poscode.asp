<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 감성모모
' Hieditor : 2009.11.11 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/momo/momo_cls.asp"-->

<%
dim linktype, fixtype
dim poscode, page

poscode = request("poscode")
page = request("page")

if poscode="" then poscode=0
if page="" then page=1

dim oposcode,oposcodeList

set oposcode = new cmomo_list
	oposcode.FRectPosCode = poscode
	oposcode.fposcode_oneitem

set oposcodeList = new cmomo_list
	oposcodeList.FPageSize=10
	oposcodeList.FCurrPage= page
	oposcodeList.fposcode_list

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
    
    
    if (frm.imagetype.value.length<1){
        alert('링크타입을 선택하세요.');
        frm.imagetype.focus();
        return;
    }
    
    if (frm.imagewidth.value.length<1){
        alert('이미지 사이즈 가로를 입력하세요.');
        frm.imagewidth.focus();
        return;
    }
    
    if (frm.imageheight.value.length<1){
        alert('이미지 사이즈 세로를 입력하세요.');
        frm.imageheight.focus();
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
<form name="frmposcode" method="post" action="/admin/momo/imagemake_poscode_process.asp" >
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
    <td width="150" bgcolor="#DDDDFF">사용할이미지수</td>
    <td>
        <input type="text" name="imagecount" value="<%= oposcode.FOneItem.fimagecount %>" maxlength="2" size="2">
        (숫자만 입력 하세요)
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
    <td width="150" bgcolor="#DDDDFF">링크타입</td>
    <td>
        <select name="imagetype">
	        <option value="" <% if oposcode.FOneItem.fimagetype = "" then response.write " selected" %>>선택</option>
	        <option value="map" <% if oposcode.FOneItem.fimagetype = "map" then response.write " selected" %>>map</option>
	        <option value="link" <% if oposcode.FOneItem.fimagetype = "link" then response.write " selected" %>>link</option>                
	        <option value="flash" <% if oposcode.FOneItem.fimagetype = "flash" then response.write " selected" %>>flash</option>
	        <option value="multi" <% if oposcode.FOneItem.fimagetype = "multi" then response.write " selected" %>>multi</option>
        </select>
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
    <td colspan="2" align="center"><input type="button" value=" 저 장 " class="button" onClick="SavePosCode(frmposcode);"></td>
</tr>
</form>
</table>
<%
set oposcode = Nothing
%>
<br>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF">
    <td colspan="6" align="right"><a href="?poscode="><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
<% if oposcodeList.FResultCount > 0 then %>
	<tr bgcolor="#DDDDFF">
	    <td width="100">code</td>
	    <td width="100">구분명</td>
	    <td width="100">링크타입</td>
	    <td width="60">Image수</td>
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
		    <td ><%= oposcodeList.FItemList(i).fimagetype %></td>
		    <td ><%= oposcodeList.FItemList(i).fimagecount %></td>
		    <td ><%= oposcodeList.FItemList(i).Fisusing %></td>
		</tr>
	<% next %>
    <tr valign="top" bgcolor="#FFFFFF">		
		<td valign="bottom" align="center" colspan=5>
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
	set oposcodeList = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
