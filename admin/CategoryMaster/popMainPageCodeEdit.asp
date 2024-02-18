<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/category_main_pageItemcls.asp" -->
<%
dim divCd, page

divCd = request("divCd")
page = request("page")

if divCd="" then divCd=0
if page="" then page=1

dim oPageDiv,oPageDivList

set oPageDiv = new CateMainPage
oPageDiv.FRectdivCd = divCd
oPageDiv.GetOnePageDivCd

set oPageDivList = new CateMainPage
oPageDivList.FPageSize=10
oPageDivList.FCurrPage= page
oPageDivList.GetPageDivList

dim i
%>
<script language='javascript'>
<!--
// 내용 검사 및 전송
function SavedivCd(frm){
    if (frm.divName.value.length<1){
        alert('구분명을 입력하세요.');
        frm.divName.focus();
        return;
    }
    
    if (frm.imgWidth.value.length<1){
        alert('이미지 사이즈W를 입력하세요.');
        frm.imgWidth.focus();
        return;
    }
    
    if (frm.imgHeight.value.length<1){
        alert('이미지 사이즈H를 입력하세요.');
        frm.imgHeight.focus();
        return;
    }
    

    if (confirm('저장 하시겠습니까?')){
        frm.submit();
    }
    
}
//-->
</script>

<table width="660" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<form name="frmdivCd" method="post" action="doMainPageCode.asp" >
<% if oPageDiv.FdivCd<>"" then %>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">구분코드</td>
    <td>
        <%= oPageDiv.FdivCd %>
        <input type="hidden" name="divCd" value="<%= oPageDiv.FdivCd %>" >
    </td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">구분명</td>
    <td>
        <input type="text" name="divName" value="<%= oPageDiv.FdivName %>" maxlength="32" size="64">
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">이미지 width</td>
    <td>
        <input type="text" name="imgWidth" value="<%= oPageDiv.FimgWidth %>" maxlength="16" size="8">
        (이미지 Width Size 숫자)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">이미지 width</td>
    <td>
        <input type="text" name="imgHeight" value="<%= oPageDiv.FimgHeight %>" maxlength="16" size="8">
        (이미지 Height Size 숫자 : 0 인경우 height 지정 안함)
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">사용형식</td>
    <td>
        <select name="divType">
        	<option value="I">상품지정</option>
        	<option value="M">이미지 선택</option>
        	<option value="B">상품지정 및 이미지추가</option>
        </select>
        <script language="javascript">
        	document.frmdivCd.divType.value="<% if oPageDiv.FdivType="" or isNull(oPageDiv.FdivType) then Response.Write "I": else Response.Write oPageDiv.FdivType: end if %>";
        </script>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td width="150" bgcolor="#DDDDFF">사용여부</td>
    <td>
        <% if oPageDiv.Fisusing="N" then %>
        <input type="radio" name="isusing" value="Y">사용함
        <input type="radio" name="isusing" value="N" checked >사용안함
        <% else %>
        <input type="radio" name="isusing" value="Y" checked >사용함
        <input type="radio" name="isusing" value="N">사용안함
        <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><input type="button" value=" 저 장 " onClick="SavedivCd(frmdivCd);"></td>
</tr>
</form>
</table>
<%
set oPageDiv = Nothing
%>
<br>

<table width="660" cellpadding="2" cellspacing="1" class="a" bgcolor="#3d3d3d">
<tr bgcolor="#FFFFFF">
    <td colspan="6" align="right"><a href="?divCd="><img src="/images/icon_new_registration.gif" border="0"></a></td>
</tr>
<tr bgcolor="#DDDDFF" align="center">
    <td width="100">code</td>
    <td width="200">코드명</td>
    <td width="150">형식</td>
    <td width="100">너비</td>
    <td width="100">높이</td>
    <td width="60">사용여부</td>
</tr>
<% for i=0 to oPageDivList.FResultCount-1 %>
<% if (CStr(oPageDivList.FItemList(i).FdivCd)=divCd) then %>
<tr bgcolor="#ECECFF" align="center">
<% elseif oPageDivList.FItemList(i).FisUsing="N" then %>
<tr bgcolor="#EEEEEE" align="center">
<% else %>
<tr bgcolor="#FFFFFF" align="center">
<% end if %>
    <td><%= oPageDivList.FItemList(i).FdivCd %></td>
    <td align="left"><a href="?divCd=<%= oPageDivList.FItemList(i).FdivCd %>&page=<%= page %>"><%= oPageDivList.FItemList(i).FdivName %></a></td>
    <td><%= oPageDivList.FItemList(i).FdivType %></td>
    <td><%= oPageDivList.FItemList(i).FimgWidth %></td>
    <td><%= oPageDivList.FItemList(i).FimgHeight %></td>
    <td><%= oPageDivList.FItemList(i).Fisusing %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="6" align="center">
    <% if oPageDivList.HasPreScroll then %>
		<a href="?page=<%= oPageDivList.StartScrollPage-1 %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + oPageDivList.StartScrollPage to oPageDivList.FScrollCount + oPageDivList.StartScrollPage - 1 %>
		<% if i>oPageDivList.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if oPageDivList.HasNextScroll then %>
		<a href="?page=<%= i %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
    </td>
</tr>
</table>

<%
set oPageDivList = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->