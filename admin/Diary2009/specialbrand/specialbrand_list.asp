<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 스페셜 브랜드 리스트
' History : 2016.09.07 유태욱 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/diary2009/classes/specialbrandCls.asp"-->
<%
dim research, isusing, page, brandid

	isusing = requestcheckvar(request("isusing"),1)
	research= requestcheckvar(request("research"),2)
	page    = requestcheckvar(request("page"),16)
	brandid    = requestcheckvar(request("brandid"),32)

if ((research="") and (isusing="")) then
    isusing = "Y"
end if

if page="" then page=1

dim oMainContents
set oMainContents = new DiaryCls
	oMainContents.FPageSize = 10
	oMainContents.FCurrPage = page
	oMainContents.FRectIsusing = isusing
	oMainContents.FRectbrandid = brandid
	oMainContents.fcontents_list

dim i
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<script type="text/javascript">
	//이미지신규등록 & 수정
	function AddNewMainContents(idx){
		var AddNewMainContents = window.open('/admin/diary2009/specialbrand/imagemake_specialbrand.asp?idx='+ idx,'AddNewMainContents','width=1024,height=768,scrollbars=yes,resizable=yes');
		AddNewMainContents.focus();
	}
	document.domain ='10x10.co.kr';
</script>
<div class="contSectFix scrl">
	<div class="pad20">
		<table class="tbType1 listTb">
			<form name="frm" method="get" action="">
			<input type="hidden" name="page" value="1">
			<input type="hidden" name="research" value="on">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="idx">
			<tr align="center" bgcolor="#FFFFFF" >
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
				<td style="text-align:left;">
					사용구분
					<select name="isusing">
					<option value="">전체
					<option value="Y" <% if isusing="Y" then response.write "selected" %> >사용함
					<option value="N" <% if isusing="N" then response.write "selected" %> >사용안함
					</select>
				</td>
				<td width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
				</td>
			</tr>
			</form>
		</table>
		<div class="tPad15">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
				<tr>
					<td align="right">
						<a href="javascript:AddNewMainContents('0');"><img src="/images/icon_new_registration.gif" border="0"></a>
					</td>
				</tr>
			</table>
		</div>
		<div class="tPad15">
			<table class="tbType1 listTb">
				<% if oMainContents.FResultCount > 0 then %>
				<tr height="25" bgcolor="FFFFFF">
					<td colspan="8" style="text-align:left;">
						검색결과 : <b><%= oMainContents.FTotalCount %></b>
					</td>
				</tr>
				<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
					<td align="center">Idx</td>
					<td align="center">브랜드ID</td>
					<td align="center">대표이미지</td>
					<td align="center">브랜드설명</td>
					<td align="center">사용여부</td>
					<td align="center">등록일</td>
					<td align="center">수정</td>
				</tr>
				<% for i=0 to oMainContents.FResultCount - 1 %>
					<tr <% if oMainContents.FItemList(i).Fisusing="N" then %>bgcolor="<%= adminColor("dgray") %>"<% else %>bgcolor="#FFFFFF" style="cursor:pointer;" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background="#FFFFFF";<% end if %>>
						<td align="center"><%= oMainContents.FItemList(i).Fidx %></td>
						<td align="center"><%= oMainContents.FItemList(i).Fbrandid %></td>
						<td align="center"><img src="<%=uploadUrl%>/diary/specialbrand/<%= oMainContents.FItemList(i).fmainbrandimg %>" border="0" width="70" height="70"></td>
						<td align="center"><%= oMainContents.FItemList(i).fbrandtext %></td>
						<td align="center"><%= oMainContents.FItemList(i).FIsusing %></td>
						<td align="center"><%= oMainContents.FItemList(i).fregdate %></td>
						<td align="center">
							<input type="button" value="수정" onclick="AddNewMainContents('<%= oMainContents.FItemList(i).Fidx %>'); return false;">
						</td>
					</tr>
				<% next %>
				<% else %>
				<tr bgcolor="#FFFFFF">
					<td colspan="8" align="center" class="page_link">[검색결과가 없습니다.]</td>
				</tr>
				<% end if %>
				<tr height="25" bgcolor="FFFFFF">
					<td colspan="8" align="center">
						<% if oMainContents.HasPreScroll then %>
							<span class="list_link"><a href="?page=<%= oMainContents.StartScrollPage-1 %>">[pre]</a></span>
						<% else %>
						[pre]
						<% end if %>
						<% for i = 0 + oMainContents.StartScrollPage to oMainContents.StartScrollPage + oMainContents.FScrollCount - 1 %>
							<% if (i > oMainContents.FTotalpage) then Exit for %>
							<% if CStr(i) = CStr(oMainContents.FCurrPage) then %>
							<span class="page_link"><font color="red"><b><%= i %></b></font></span>
							<% else %>
							<a href="?page=<%= i %>" class="list_link"><font color="#000000"><%= i %></font></a>
							<% end if %>
						<% next %>
						<% if oMainContents.HasNextScroll then %>
							<span class="list_link"><a href="?page=<%= i %>">[next]</a></span>
						<% else %>
						[next]
						<% end if %>
					</td>
				</tr>
			</table>
		</div>
	</div>
</div>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->

