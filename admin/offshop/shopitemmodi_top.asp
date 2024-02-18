<%
dim refaddr
refaddr = request.ServerVariables("PATH_INFO")

%>

<!-- <h2>수정중 입니다.</h2> -->

<table width="98%" border=0 cellspacing=1 cellpadding=2 class="a" bgcolor="#7777FF">
<tr bgcolor="#FFFFFF">
	<td width="25%">
	<% if refaddr="/admin/offshop/shopitemmodi_itemname.asp" then %>
	<b>1.상품명 수정<br> (온라인과 다른내용)</b>
	<% else %>
	<a href="/admin/offshop/shopitemmodi_itemname.asp?menupos=<%= menupos %>">3.상품명 수정<br> (온라인과 다른내용)</a>
	<% end if %>
	</td>
	<td width="25%">
	<% if refaddr="/admin/offshop/shopitemmodi_brand.asp" then %>
	<b>2.브랜드 수정<br> (온라인과 다른내용)</b>
	<% else %>
	<a href="/admin/offshop/shopitemmodi_brand.asp?menupos=<%= menupos %>">4.브랜드 수정<br> (온라인과 다른내용)</a>
	<% end if %>
	</td>
	<td width="25%">
	<% if refaddr="/admin/offshop/shopitemmodi_itemprice.asp" then %>
	<b>3.상품가격 수정<br> (온라인과 다른내용)</b>
	<% else %>
	<a href="/admin/offshop/shopitemmodi_itemprice.asp?menupos=<%= menupos %>">2.상품가격 수정<br> (온라인과 다른내용)</a>
	<% end if %>
	</td>
	<td width="25%">
	<% if refaddr="/admin/offshop/shopitemmodi.asp" then %>
	<b>4.매입가수정, 공급가수정 <br> (브랜드별로 가능.)</b>
	<% else %>
	<a href="/admin/offshop/shopitemmodi.asp?menupos=<%= menupos %>">1.매입가수정, 공급가수정 <br> (브랜드별로 가능.)</a>
	<% end if %>
	</td>
</tr>
</table>
<br>