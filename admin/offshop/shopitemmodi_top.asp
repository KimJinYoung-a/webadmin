<%
dim refaddr
refaddr = request.ServerVariables("PATH_INFO")

%>

<!-- <h2>������ �Դϴ�.</h2> -->

<table width="98%" border=0 cellspacing=1 cellpadding=2 class="a" bgcolor="#7777FF">
<tr bgcolor="#FFFFFF">
	<td width="25%">
	<% if refaddr="/admin/offshop/shopitemmodi_itemname.asp" then %>
	<b>1.��ǰ�� ����<br> (�¶��ΰ� �ٸ�����)</b>
	<% else %>
	<a href="/admin/offshop/shopitemmodi_itemname.asp?menupos=<%= menupos %>">3.��ǰ�� ����<br> (�¶��ΰ� �ٸ�����)</a>
	<% end if %>
	</td>
	<td width="25%">
	<% if refaddr="/admin/offshop/shopitemmodi_brand.asp" then %>
	<b>2.�귣�� ����<br> (�¶��ΰ� �ٸ�����)</b>
	<% else %>
	<a href="/admin/offshop/shopitemmodi_brand.asp?menupos=<%= menupos %>">4.�귣�� ����<br> (�¶��ΰ� �ٸ�����)</a>
	<% end if %>
	</td>
	<td width="25%">
	<% if refaddr="/admin/offshop/shopitemmodi_itemprice.asp" then %>
	<b>3.��ǰ���� ����<br> (�¶��ΰ� �ٸ�����)</b>
	<% else %>
	<a href="/admin/offshop/shopitemmodi_itemprice.asp?menupos=<%= menupos %>">2.��ǰ���� ����<br> (�¶��ΰ� �ٸ�����)</a>
	<% end if %>
	</td>
	<td width="25%">
	<% if refaddr="/admin/offshop/shopitemmodi.asp" then %>
	<b>4.���԰�����, ���ް����� <br> (�귣�庰�� ����.)</b>
	<% else %>
	<a href="/admin/offshop/shopitemmodi.asp?menupos=<%= menupos %>">1.���԰�����, ���ް����� <br> (�귣�庰�� ����.)</a>
	<% end if %>
	</td>
</tr>
</table>
<br>