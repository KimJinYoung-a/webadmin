<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �𺰱�������
' Hieditor : 2010.12.29 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/zone/zone_cls.asp"-->

<%
Dim ozone,i,page,isusing , parameter , shopid , zonegroup ,racktype ,menupos
	isusing = requestCheckVar(request("isusing"),1)
	page = requestCheckVar(request("page"),10)
	shopid = requestCheckVar(request("shopid"),32)
	zonegroup = requestCheckVar(request("zonegroup"),10)
	racktype = requestCheckVar(request("racktype"),10)
	menupos = requestCheckVar(request("menupos"),10)

if page = "" then page = 1
if isusing = "" then isusing = "Y"
	
set ozone = new czone_list
	ozone.FPageSize = 20
	ozone.FCurrPage = page
	ozone.frectisusing = isusing
	ozone.frectzonegroup = zonegroup
	ozone.frectracktype = racktype
	ozone.frectshopid = shopid
	ozone.fzone_list()
	
parameter = "isusing="&isusing&"&zonegroup="&zonegroup&"&racktype="&racktype&"&shopid="&shopid	
%>

<script language="javascript">

function reg(idx){
	location.href='zone_reg.asp?idx='+idx+'&menupos=<%=menupos%>';
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		����:<% drawSelectBoxOffShop "shopid",shopid %>
		�׷�:<% drawSelectBoxOffShopzonegroup "zonegroup",zonegroup,"" %>
		�Ŵ�Ÿ��:<% drawSelectBoxOffShopracktype "racktype",racktype,"" %>
		<Br>��뿩��:<select name="isusing" value="<%=isusing%>">
			<!--<option value="" <%' if isusing = "" then response.write " selected" %>>��ü</option>-->
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>		
	</td>	
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<br>		
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">		
	</td>
	<td align="right">	
		<input type="button" class="button" value="�űԵ��" onclick="reg('');">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= ozone.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= ozone.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��ȣ</td>
	<td>SHOPID</td>
	<td>�׷�</td>	
	<td>�Ŵ�Ÿ��</td>
	<td>�󼼱���</td>
	<td>UNIT</td>	
	<td>���</td>
</tr>
<% if ozone.FresultCount>0 then %>
<% for i=0 to ozone.FresultCount-1 %>
<% if ozone.FItemList(i).fisusing = "Y" then %>
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>
<% else %>    
<tr align="center" bgcolor="#FFFFaa" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#FFFFaa';>
<% end if %>
	<td align="center">
		<%= ozone.FItemList(i).fidx %>
	</td>		
	<td align="center">
		<%= ozone.FItemList(i).fshopid %>
	</td>
	<td align="center">
		<%= getOffShopzonegroup(ozone.FItemList(i).fzonegroup) %>
	</td>	
		
	<td align="center">
		<%= getOffShopracktype(ozone.FItemList(i).fracktype) %>
	</td>
	<td align="center">
		<%= ozone.FItemList(i).fzonename %>
	</td>
	<td align="center">
		<%= ozone.FItemList(i).funit %>
	</td>
	<td align="center">
		<input type="button" value="����" class="button" onclick="reg('<%= ozone.FItemList(i).fidx %>');">
	</td>	
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if ozone.HasPreScroll then %>
			<span class="list_link"><a href="?page=<%= ozone.StartScrollPage-1 %>&<%=parameter%>">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + ozone.StartScrollPage to ozone.StartScrollPage + ozone.FScrollCount - 1 %>
			<% if (i > ozone.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(ozone.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="?page=<%= i %>&<%=parameter%>" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if ozone.HasNextScroll then %>
			<span class="list_link"><a href="?page=<%= i %>&<%=parameter%>">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<%
set ozone = nothing
%>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->