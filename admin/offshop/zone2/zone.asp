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
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/zone2/zone_cls.asp"-->

<%
Dim ozone,i,page,isusing , shopid , zonegroup ,racktype
	isusing = requestCheckVar(request("isusing"),1)
	page = requestCheckVar(request("page"),10)
	shopid = requestCheckVar(request("shopid"),32)
	zonegroup = requestCheckVar(request("zonegroup"),32)
	racktype = requestCheckVar(request("racktype"),10)
	menupos = requestCheckVar(request("menupos"),10)

if page = "" then page = 1
if isusing = "" then isusing = "Y"

'����/������
if (C_IS_SHOP) then
	
	'/���α��� ���� �̸�
	if getlevel_sn("",session("ssBctId")) > 6 then
		shopid = C_STREETSHOPID
	end if	
end if	

set ozone = new czone_list
	ozone.FPageSize = 50
	ozone.FCurrPage = page
	ozone.frectisusing = isusing
	ozone.frectshopid = shopid
	ozone.fzone_list()
%>

<script language="javascript">
	
//window.resizeTo(700, 700);

function reg(idx){
	var reg = window.open('/admin/offshop/zone2/zone_reg.asp?menupos=<%=menupos%>&idx='+idx,'reg','width=1024,height=768,scrollbars=yes,resizable=yes');
	reg.focus();	
}

function divch(divid,zoneidx){
	frmdiv.divid.value = divid;
	frmdiv.zoneidx.value = zoneidx;
	frmdiv.target="view";
	frmdiv.action='/admin/offshop/zone2/zone_manager_search.asp';
	frmdiv.submit();
}

function frmsubmit(page){
	frm.page.value=page;
	frm.submit();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmdiv" method="get" action="">
	<input type="hidden" name="divid">
	<input type="hidden" name="zoneidx">
</form>
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value=1>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<%
		'����/������
		if (C_IS_SHOP) then
		%>	
			<% if getoffshopdiv(shopid) <> "1" and shopid <> "" then %>
				* ���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
			<% end if %>
		<% else %>
			* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
		<% end if %>
		&nbsp;&nbsp;
		* ��뿩��:<select name="isusing" value="<%=isusing%>">
			<!--<option value="" <%' if isusing = "" then response.write " selected" %>>��ü</option>-->
			<option value="Y" <% if isusing = "Y" then response.write " selected" %>>Y</option>
			<option value="N" <% if isusing = "N" then response.write " selected" %>>N</option>
		</select>		
	</td>	
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frmsubmit('');">
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
		<font color="red">[�ʵ�]</font> ������ ������� [OFF]����_�������>>�����޸���Ʈ ���� ���� �����մϴ�	
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
	<td>��<br>���</td>
	<td>��<br>�����</td>
	<td>�ǻ����</td>
	<td>���׸�</td>
	<td>����<br>ũ��</td>	
	<td>����<br>������</td>
	<td>���峻<br>�����</td>
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
		<%= ozone.FItemList(i).fpyeong %>
	</td>	
	<td align="center">
		<%= ozone.FItemList(i).frealpyeong %>
	</td>
	<td align="center">
		<% if ozone.FItemList(i).fpyeong <> 0 then %>
			<%= Clng( ((ozone.FItemList(i).frealpyeong / ozone.FItemList(i).fpyeong) * 10000)) / 100 %> %
		<% end if %>
	</td>
	<td align="center">
		<%= ozone.FItemList(i).fzonename %>
	</td>	
	<td align="center">
		<%= ozone.FItemList(i).funit %>
	</td>
	<td align="center">
		<% if ozone.FItemList(i).frealpyeong<>0 then %>
			<%= Clng( ((ozone.FItemList(i).funit / ozone.FItemList(i).frealpyeong) * 10000)) / 100 %> %
		<% end if %>
	</td>	
	<td align="center">
		<% if ozone.FItemList(i).fmanagershopyn = "Y" then %>
			<div name="div<%=i%>" id="div<%=i%>">
				<img src="/images/icon_search.jpg" onmouseover="javascript:divch('div<%=i%>','<%=ozone.FItemList(i).fidx%>');">
			</div>
		<% end if %>
	</td>
	<td align="center">
		<input type="button" value="����" class="button" onclick="reg('<%= ozone.FItemList(i).fidx %>');">
	</td>	
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if ozone.HasPreScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit('<%= ozone.StartScrollPage-1 %>');">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + ozone.StartScrollPage to ozone.StartScrollPage + ozone.FScrollCount - 1 %>
			<% if (i > ozone.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(ozone.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:frmsubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if ozone.HasNextScroll then %>
			<span class="list_link"><a href="javascript:frmsubmit('<%= i %>');">[next]</a></span>
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
<iframe id="view" name="view" src="" width=0 height=0 frameborder="0" scrolling="no" ></iframe>
</table>

<%
set ozone = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->