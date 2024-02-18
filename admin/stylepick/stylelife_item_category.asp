<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/CategoryCls.asp"-->
<%
dim cdl,cdm,cds
cdl = request("cdl")
cdm = request("cdm")
cds = request("cds")

dim oLcate
set oLcate = new CCatemanager
oLcate.GetNewCateMaster


dim oMcate
set oMcate = new CCatemanager
if (cdl<>"") then
	oMcate.GetNewCateMasterMid cdl
end if

dim oScate
set oScate = new CCatemanager
if (cdl<>"") and (cdm<>"") then
	oScate.GetNewCateMasterSmall cdl,cdm
end if

dim i,currposStr

if cdl<>"" then
	'currposStr = oLcate.GetNewCateCurrentPos(cdl,cdm,cds)
end if

Dim vNowCateName
%>
<script language='javascript'>

</script>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
	<td width="22%" valign=top>
		<table border=1 cellspacing=1 cellpadding=0 class=a width="100%" bgcolor="#FFFFFF">
		<% for i=0 to oLcate.FResultCount-1 %>
		<tr>
			<% if oLcate.FItemList(i).Fcdlarge=cdl then %>
			<td><b><a href="?menupos=<%=request("menupos")%>&cdl=<%= oLcate.FItemList(i).Fcdlarge %>">[<%= oLcate.FItemList(i).Fcdlarge %>]<%= oLcate.FItemList(i).Fnmlarge %></a></b></td>
			<%
				vNowCateName = "[" & oLcate.FItemList(i).Fcdlarge & "]" & oLcate.FItemList(i).Fnmlarge
			else %>
			<td><a href="?menupos=<%=request("menupos")%>&cdl=<%= oLcate.FItemList(i).Fcdlarge %>">[<%= oLcate.FItemList(i).Fcdlarge %>]<%= oLcate.FItemList(i).Fnmlarge %></a></td>
			<% end if %>
		</tr>
		<% next %>
		</table>
	</td>
	<td width="22%" valign=top>
		<table border=1 cellspacing=1 cellpadding=1 class=a width="100%" bgcolor="#FFFFFF">
		<% for i=0 to oMcate.FResultCount-1 %>
		<tr>
			<% if oMcate.FItemList(i).Fcdmid=cdm then %>
				<td><%= oMcate.FItemList(i).ForderNo %></td>
				<td><b><a href="?menupos=<%=request("menupos")%>&cdl=<%= oMcate.FItemList(i).Fcdlarge %>&cdm=<%= oMcate.FItemList(i).Fcdmid %>">[<%= oMcate.FItemList(i).Fcdmid %>]<%= oMcate.FItemList(i).Fnmlarge %></a></b></td>
			<%
				vNowCateName = vNowCateName & " - [" & oMcate.FItemList(i).Fcdmid & "]" & oMcate.FItemList(i).Fnmlarge
			else %>
				<td><%= oMcate.FItemList(i).ForderNo %></td>
				<td><a href="?menupos=<%=request("menupos")%>&cdl=<%= oMcate.FItemList(i).Fcdlarge %>&cdm=<%= oMcate.FItemList(i).Fcdmid %>">[<%= oMcate.FItemList(i).Fcdmid %>]<%= oMcate.FItemList(i).Fnmlarge %></a></td>
			<% end if %>
		</tr>
		<% next %>
		</table>
	</td>
	<td width="22%" valign=top>
		<table border=1 cellspacing=1 cellpadding=1 class=a width="100%" bgcolor="#FFFFFF">
		<% for i=0 to oScate.FResultCount-1 %>
		<tr>
		<% if oScate.FItemList(i).Fcdsmall=cds then %>
			<td><%= oScate.FItemList(i).ForderNo %></td>
			<td><b><a href="?menupos=<%=request("menupos")%>&cdl=<%= oScate.FItemList(i).Fcdlarge %>&cdm=<%= oScate.FItemList(i).Fcdmid %>&cds=<%= oScate.FItemList(i).Fcdsmall %>">[<%= oScate.FItemList(i).Fcdsmall %>]<%= oScate.FItemList(i).Fnmlarge %></b></a></td>
		<%
			vNowCateName = vNowCateName & " - [" & oScate.FItemList(i).Fcdsmall & "]" & oScate.FItemList(i).Fnmlarge
		else %>
			<td><%= oScate.FItemList(i).ForderNo %></td>
			<td><a href="?menupos=<%=request("menupos")%>&cdl=<%= oScate.FItemList(i).Fcdlarge %>&cdm=<%= oScate.FItemList(i).Fcdmid %>&cds=<%= oScate.FItemList(i).Fcdsmall %>">[<%= oScate.FItemList(i).Fcdsmall %>]<%= oScate.FItemList(i).Fnmlarge %></a></td>
		<% end if %>
			<td width=20><%= oScate.FItemList(i).Fcatecnt %></td>
		</tr>
		<% next %>
		</table>
	</td>
</tr>
</table>
<input type="hidden" id="nowcatename" name="nowcatename" value="<%=vNowCateName%>">
<% If cdl <> "" Then %>
<iframe name=imatchitem src="stylelife_item_list.asp?cdl=<%= cdl %>&cdm=<%= cdm %>&cds=<%= cds %>" width="100%" height="100%" frameborder="0" marginheight="0" marginwidth="0" scrolling="no" onload="resizeIfr(this, 10)"></iframe>
<%
End If

set oLcate = Nothing
set oMcate = Nothing
set oScate = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->