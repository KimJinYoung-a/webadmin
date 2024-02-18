<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->

<%
	Dim vItemID, strSQL, vImage, vDiv
	vItemID = Request("itemid")
	vDiv	= Request("div")
	
	If vItemID = "" Then
		Response.End
	End If
	
	strSQL = "SELECT smallimage FROM [db_item].[dbo].[tbl_item] WHERE itemid = '" & vItemID & "'"
	rsget.Open strSQL,dbget,1
	
	If Not rsget.Eof Then
		vImage = "http://webimage.10x10.co.kr/image/small/" + GetImageSubFolderByItemid(vItemID) + "/" + rsget("smallimage")
	End If
	rsget.close
%>

<script language="javascript">
parent.document.getElementById("<%=vDiv%>").innerHTML = "<img src='<%=vImage%>'>";
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->