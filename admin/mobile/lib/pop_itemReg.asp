<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim idx, vQuery, i, vCount , number
idx  = requestCheckVar(request("idx"),10)
number = requestCheckVar(request("number"),10)

IF idx = "" THEN
	Response.Write "<script>alert('�߸��� ����Դϴ�.\nNo. ��ȣ�� �־�� �մϴ�.');</script>"
	dbget.close()
	Response.End
END IF	
IF IsNumeric(idx) = False THEN
	Response.Write "<script>alert('�߸��� ����Դϴ�.\nNo. ��ȣ�� �־�� �մϴ�.');</script>"
	dbget.close()
	Response.End
END IF
%>
<script>
	function jsGoedit(idx){
		window.opener.document.location.href='/admin/mobile/showbanner/popShowbannerEdit.asp?idx=' + idx + '#itemlist';
		self.close();
	}
</script>
<center><b>Idx: <%=idx%> </b> ��ǰ Ȯ��</center>
<%If idx <> "" Or idx <> 0 Then %><div style="padding:0 10 10 0"><input type="button" value=" ��ǰ ���� " class="button" onclick="jsGoedit('<%=idx%>');"/></div><% End If %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<% If number = "0" Then %>
	<td>�̹��� ����</td>
	<% End If %>
	<td>��ǰ�ڵ�</td>
	<td>�̹���</td>
	<td>��ǰ��</td>
</tr>
<%
	Dim addquery
	If number <> "0" Then
		addquery = " and S.viewidx = "& number
	End If 
	vCount = 0
	vQuery = "select S.showitemidx, S.showidx , S.itemid, i.itemname, i.smallimage , S.sortnum from db_sitemaster.dbo.tbl_mobile_showbanner_subitem as S " & _
			 "	left outer join db_item.dbo.tbl_item as i on S.itemid = i.itemid " & _
			 " where S.showidx = '" & idx & "'"& addquery & " and S.isusing = 'Y' order by S.sortnum asc"
	rsget.Open vQuery, dbget, 1
	If rsget.Eof Then
		Response.Write "<tr><td bgcolor='#FFFFFF' colspan='10' align='center'>�����Ͱ� �����ϴ�.</td></tr>"
	Else
		Do Until rsget.Eof
%>
		<tr>
			<% If number = "0" Then %>
			<td align="center" bgcolor="#FFFFFF"><%=rsget("sortnum")%>���̹���</td>
			<% End If %>
			<td align="center" bgcolor="#FFFFFF"><%=rsget("itemid")%></td>
			<td align="center" bgcolor="#FFFFFF"><img src="http://webimage.10x10.co.kr/image/small/<%=GetImageSubFolderByItemid(rsget("itemid"))%>/<%=rsget("smallimage")%>"></td>
			<td bgcolor="#FFFFFF"><%=rsget("itemname")%></td>
		</tr>
<%
		vCount = vCount + 1
		rsget.MoveNext
		Loop
	End IF
	
	rsget.close()
%>
</table>
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->