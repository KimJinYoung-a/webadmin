<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylelifeCls.asp"-->

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

<% If number <>"0" Then %>
<center><b>Idx: <%=idx%> No.<%=number%></b> ��ǰ ���</center>
<form name="frm1" action="itemProc.asp" method="post" style="margin:0px;">
<input type="hidden" name="mode" value="insert">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="number" value="<%=number%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr>
	<td align="right" bgcolor="#FFFFFF" style="padding:10px 10px 10px 0;">
		��ǰ�ڵ� : <input type="text" name="itemid" value="" size="53">
		<br>
		<b>�� ���ڷθ� �Է��ϰ� �ΰ� �̻��� ���� ��ǥ(,)�� �����ؼ� �Է�.</b>
		<input type="submit" class="button" value=" ��  �� ">
	</td>
</tr>
</table>
</form>
<% Else %>
<center><b>Idx: <%=idx%> </b> ��ǰ Ȯ��</center>
<% End If %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><!--<input type="checkbox" name="chkAll" onClick="jsChkAll();">--></td>
	<% If number = "0" Then %>
	<td>�̹��� ��Ī No</td>
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
	vQuery = "select S.styleitemidx, S.styleidx , S.itemid, i.itemname, i.smallimage , S.viewidx from db_sitemaster.dbo.tbl_play_style_item as S " & _
			 "	left outer join db_item.dbo.tbl_item as i on S.itemid = i.itemid " & _
			 " where S.styleidx = '" & idx & "'"& addquery & " order by S.itemid asc"
	rsget.Open vQuery, dbget, 1
	If rsget.Eof Then
		Response.Write "<tr><td bgcolor='#FFFFFF' colspan='10' align='center'>�����Ͱ� �����ϴ�.</td></tr>"
	Else
		Do Until rsget.Eof
%>
		<tr>
			<td align="center" bgcolor="#FFFFFF"><input type="checkbox" name="itemids" value="<%=rsget("itemid")%>"></td>
			<% If number = "0" Then %>
			<td align="center" bgcolor="#FFFFFF"><%=rsget("viewidx")%>���̹���</td>
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
<br>
<% If number <>"0" Then %>
<input type="button" value="���û�ǰ����" onClick="goproddel()">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<% End If %>
<input type="button" value="�ݱ�" onClick="window.close()">
<script type="text/javascript">
function checkboxCheck()
{
	var j = document.getElementsByName("itemids").length;
	var k = new Array();
	var m = 0;

	for(var i=0; i < <%=CHKIIF(vCount=1,"1","j")%> ; i++){
	    if (document.getElementsByName("itemids")<%=CHKIIF(vCount=1,"","[i]")%>.checked == true)
	    {
	        k[m] = document.getElementsByName("itemids")<%=CHKIIF(vCount=1,"","[i]")%>.value;
	        m = m+1;
	    }
	}
	return k;
}
function goproddel()
{
	var i = checkboxCheck();
	if(i == "")
	{
		alert("��ǰ�� ������ �ּ���.");
		return;
	}
	else
	{
		document.frm1.itemid.value = i;
		document.frm1.mode.value = "delete";
		document.frm1.submit();
	}
}
</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->