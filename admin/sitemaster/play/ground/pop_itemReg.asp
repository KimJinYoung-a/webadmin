<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim idx, vQuery, i, vCount
idx  = requestCheckVar(request("idx"),10)

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

<center><b>Idx: <%=idx%></b> ��ǰ ���/Ȯ��</center>
<form name="frm1" action="itemProc.asp" method="post" style="margin:0px;">
<input type="hidden" name="mode" value="insert">
<input type="hidden" name="idx" value="<%=idx%>">
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

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><!--<input type="checkbox" name="chkAll" onClick="jsChkAll();">--></td>
	<td>��ǰ�ڵ�</td>
	<td>�̹���</td>
	<td>��ǰ��</td>
</tr>
<%
	vCount = 0
	vQuery = "select S.itemidx, S.subidx , S.itemid, i.itemname, i.smallimage from db_sitemaster.dbo.tbl_play_ground_item as S " & _
			 "	left outer join db_item.dbo.tbl_item as i on S.itemid = i.itemid " & _
			 " where S.subidx = '" & idx & "' order by S.itemid asc"
	rsget.Open vQuery, dbget, 1
	If rsget.Eof Then
		Response.Write "<tr><td bgcolor='#FFFFFF' colspan='10' align='center'>�����Ͱ� �����ϴ�.</td></tr>"
	Else
		Do Until rsget.Eof
%>
		<tr>
			<td align="center" bgcolor="#FFFFFF"><input type="checkbox" name="itemid" value="<%=rsget("itemid")%>"></td>
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
<div>
	<div style="float:left"><input type="button" value="���û�ǰ����" onClick="goproddel()"></div>
	<div style="float:right"><input type="button" value="�ݱ�" onClick="window.close()"></div>
</div>
<script type="text/javascript">
function checkboxCheck()
{
	var j = document.getElementsByName("itemid").length;
	var k = new Array();
	var m = 0;

	for(var i=0; i < <%=CHKIIF(vCount=1,"1","j")%> ; i++){
	    if (document.getElementsByName("itemid")<%=CHKIIF(vCount=1,"","[i]")%>.checked == true)
	    {
	        k[m] = document.getElementsByName("itemid")<%=CHKIIF(vCount=1,"","[i]")%>.value;
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