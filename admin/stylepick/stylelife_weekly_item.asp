<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/stylepick/stylelifeCls.asp"-->

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

<center><b>No. <%=idx%></b> ��ǰ ���</center>
<form name="frm1" action="stylelife_weekly_item_proc.asp" method="post" style="margin:0px;">
<input type="hidden" name="action" value="insert">
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="imgsize" value="">
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
	<td>�̹���ũ��</td>
</tr>
<%
	vCount = 0
	vQuery = "select w.itemidx, w.itemid, i.itemname, i.smallimage, w.imgsize from db_giftplus.dbo.tbl_stylelife_weekly_item as w " & _
			 "	inner join db_item.dbo.tbl_item as i on w.itemid = i.itemid " & _
			 " where w.idx = '" & idx & "' order by w.imgsize desc, w.itemidx asc"
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
			<td align="center" bgcolor="#FFFFFF">
				<select name="selimgsize" class="select" onchange="changesize(this.value,<%=rsget("itemid")%>)">	
				<option value="100" <%=CHKIIF(CInt(rsget("imgsize"))=100,"selected","")%>>100px</option>
				<option value="200" <%=CHKIIF(CInt(rsget("imgsize"))=200,"selected","")%>>200px</option>
				</select>
			</td>
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
<input type="button" value="���û�ǰ����" onClick="goproddel()">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="button" value="�ݱ�" onClick="window.close()">

<script language="javascript">
// ��ü �̹���ũ�� �ϰ� ��ȯ
function jsSizeChg(selv) {
    var frm, blnChk;
	frm = document.fitem;
	if (frm.chkI.length > 1){
		for (var i=0;i<frm.itemimgsize.length;i++){
			frm.itemimgsize[i].value=selv;
		}
	} else {
		frm.itemimgsize.value=selv;
	}
}
function changesize(a,b)
{
	document.frm1.imgsize.value = a;
	document.frm1.itemid.value = b;
	document.frm1.action.value = "update";
	document.frm1.submit();
}
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
		document.frm1.action.value = "delete";
		document.frm1.submit();
	}
}
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->