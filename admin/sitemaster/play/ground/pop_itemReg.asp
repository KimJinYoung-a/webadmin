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
	Response.Write "<script>alert('잘못된 경로입니다.\nNo. 번호가 있어야 합니다.');</script>"
	dbget.close()
	Response.End
END IF	
IF IsNumeric(idx) = False THEN
	Response.Write "<script>alert('잘못된 경로입니다.\nNo. 번호가 있어야 합니다.');</script>"
	dbget.close()
	Response.End
END IF
%>

<center><b>Idx: <%=idx%></b> 상품 등록/확인</center>
<form name="frm1" action="itemProc.asp" method="post" style="margin:0px;">
<input type="hidden" name="mode" value="insert">
<input type="hidden" name="idx" value="<%=idx%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr>
	<td align="right" bgcolor="#FFFFFF" style="padding:10px 10px 10px 0;">
		상품코드 : <input type="text" name="itemid" value="" size="53">
		<br>
		<b>※ 숫자로만 입력하고 두개 이상일 경우는 쉼표(,)로 구분해서 입력.</b>
		<input type="submit" class="button" value=" 저  장 ">
	</td>
</tr>
</table>
</form>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" valign="top" border="0">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><!--<input type="checkbox" name="chkAll" onClick="jsChkAll();">--></td>
	<td>상품코드</td>
	<td>이미지</td>
	<td>상품명</td>
</tr>
<%
	vCount = 0
	vQuery = "select S.itemidx, S.subidx , S.itemid, i.itemname, i.smallimage from db_sitemaster.dbo.tbl_play_ground_item as S " & _
			 "	left outer join db_item.dbo.tbl_item as i on S.itemid = i.itemid " & _
			 " where S.subidx = '" & idx & "' order by S.itemid asc"
	rsget.Open vQuery, dbget, 1
	If rsget.Eof Then
		Response.Write "<tr><td bgcolor='#FFFFFF' colspan='10' align='center'>데이터가 없습니다.</td></tr>"
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
	<div style="float:left"><input type="button" value="선택상품삭제" onClick="goproddel()"></div>
	<div style="float:right"><input type="button" value="닫기" onClick="window.close()"></div>
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
		alert("상품을 선택해 주세요.");
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