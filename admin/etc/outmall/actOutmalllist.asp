<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/etc/outmallConfirmCls.asp"-->
<%
Response.CharSet = "euc-kr"
Dim midx, i, arrRows, oOutMall, vidx, cmdparam
Dim vSql
midx		= request("idx")
vidx		= request("vidx")
cmdparam	= request("cmdparam")

If cmdparam = "I" Then
	Dim mallarr, splitMall, lp
	mallarr =  request("chk_"&vidx)
	splitMall = Split(mallarr, ",")

	vSql = ""
	vSql = vSql & " DELETE FROM db_etcmall.[dbo].[tbl_outmall_not_in_keywords_mallid] WHERE midx = '"& vidx &"' "
	dbget.execute vSql

	For lp = 0 To Ubound(splitMall)
		vSql = ""
		vSql = vSql & " INSERT INTO db_etcmall.[dbo].[tbl_outmall_not_in_keywords_mallid] (midx, mallid) values ('"& vidx &"', '"& Trim(splitMall(lp)) &"') "
		dbget.execute vSql
	Next
	Response.Write "<script>parent.location.reload();</script>"
	Response.End
End If

SET oOutMall = new cOutmall
	oOutMall.FRectIdx = midx
	arrRows = oOutMall.fnOutmallList2
SET oOutMall = nothing
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function ckAll(v){
	if($("#cbx_chkAll_"+v).is(":checked")){
		$("input[name=chk_"+v+"]").prop("checked", true);
	}else{
		$("input[name=chk_"+v+"]").prop("checked", false);
	}
}
function fnkeywordMallProc(f){
	f.target = "xLink";
	f.submit();
}
</script>
<form name="frmk" method="post" action="actOutmalllist.asp" onSubmit="return false;" action="" style="margin:0px;">
<input type="hidden" id="vidx" name="vidx" value="<%= midx %>">
<input type="hidden" name="cmdparam" value="I">
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="0" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#FFFFFF">
    <td align="left">
	<input type="checkbox" id="cbx_chkAll_<%=midx%>" onclick="ckAll('<%= midx %>')" />전체<br />
	<%
		Dim chkmallid
		For i = 0 To Ubound(arrRows, 2)
			chkmallid = ""
			If lcase(arrRows(0, i)) = lcase(arrRows(1, i)) Then
				chkmallid = "Y"
			End If
	%>
			<label><input type="checkbox" class="checkbox" id="chk_<%=midx%>" name="chk_<%=midx%>" value="<%= arrRows(0, i) %>" <%= Chkiif(chkmallid="Y", "checked", "") %> ><%= arrRows(0, i) %></label><br />
	<%
		Next
	%>
	</td>
</tr>
<tr align="center" height="25" bgcolor="#FFFFFF">
    <td>
		<input type="button" class="button" value="저장" onclick="fnkeywordMallProc(this.form);" style=color:blue;font-weight:bold>
		<input type="button" class="button" value="닫기" onclick="ajaxOutmall222('<%=midx%>');" style=color:black;font-weight:bold>
	</td>
</tr>
</table>
</form>
<!-- #include virtual="/lib/db/dbclose.asp" -->
