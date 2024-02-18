<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/sitemaster/EmoDic/EmoDicCls.asp" -->
<%

dim eNumber
eNumber = request("eno")
dim eType 
eType = request("etp")

IF eNumber="" or eType="" Then
	response.write "Àß¸øµÈ Á¢±ÙÀÔ´Ï´Ù"
	dbget.close()	:	response.End
End if


function getETypeStr(eTp)
	dim eStr 
	Select Case etp
	Case "1"
		eStr = "²ô´ö²ô´ö"	
	Case"2"
		eStr = "¾ó··¶×¶¥"	
	Case "3"
		eStr = "½Ì¼þ»ý¼þ"		
	Case "4"
		eStr = "³¢¸®³¢¸®"
	End Select
	getETypeStr=eStr
End Function
%>

<script language="javascript" type="text/javascript">
function fncgsel(){
	document.rFrm.submit();
}
window.resizeTo(620,370);
</script>

<table width="550" border="0" class="a" cellpadding="5" cellspacing="1" align="left" bgcolor="<%=adminColor("tablebg") %>">
<form name="" method="post" action="EmoDic_Proc.asp">
<input type="hidden" name="eno" value="<%= eNumber %>">
<input type="hidden" name="etp" value="<%= eType %>">
<input type="hidden" name="mode" value="batch">
<tr bgcolor="#FFFFFF">
	<td colspan="2" bgcolor="<%=adminColor("tablebar")%>">
		<b><%=eNumber %>Â÷ - <%=getETypeStr(eType) %></b> ÀÏ°ýµî·Ï
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF"></td>
	<td bgcolor="#FFFFFF" style="padding-left:10;">
		<textarea name="awrd" cols="70" rows="10" ></textarea>
		<br>
		<font size="2" color="red">
		´Ü¾î´Â ÄÞ¸¶(,)·Î ±¸ºÐÇØÁÖ¼¼¿ä</font>
	</td>
	
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan="2" align="right"><input type="submit" class="button" value="µî·Ï"></td>
</tr>
</form>
</table>		
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->