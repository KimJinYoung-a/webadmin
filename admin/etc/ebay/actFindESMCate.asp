<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/ebay/ebayCls.asp"-->
<%
Response.CharSet = "euc-kr"
Dim oEbay, i, page, srcKwd, tmpKey
page		= requestCheckVar(request("page"),10)
srcKwd		= Trim(requestCheckVar(request("srcKwd"),50))

If page = ""	Then page = 1
'// ��� ����
Set oEbay = new CEbay
	oEbay.FPageSize = 1000
	oEbay.FCurrPage = page
	oEbay.FRectSearchName = srcKwd
	oEbay.getESMCateList
%>
<script>
function chkThis(comp){
    //AnCheckClick(comp);
}

function fnChkThisCate(ii,stdcate,dispcate){
    var iobj;
    if (document.resultFrm.chk.length){
        iobj = document.resultFrm.chk[ii];
    }else{
        iobj = document.resultFrm.chk
    }
    var pchecked = iobj.checked;
    iobj.checked = !pchecked;

}
</script>
<p>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<tr height="5" valign="top">
	<td align="right"><img src="/images/icon_arrow_down.gif" border="0" vspace="5" align="absmiddle"> �˻���� : <strong><%=oEbay.FtotalCount%></strong>&nbsp;&nbsp;</td>
</tr>
</table>
<form name="resultFrm" >
<input type="hidden" name="cdl" value="">
<input type="hidden" name="cdm" value="">
<input type="hidden" name="cds" value="">
<input type="hidden" name="mode" value="saveCateArr">

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#DDDDFF">
    <td></td>
	<td>ESMī�װ�</td>
	<td>ESMī�װ���</td>
	<td>���޸�</td>
	<td>Siteī�װ�</td>
	<td>Siteī�װ���</td>
</tr>
<% If oEbay.FresultCount < 1 Then %>
<tr bgcolor="#FFFFFF">
	<td colspan="8" height="40" align="center">[�˻������ �����ϴ�.]</td>
</tr>
<%
	Else
		For i = 0 to oEbay.FresultCount - 1
			If i <> 0 AND tmpKey <> oEbay.FItemList(i).FSDCategoryCode Then
%>
<tr align="center" height="25"  title="ī�װ� ����" bgcolor="#999999">
	<td colspan="6"></td>
</tr>
<%
			End If
%>
<tr align="center" height="25"  title="ī�װ� ����" bgcolor="#FFFFFF">
	<td>
	    <input type="checkbox" name="chk" id="chk" value="<%=i%>" onClcik="chkThis(this)";>
	    <input type="hidden" name="cateCode" value="<%= oEbay.FItemList(i).FCateCode %>">
	    <input type="hidden" name="stdcode" value="<%= oEbay.FItemList(i).FSDCategoryCode %>">
	    <input type="hidden" name="gubun" value="<%= oEbay.FItemList(i).FGubun %>">
	</td>
	<td><%= oEbay.FItemList(i).FSDCategoryCode %></td>
	<td><%= oEbay.FItemList(i).FSDCategoryName %></td>
	<td>
		<%
			Select Case oEbay.FItemList(i).FGubun
				Case "A"	response.write "<font color='RED'>����</font>"
				Case "G"	response.write "<font color='GREEN'>������</font>"
			End Select
		%>
	</td>
	<td align="left"><%= oEbay.FItemList(i).FCateCode %></td>
	<td><%= oEbay.FItemList(i).FCateName %></td>
</tr>
<%
			tmpKey = oEbay.FItemList(i).FSDCategoryCode
		Next
	End If
%>
</table>
</form>
<% Set oEbay = Nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
