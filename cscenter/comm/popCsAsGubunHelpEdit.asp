<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/cscenter/CsCommCdcls.asp"-->

<%
dim comm_cd
comm_cd = requestCheckVar(request("comm_cd"),32)

dim oCommHelp
set oCommHelp = new CCommCd
oCommHelp.FRectCommCd = comm_cd

oCommHelp.GetCommHelpStr

dim i, infoHtml_B001, infoHtml_B007

for i=0 to oCommHelp.FResultCount-1
    if (oCommHelp.FItemList(i).Fstate_comm_cd="B001") then
        infoHtml_B001 = oCommHelp.FItemList(i).FinfoHtml
    end if 
    
    if (oCommHelp.FItemList(i).Fstate_comm_cd="B007") then
        infoHtml_B007 = oCommHelp.FItemList(i).FinfoHtml
    end if 
next

%>
<script language='javascript'>
function EditCsAsGubunHelp(frm){
    if (confirm('저장 하시겠습니까?')){
        frm.submit();
    }
}
</script>
<p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm_search" method="GET" >
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			코드:
			<input type="text" class="text" name="comm_cd" size="20" value="<%= comm_cd %>" maxlength=4>
			
			<% if (oCommHelp.FREsultCount>0) then %>
			   <%= oCommHelp.FItemList(0).Fdiv_comm_name %>
			<% else %>
			    <font color="red">코드오류</font>
			<% end if %>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm_search.submit();">
		</td>
	</tr>
	</form>
</table>
<p>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frmEdit" method="post" action="popCsAsGubunHelpEdit_Process.asp">
    <input type="hidden" name="comm_cd" value="<%= comm_cd %>">
	<tr height="25" bgcolor="FFFFFF">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>" rowspan="2">접수</td>
		<td ><textarea name="infoHtml_B001" cols="70" rows="10"><%= infoHtml_B001 %></textarea></td>
	</tr>
	<tr height="60" bgcolor="FFFFFF">
	    <td ><%= infoHtml_B001 %></td>
	</tr>
	<tr height="25" bgcolor="FFFFFF">
		<td width="100" bgcolor="<%= adminColor("tabletop") %>" rowspan="2">완료</td>
		<td ><textarea name="infoHtml_B007" cols="70" rows="10"><%= infoHtml_B007 %></textarea></td>
	</tr>
	<tr height="60" bgcolor="FFFFFF">
	    <td ><%= infoHtml_B007 %></td>
	</tr>
	<tr bgcolor="FFFFFF">
	    <td colspan="2" align="center"><input type="button" value=" 저 장 " <%= ChkIIF(oCommHelp.FREsultCount>0,"","disabled") %> onClick="EditCsAsGubunHelp(frmEdit);"></td>
	</tr>
	
	</form>
</table>
<%
set oCommHelp = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->