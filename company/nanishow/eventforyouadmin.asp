<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/company/menu/nanishowhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim filename,fso,ofile, buf
filename = "C:/home/cube1010/www" & "/ext/nanishow/notice.inc"
set fso = server.CreateObject("scripting.filesystemobject")
set ofile = fso.OpenTextFile (filename, 1, false )

buf = ofile.readAll

ofile.close
set ofile = nothing
set fso = nothing
%>
<script language='javascript'>
function SvNoticeConfirm(frm){
	var ret = confirm('저장 하시겠습니까?');
	if (ret) {
		frm.submit();
	}
}
</script>
	<form name="frm_notice" method="post" action="/lib/donotice.asp" onsubmit="SvNoticeConfirm(this); return false;">
	<input type="hidden" name="pid" value="<%= session("btcid") %>">
	<table border="1" width="700" bordercolorlight="#808080" cellspacing="0" bordercolordark="#FFFFFF" valign="top">
	  <tr>
	    <td class="a" colspan="5">EventForyou</td>
	  </tr>
	  <tr>
	    <td>
	    	
	    </td>
	  </tr>
	</form>  
	</table>

<!-- #include virtual="/company/lib/companybodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->