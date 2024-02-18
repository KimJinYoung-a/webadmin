<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/company/incSessionCompany.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/company/menu/nanishowhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/extstylecls.asp"-->

<%
dim filename,fso,ofile, buf
filename = "C:/home/cube1010/www" & "/ext/nanishow/notice.inc"
set fso = server.CreateObject("scripting.filesystemobject")
set ofile = fso.OpenTextFile (filename, 1, false )

buf = ofile.readAll

ofile.close
set ofile = nothing
set fso = nothing

dim ostyle,i
set ostyle = new CExtSpecial
ostyle.GetSpecialData "nanistyle",0

%>
<script language='javascript'>
function SvNoticeConfirm(frm){
	var ret = confirm('저장 하시겠습니까?');
	if (ret) {
		frm.submit();
	}
}

function ShowItemList(frm){
	var paramstr ='';
	paramstr = '?frmname=' + frm.name + '&bxitemid=iitemid&bxitemname=idisptext&bxitemimage=Img_' + frm.iitemid.value;
	var newwin = window.open('/company/lib/popitemlist.asp' + paramstr, 'popitemlist', 'width=660,height=600,menubar=no,scrollbars=yes');
	
}

function saveonebest(frm){
	if (frm.idispnum.length<2) {
		alert('표시순서는 03 형태로 입력하시기 바랍니다.');
		return;
	}
	
	var ret = confirm('저장하시겠습니까?');
	if (ret) {
		frm.submit();
	}
}

</script>

<table border="0" width="700" bordercolorlight="#808080" cellspacing="0" bordercolordark="#FFFFFF" valign="top">
  <tr>
    <td class="a" colspan="5">BestNaniShow</td>
  </tr>
  <tr>
    <td>
    	<table width="700" border="1" cellspacing="0" cellpaddin="0" class="a">
    	<tr>
    		<td width="50">표시순서</td>
    		<td width="100">ItemID</td>
    		<td width="100">Image</td>
    		<td width="100">표시명</td>
    		<td width="100">저장</td>
    		<td></td>
    	</tr>
    	<% for i=0 to ostyle.FResultCount-1 %>
    	<form name="frm_<%= ostyle.FSpecialList(i).FItemID %>" method="post" action="/company/nanishow/dobestnani.asp" >
    	<tr>
    		
    		<td><input type="textbox" name="idispnum" value="<%= ostyle.FSpecialList(i).FDispNum %>" size="5" maxlength="2"></td>
    		<td>
    			<input type="hidden" name="olditemid" value="<%= ostyle.FSpecialList(i).FItemID %>">
    			<input type="textbox" name="iitemid" value="<%= ostyle.FSpecialList(i).FItemID %>" size="5">
    			<input type="button" value="선택" onClick="ShowItemList(frm_<%= ostyle.FSpecialList(i).FItemID %>)">
    		</td>
    		<td><img name="Img_<%= ostyle.FSpecialList(i).FItemID %>" src="<%= ostyle.FSpecialList(i).FImageList %>"></td>
    		<td><input type="textbox" name="idisptext" value="<%= ostyle.FSpecialList(i).getDispTitleorItemName %>" size="20"></td>
    		<td><input type="button" value="저장" onclick="saveonebest(frm_<%= ostyle.FSpecialList(i).FItemID %>)"></td>
    		<td></td>
    	</tr>
    	</form>
    	<% next %>
    	</table>
    </td>
  </tr>
</table>

<%
set ostyle = Nothing
%>
<!-- #include virtual="/company/lib/companybodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->