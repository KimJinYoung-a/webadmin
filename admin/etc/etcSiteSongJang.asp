<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim sitename

sitename = RequestCheckVar(Request("sitename"),32)

%>
<script language='javascript'>
function MakeSongJangFile(frm){
    if (frm.sitename.value.length<1){
        alert('제휴사를 선택하세요.');
        frm.sitename.focus();
        return;
    }
    
    if ((frm.sitename.value=="gseshop")){  //(frm.sitename.value=="interpark")||
        alert(frm.sitename.value + '는 현재 지원되지 않습니다.');
        frm.sitename.focus();
        return;
    }
    
    if (frm.orgFile.value.length<1){
        alert('송장생성 원본 파일을 넣어 주세요.');
        frm.orgFile.focus();
        return;
    }
    
    frm.target="isongjangFrm";
    frm.submit();
}
</script>

<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#EEEEEE" class="a">
	<form name="frm" method="post" action="iFrameEtcSiteSongjang.asp">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr >
	    <td width="100">제휴사 선택</td>
		<td >
    		<select name="sitename" >
    		<option value="dnshop" <%= chkIIF(sitename="dnshop","selected","") %> >DnShop
    		<option value="interpark" <%= chkIIF(sitename="interpark","selected","") %> >InterPark
    		<option value="gseshop" <%= chkIIF(sitename="gseshop","selected","") %> >GsEshop
    		</select>
		</td>
		<td class="a" align="right">
		</td>
	</tr>
	<tr>
	    <td width="100">송장 입력폼</td>
	    <td>
	    <textarea name="orgFile" cols="80" rows="10"></textarea>    
	    <input type="button" value="Clear" onClick="frm.orgFile.value='';">
	    </td>
	</tr>
	<tr>
	    <td width="100">송장 입력폼</td>
	    <td><input type="button" value="송장입력파일 생성" onClick="MakeSongJangFile(frm);"></td>
	</tr>
	</form>
</table>
<iframe name="isongjangFrm" id="isongjangFrm" width="800" height="100"></iframe>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->