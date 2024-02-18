<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  기프트플러스
' History : 2010.04.02 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/giftplus/giftplus_cls.asp"-->
<% 
dim Frmvalue,imageName
	Frmvalue= request("Frmvalue")
	imageName = request("imageName")
%>

<body leftmargin="0">

<script language="javascript">

function subchk(){
	if (isNaN(UpdateFRM.viewidx.value)) {
		alert('숫자만 입력가능합니다');
		return false;
	}
}

function popInputImg(){
	var pop = window.open('','','resizable=yes');
}

</script>

<table width="290" border="0" cellpadding="2" cellspacing="1" class="a"  align="center" bgcolor="<%= adminColor("tablebg") %>">
	<form name="upimgfrm" method="post" action="<%= uploadUrl %>/linkweb/giftplus/menuimage_process.asp" enctype="MULTIPART/FORM-DATA">
	<input type="hidden" name="Frmvalue" value="<%= Frmvalue %>">
	<input type="hidden" name="imageName" value="<%= imageName %>">
	<tr>
		<td bgcolor="#FFFFFF" align="center"><input type="file" name="imagefile" class="button" size="26"></td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF" align="center"><input type="submit" class="button" value="저장"></td>
	</tr>
	</form>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->