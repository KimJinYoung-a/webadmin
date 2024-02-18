<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 전시 카테고리 선택
' History : 2018/01/18 eastone 전시카테고리 팝업이 없어서 만들었음.
'####################################################
%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim dispCate
dim frmname, targetcompname, targetcpndtlnm
dispCate = requestCheckvar(request("dispCate"),32)
frmname  = requestCheckvar(request("frmname"),64)
targetcompname  = requestCheckvar(request("targetcompname"),64)
targetcpndtlnm  = requestCheckvar(request("targetcpndtlnm"),64)
%>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
function finselect(){
    var selcate = document.frmcate.disp.value;
    if (selcate==""){
        alert('전시카테고리를 선택해주세요.');
        return;    
    }
    
    var catename="";
    var depth = (selcate.length)/3;
    $(".formSlt").each(function(i) {
        if (i<=depth-1){
            catename = catename + $(this).find("option:selected").text() ;
            if (i<depth-1){
                catename = catename + '>';
            }
        }
        //alert("NAME :" + $(this).find("option:selected").text() +"\n"+ "VALUE :" +$(this).find("option:selected").val()) ;
    })
    var parentcomp = opener.<%=frmname%>.<%=targetcompname%>;
    parentcomp.value=selcate;
    
    var parentcompdesc = opener.<%=frmname%>.<%=targetcpndtlnm%>;
    if (parentcompdesc){
        parentcompdesc.value=catename;
    }
    window.close();
}    
</script>
<body bgcolor="#F4F4F4">
<!-- 해더 -->
<table width="700" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="1" cellspacing="0" class="a" bgcolor="#999999">
		<tr>
			<td width="400" style="padding:5; border-top:1px solid #999999;border-left:1px solid #999999;border-right:1px solid #999999"  background="/images/menubar_1px.gif">
				<font color="#333333"><b>전시 카테고리 선택</b></font>
			</td>
			<td align="right" style="border-bottom:1px solid #999999" bgcolor="#F4F4F4">&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<table width="630" border="0" cellspacing="3" cellpadding="0" align="center">
<form name="frmcate">
<tr><td colspan="2" height="5"></td></tr>
<tr>
	<td colspan="2" align="center">
		<!-- #include virtual="/common/module/dispCateSelectBox.asp" -->
	</td>
</tr>
<tr><td colspan="2" height="5"></td></tr>
<tr>
	<td align="center">
	    <input type="button" class="button" value="선택완료" onclick="finselect()">
	    &nbsp;
		<input type="button" class="button" value="창닫기" onclick="self.close()">
	</td>
</tr>
</form>
</table>
</body>
<!-- #include virtual="/common/lib/poptail.asp"-->