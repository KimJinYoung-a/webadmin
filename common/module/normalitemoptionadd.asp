<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<body onload="window.resizeTo(560,500);">
<table width="500" border="0" cellspacing="1" cellpadding="2" align="center" class="a"  bgcolor="#3d3d3d">
<form name="itemopt">
	<tr height="30" bgcolor="#DDDDFF" align="center">
		<td>�ɼ� ����</td>
		<td>�ɼ� ��</td>
	</tr>
	<tr bgcolor="#FFFFFF" align="center">
		<td>
		  <select name="opt1" size="20" style='width:240;' onchange="javascript:searchOption(this.options[this.selectedIndex].value);" >
		  <option value="">-----------------------</option>
		  </select>
		</td>
		<td>
		  <select multiple name="opt2" size="20" style='width:240;'>
		  <option value="">-----------------------</option>
		  </select>&nbsp;
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="4" align="center">
			<input type="button" value="���ÿɼ��߰�" onclick="MoveOptionWithGubun(document.itemopt.elements['opt1'],document.itemopt.elements['opt2'])">
			<input type="button" value=" �� �� " onclick="self.close()">
		</td>
	</tr>
</form>
</table>
<iframe name="FrameSearchOption" src="/lib/frame_option_select.asp?form_name=itemopt&element_name=opt1" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
<script language="JavaScript">
<!--

//�ɼ��������ý� �����ɼ� ����
function searchOption(paramCode1) {

	resetOption1() ;
	//resetRealOption() ;

	if(paramCode1 != '') {
		FrameSearchOption.location.href="/lib/frame_option_select.asp?search_code=" + paramCode1 + "&form_name=itemopt&element_name=opt2";
	}
}

//�ɼǸ���Ʈ �ʱ�ȭ
function resetOption1() {
	document.itemopt.opt2.length = 1;
	document.itemopt.opt2.selectedIndex = 1 ;
}

//���ÿɼ� �ʱ�ȭ
function resetRealOption() {
	opener.document.itemreg.realopt.length = 0;
	opener.document.itemreg.realopt.selectedIndex = 0 ;
}

function MoveOption(fbox) {
	for(i=0; i<fbox.options.length; i++){
		if(fbox.options[i].selected){
			opener.InsertOption(fbox.options[i].text, fbox.options[i].value)
			fbox.options[i] = null;
			i=i-1;
		}
	}
}

function MoveOptionWithGubun(fbox1,fbox2) {
    var optTypeName = "";
    
	
    for(i=0; i<fbox1.options.length; i++){
        if(fbox1.options[i].selected){
            optTypeName = fbox1.options[i].text;
        }
    }
    
    
    optTypeName = optTypeName.replace(/\(�ѱ�\)/gi,'');
	optTypeName = optTypeName.replace(/\(����\)/gi,'');
	optTypeName = optTypeName.replace(/\(1-99\)/gi,'');
	optTypeName = optTypeName.replace(/����Ŭ��2/gi,'����Ŭ��');
	
	
	for(i=0; i<fbox2.options.length; i++){
		if(fbox2.options[i].selected){
			opener.InsertOptionWithGubun(optTypeName , fbox2.options[i].text, fbox2.options[i].value)
			fbox2.options[i] = null;
			i=i-1;
		}
	}
}
//-->
</script>
</body>
<!-- #include virtual="/admin/lib/poptail.asp"-->
