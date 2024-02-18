<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/partner/lib/adminHead.asp" -->
<body  >
	<div class="popupWrap">
		<div class="popHead">
		<h1><img src="/images/partner/pop_admin_logo.gif" alt="10x10" /></h1>
		<p class="btnClose"><input type="image" src="/images/partner/pop_admin_btn_close.gif" alt="창닫기" onclick="window.close();" /></p>
	</div>
	<form name="itemopt">
		<div class="cont">
	<table class="tbType1 listTb tMar10"> 
	<tr>
		<th><div>옵션 구분</div></th>
		<th><div>옵션 명</div></th>
	</tr>
	<tr>
		<td>
		  <select class="formSlt" name="opt1" size="20" style='width:240;height:400px;' onchange="javascript:searchOption(this.options[this.selectedIndex].value);" >
		  <option value="">-----------------------</option>
		  </select>
		</td>
		<td>
		  <select class="formSlt" multiple name="opt2" size="20" style='width:240;height:400px;'>
		  <option value="">-----------------------</option>
		  </select>&nbsp;
		</td>
	</tr>
</table>	
</div>
<div class="tPad15 ct"> 
			<input type="button" class="btn3 btnDkGy" value="선택옵션추가" onclick="MoveOptionWithGubun(document.itemopt.elements['opt1'],document.itemopt.elements['opt2'])"> 
</div>

</div>
</form>
<iframe name="FrameSearchOption" src="/lib/frame_option_select.asp?form_name=itemopt&element_name=opt1" width="0" height="0" frameborder="0" hspace="0" vspace="0" scrolling="no"></iframe>
<script type="text/JavaScript">
<!--

//옵션종류선택시 개별옵션 셋팅
function searchOption(paramCode1) {

	resetOption1() ;
	//resetRealOption() ;

	if(paramCode1 != '') {
		FrameSearchOption.location.href="/lib/frame_option_select.asp?search_code=" + paramCode1 + "&form_name=itemopt&element_name=opt2";
	}
}

//옵션리스트 초기화
function resetOption1() {
	document.itemopt.opt2.length = 1;
	document.itemopt.opt2.selectedIndex = 1 ;
}

//선택옵션 초기화
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
    
    
    optTypeName = optTypeName.replace(/\(한글\)/gi,'');
	optTypeName = optTypeName.replace(/\(영문\)/gi,'');
	optTypeName = optTypeName.replace(/\(1-99\)/gi,'');
	optTypeName = optTypeName.replace(/프랭클린2/gi,'프랭클린');
	
	
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
</div>
</body>
</html>
