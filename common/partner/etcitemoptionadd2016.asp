<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/partner/lib/adminHead.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'###########################################################
' Description : 상품 전용옵션 등록
' History : 2013.12.16 정윤정 옵션갯수 수정   
'###########################################################
%>
<%
dim i,iRowMax
iRowMax = 19 '옵션 최대갯수
%>

<script type/text="javascript">
<!--
function AddOption()
{
	var frm = document.itemopt;
    var addedCnt = 0;
    
	if(!frm.optTypeNm.value){
		alert("추가할 옵션 구분 명을 입력해주십시오.");
		frm.optTypeNm.focus();
		return false;
	}
	
	if(GetByteLength(frm.optTypeNm.value)>32){
		alert("옵션구분명은 32byte (한글 16자, 영문 32자) 이내로 입력해주세요"); 
		frm.optTypeNm.focus();
		return false;
	}

    for (var i=0;i<frm.optNm.length;i++){ 
        if(GetByteLength(frm.optNm[i].value) >32 ){
        	alert("옵션명은 32byte (한글 16자, 영문 32자) 이내로 입력해주세요");
        	frm.optNm[i].focus(); 
        	return false;
        }
     }
      for (var i=0;i<frm.optNm.length;i++){    
         if (frm.optNm[i].value.length>0){
            opener.InsertOptionWithGubun(frm.optTypeNm.value, frm.optNm[i].value, "0000");
            addedCnt++;
        }
    }

    if (addedCnt>0){
	    self.close();
	}else{
	    alert('추가할 옵션을 입력해 주세요.');
	}
}
//-->
</script>
<body onload="document.itemopt.optTypeNm.focus();">
<div class="popupWrap">
		<div class="popHead">
		<h1><img src="/images/partner/pop_admin_logo.gif" alt="10x10" /></h1>
		<p class="btnClose"><input type="image" src="/images/partner/pop_admin_btn_close.gif" alt="창닫기" onclick="window.close();" /></p>
	</div>	
	<form name="itemopt" >
	<div class="cont">
	<table class="tbType1 writeTb tMar10">  
    <tr>
		<th>옵션 구분 명</th>
		<td><input type="text" name="optTypeNm" size="20" maxlength="32"> 색상</td>
	</tr>
	<% for i=0 to iRowMax %>
	<tr>
		<th>옵션 명 <%= i+1 %></td>
		<td bgcolor="#FFFFFF" align="left"><input type="text" name="optNm" size="32" maxlength="32"> <%= chkIIF(i=0,"빨강","") %><%= chkIIF(i=1,"파랑","") %><%= chkIIF(i=2,"노랑","") %></td>
	</tr>
	<% next %> 
</table>
</div>
<div class="tPad15 ct"> 
			<input type="button" class="btn3 btnDkGy" value="옵션추가" onclick="AddOption();"> 
</div>

</div>
</form> 
 
</body>
</html>