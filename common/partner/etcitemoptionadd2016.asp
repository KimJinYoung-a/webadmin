<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/partner/lib/adminHead.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
'###########################################################
' Description : ��ǰ ����ɼ� ���
' History : 2013.12.16 ������ �ɼǰ��� ����   
'###########################################################
%>
<%
dim i,iRowMax
iRowMax = 19 '�ɼ� �ִ밹��
%>

<script type/text="javascript">
<!--
function AddOption()
{
	var frm = document.itemopt;
    var addedCnt = 0;
    
	if(!frm.optTypeNm.value){
		alert("�߰��� �ɼ� ���� ���� �Է����ֽʽÿ�.");
		frm.optTypeNm.focus();
		return false;
	}
	
	if(GetByteLength(frm.optTypeNm.value)>32){
		alert("�ɼǱ��и��� 32byte (�ѱ� 16��, ���� 32��) �̳��� �Է����ּ���"); 
		frm.optTypeNm.focus();
		return false;
	}

    for (var i=0;i<frm.optNm.length;i++){ 
        if(GetByteLength(frm.optNm[i].value) >32 ){
        	alert("�ɼǸ��� 32byte (�ѱ� 16��, ���� 32��) �̳��� �Է����ּ���");
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
	    alert('�߰��� �ɼ��� �Է��� �ּ���.');
	}
}
//-->
</script>
<body onload="document.itemopt.optTypeNm.focus();">
<div class="popupWrap">
		<div class="popHead">
		<h1><img src="/images/partner/pop_admin_logo.gif" alt="10x10" /></h1>
		<p class="btnClose"><input type="image" src="/images/partner/pop_admin_btn_close.gif" alt="â�ݱ�" onclick="window.close();" /></p>
	</div>	
	<form name="itemopt" >
	<div class="cont">
	<table class="tbType1 writeTb tMar10">  
    <tr>
		<th>�ɼ� ���� ��</th>
		<td><input type="text" name="optTypeNm" size="20" maxlength="32"> ����</td>
	</tr>
	<% for i=0 to iRowMax %>
	<tr>
		<th>�ɼ� �� <%= i+1 %></td>
		<td bgcolor="#FFFFFF" align="left"><input type="text" name="optNm" size="32" maxlength="32"> <%= chkIIF(i=0,"����","") %><%= chkIIF(i=1,"�Ķ�","") %><%= chkIIF(i=2,"���","") %></td>
	</tr>
	<% next %> 
</table>
</div>
<div class="tPad15 ct"> 
			<input type="button" class="btn3 btnDkGy" value="�ɼ��߰�" onclick="AddOption();"> 
</div>

</div>
</form> 
 
</body>
</html>