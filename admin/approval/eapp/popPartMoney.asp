<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ڱݰ��� �μ����
' History : 2011.03.22 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"--> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/approval/partMoneyCls.asp"-->
<%
Dim clsPart,istep1partidx, istep2partidx, iType
Dim arrList, intLoop , arrData, intD  

istep1partidx = requestCheckvar(Request("selp1"),10)
IF istep1partidx = "" THEN istep1partidx = 0
	
istep2partidx = requestCheckvar(Request("selp2"),10)
IF istep2partidx = "" THEN istep2partidx = 0
	
iType	= 	 requestCheckvar(Request("iType"),10)
%> 
 
<script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script>
<script language="javascript">
<!-- 
$(document).ready(function(){
	$("#selp1").change(function(){  
		var selp1 = $("#selp1").val();
	 
		 var url="/admin/approval/partMoney/ajaxPart.asp";
		 var params = "hidM=S&iDP=2&is1="+selp1;
		 var default_args = "<select name='selp3' id='selp3'>" 
						 + "<option value='0'>--����--</option>" 
						  + "</select>";
		 $.ajax({
		 	type:"POST",
		 	url:url,
		 	data:params,
		 	success:function(args){  
		 		$("#sp2").html(args);
		 		$("#sp3").html(default_args);	 
		 	}, 
		 	error:function(e){ 
		 		alert("�����ͷε��� ������ ������ϴ�. �ý������� �������ּ���");
		 		//alert(e.responseText);
		 	}
		 }); 
	}); 
});

$(function(){ 
	//�߰���ư Ŭ���� �̺�Ʈ
	$("#btnAdd").click(function(){  
	 
	 var sValue;
	 var sText;
	 
	if( $("#selp1 option:selected").val() == 0 || $("#selp2 option:selected").val() ==0  || $("#selp3 option:selected").val() ==0){
	alert("�߰��Ͻ� �μ��� �������ּ���");
	return;
	};  
	 
	sValue = $("#selp1 option:selected").val()+"^"+$("#selp2 option:selected").val()+"^"+$("#selp3 option:selected").val();  //���ð�
	sText  = $("#selp1 option:selected").text()+">"+$("#selp2 option:selected").text()+">"+$("#selp3 option:selected").text(); //���ð��� �ؽ�Ʈ
	  
	   
	for(j=0; j<$("#selU option").size();j++){	 
		if($("#selU option:eq("+j+")").val()==sValue){
			alert("�̹� ��ϵǾ��ֽ��ϴ�.");
			return;
		}
	}	 
	$("#selU").append("<option value='"+sValue+"'>"+sText+"</option>"); //�߰�ó��  
	});	
	
	//����
	$("#btnDel").click(function(){ 
		 $("#selU option:selected").remove(); 
	});	 			
});
    

//�μ� ���
function jsSubmitPartMoney(){   
	var strT  = "<table border=0 cellpadding=5 cellspacing=1 class=a bgcolor=#BABABA>"	;
	strT = strT + "<tr align='center' bgcolor=#E6E6E6><td  width=200>�μ�</td><td colspan=3 width=200>�ݾ�</td></tr>";
   for(i=0;i<document.frm.selU.length;i++){  
       if(i==0){
   		opener.document.frm.iP.value = document.frm.selU[i].value; 
   		opener.document.frm.sP.value = document.frm.selU[i].text; 
   		}else{
   		opener.document.frm.iP.value = opener.document.frm.iP.value +","+ document.frm.selU[i].value; 
   		opener.document.frm.sP.value = opener.document.frm.sP.value +","+ document.frm.selU[i].text; 
   		}
   		strT = strT+  "<tr><td bgcolor=#EEEEEE>"+document.frm.selU[i].text+"</td><td bgcolor=#FFFFFF align='center'><input type='text' name='mPM' id='mPM'  value='' size='10' onKeyUp=jsSetMoney('m',"+i+",<%=iType%>) >��</td><td align='center' bgcolor=#FFFFFF><input type='text' name='iPM' id='iPM' value='' size='3'  onKeyUp=jsSetMoney('i',"+i+",<%=iType%>)>%</td></tr>";
	} 
	strT = strT+"</table>";  
	opener.document.all.divPM.innerHTML = strT;		
	self.close();
}
 
 
$(window).load(function(){ //������ �ε��   
	if($("#iP",window.opener.document).val() != ""){ //���� ���ð� ���� ���  
		var arrI = $("#iP",window.opener.document).val().split(",");
		var arrN = $("#sP",window.opener.document).val().split(",");   
		 
		for(i=0;i<arrI.length;i++){ 
			$("#selU").append("<option value='"+arrI[i]+"'>"+arrN[i]+"</option>"); //�ɼǰ� �߰�
		}
	} 
});
	
//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
<form name="frm" method="post"> 
<input type="hidden" name="iType" value="<%=iType%>">
<tr>
	<td>�ڱݱ��� �μ� ���<br><hr width=100%></td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
		<tr> 
			<td  align="center">
			<div id="sp1">
			<select name="selp1" id="selp1">
			<option value="0">--����--</option>
			<%  
			Set clsPart = new CpartMoneyCls 
				clsPart.Fstep1partidx = 0
				clsPart.Fstep2partidx = 0
				clsPart.FeappDepth = 1
				arrData = clsPart.fnGetPartList  
			Set clsPart = nothing	
			IF isArray(arrData) THEN
				For intD = 0 To UBound(arrData,2)
			%>
				<option value="<%=arrData(0,intD)%>" <%IF Cstr(istep1partidx) = Cstr(arrData(0,intD)) THEN%>selected<%END IF%>><%=arrData(4,intD)%></option>
			<%	
				Next
			END IF
			%>
			</select>
			</div>	<br>
			<div id="sp2"> 
			<select name="selp2" id="selp2">
			<option value="0">--����--</option>
			<%  
			IF istep1partidx > 0 THEN
			Set clsPart = new CpartMoneyCls 
				clsPart.Fstep1partidx = istep1partidx
				clsPart.Fstep2partidx = 0
				clsPart.FeappDepth = 2
				arrData = clsPart.fnGetPartList  
			Set clsPart = nothing	
			IF isArray(arrData) THEN
				For intD = 0 To UBound(arrData,2)
			%>
				<option value="<%=arrData(0,intD)%>"  <%IF Cstr(istep2partidx) = Cstr(arrData(0,intD)) THEN%>selected<%END IF%>><%=arrData(4,intD)%></option>
			<%	
				Next
			END IF
			END IF
			%>
			</select>
			</div><br>
			<div id="sp3"> 
			<select name="selp3" id="selp3">
			<option value="0">--����--</option>
			<%  
			IF istep2partidx > 0 THEN
			Set clsPart = new CpartMoneyCls 
				clsPart.Fstep1partidx = istep1partidx
				clsPart.Fstep2partidx = istep2partidx
				clsPart.FeappDepth = 3
				arrData = clsPart.fnGetPartList  	
			Set clsPart = nothing	
			IF isArray(arrData) THEN
				For intD = 0 To UBound(arrData,2)
			%>
				<option value="<%=arrData(0,intD)%>"  <%IF Cstr(istep3partidx) = Cstr(arrData(0,intD)) THEN%>selected<%END IF%>><%=arrData(4,intD)%></option>
			<%	
				Next
			END IF
			END IF
			%>
			</select>
			</div>
			</td>
			<td align="center">
				<input type="button" id="btnAdd" value="�߰���" class="button"> <br><br>
				<input type="button" id="btnDel" value="������" class="button"> 
			</td>
			<td align="center">
				<select name="selU"  id="selU" multiple size="15" style="width:300px">
				</select>
			</td>
		</tr>
		<Tr>
			<td align="center" colspan="3" height="50"><input type="button" value="Ȯ��" class="button" onClick="jsSubmitPartMoney();"></td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
 </body>
 </html>