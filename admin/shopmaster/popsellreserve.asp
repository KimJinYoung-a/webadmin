<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
dim iCnt, arrItemid, itemid, i 
iCnt = requestCheckVar(request("iCnt"),10)
arrItemid = requestCheckVar(request("Ritemid"),500)
itemid = split(arrItemid,",") 
%>
<script type="text/javascript">
	//�޷�
	function jsPopCal(sName){
	 if(!document.all.chkSR.checked){
	 	 document.all.chkSR.checked= true;
	 	}
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	 
	}
	
	//���¿���
	function jsChkSellReserve(){
		if(!document.all.chkSR.checked){ 
			document.all.dSR.value = "";
		} 
	}
	
	function jsSetSellDate(){   
	 	var frm = opener.document.frmSvArr;
	  var itemKey = "";
	  var pass = false;
    var val_sellyn, val_defaultDeliveryType,var_realstock;
		frm.dSR.value = document.all.dSR.value;  
 	
    if (!frm.cksel.length){
        if (frm.cksel.checked){ 
            itemKey = frm.cksel.value;
            val_sellyn = eval("frm.defaultsellyn_" + itemKey ).value;
            val_deliveryTypePolicy = eval("frm.deliveryTypePolicy_" + itemKey ).value; 
            var_realstock =  eval("document.frmSR.chkRStock_" + itemKey ).value;
            
            if (val_sellyn!="N"){
            	alert('[' + itemKey + '] - �Ǹſ��ΰ� N�� ��ǰ�� ���¿��� �� �� �ֽ��ϴ�.');
                return;
            }
            
           
           if ((val_deliveryTypePolicy=="1" || val_deliveryTypePolicy =="4") && var_realstock==0){ 
            	alert('[' + itemKey + '] - ���� �԰�Ȯ�ε��� ���� ��ǰ�Դϴ�. �ٹ����� ����� ���, �԰�Ȯ�� �� ���¿����� �����մϴ�. Ȯ�� �� �ٽ� �õ����ּ���');
                return;
            }
            
            eval("opener.document.all.sellreserve_"+itemKey).innerHTML = "���¿���:"+document.all.dSR.value;
        }
    }else{
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                itemKey = frm.cksel[i].value;
                val_sellyn = eval("frm.defaultsellyn_" + itemKey ).value;
		            val_deliveryTypePolicy = eval("frm.deliveryTypePolicy_" + itemKey ).value;
		            var_realstock =  eval("document.frmSR.chkRStock_" + itemKey ).value;
		             
		            if (val_sellyn!="N"){ 
		            	alert('[' + itemKey + '] - �Ǹſ��ΰ� N�� ��ǰ�� ���¿��� �� �� �ֽ��ϴ�.');
		                return;
		            }
		            
		           if ((val_deliveryTypePolicy=="1" || val_deliveryTypePolicy =="4") && var_realstock==0){ 
		            	alert('[' + itemKey + '] - ���� �԰�Ȯ�ε��� ���� ��ǰ�Դϴ�. �ٹ����� ����� ���, �԰�Ȯ�� �� ���¿����� �����մϴ�. Ȯ�� �� �ٽ� �õ����ּ���');
		                return;
		            }
		          
		           eval("opener.document.all.sellreserve_"+itemKey).innerHTML = "���¿���:"+document.all.dSR.value;
		           
            }
        }
     }   
		 self.close();
	}
	
	function jsCancel(){ 
		self.close();
	}
	
</script>
<form name="frmSR" method="post">
<%
'### ���¿��� ����(�ٹ��ǰ ������) üũ 
dim objCmd, returnValue   
For i = 0 To UBound(itemid) 
set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call db_item.[dbo].[sp_Ten_item_sellreserve_chkStock]("&trim(itemid(i))&")}"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With
	    returnValue = objCmd(0).Value
Set objCmd = nothing
%>
<input type="hidden" name="chkRStock_<%=trim(itemid(i))%>" value="<%=returnValue%>">
<%Next%>
</form>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td>��ǰ���¿���(<%=iCnt%>��)<hr width="100%"></td>
	</tr> 
	 <tr> 
			<td style="padding:3px">
				<input type="checkbox" name="chkSR" value="Y" onClick="jsChkSellReserve();"> ��ǰ���¿���: 
				<input type="text" name="dSR" value="" size="10" class="input"   onClick="jsPopCal('dSR');"> 
				<input type="image" name="imgSR" src="/images/admin_calendar.png" onClick="jsPopCal('dSR');"  > 
				  <div style="padding:3px;">������ ������ ��� ����� �ð��� ������ ���� �ʽ��ϴ�. 
			   �ٹ����� ����� ���, �԰� Ȯ�� �� ���¿����� �����մϴ�.  <br><br> 
<font color="blue">���³�¥ ���� ��, [���û�ǰ����]�� �����ּž� ������ �Ϸ�˴ϴ�.</font>
			   </div>
			</td> 
	</tR> 
	<tr>
		<td align="center"><input type="button" class="button" value="���" onclick="javascript:jsCancel();"> <input type="button" class="button" value="Ȯ��" onClick="jsSetSellDate();"></td>
	</tr>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->