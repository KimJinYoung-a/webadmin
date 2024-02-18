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
	//달력
	function jsPopCal(sName){
	 if(!document.all.chkSR.checked){
	 	 document.all.chkSR.checked= true;
	 	}
		var winCal;
		winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	 
	}
	
	//오픈예약
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
            	alert('[' + itemKey + '] - 판매여부가 N인 상품만 오픈예약 할 수 있습니다.');
                return;
            }
            
           
           if ((val_deliveryTypePolicy=="1" || val_deliveryTypePolicy =="4") && var_realstock==0){ 
            	alert('[' + itemKey + '] - 아직 입고확인되지 않은 상품입니다. 텐바이텐 배송일 경우, 입고확인 후 오픈예약이 가능합니다. 확인 후 다시 시도해주세요');
                return;
            }
            
            eval("opener.document.all.sellreserve_"+itemKey).innerHTML = "오픈예약:"+document.all.dSR.value;
        }
    }else{
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                itemKey = frm.cksel[i].value;
                val_sellyn = eval("frm.defaultsellyn_" + itemKey ).value;
		            val_deliveryTypePolicy = eval("frm.deliveryTypePolicy_" + itemKey ).value;
		            var_realstock =  eval("document.frmSR.chkRStock_" + itemKey ).value;
		             
		            if (val_sellyn!="N"){ 
		            	alert('[' + itemKey + '] - 판매여부가 N인 상품만 오픈예약 할 수 있습니다.');
		                return;
		            }
		            
		           if ((val_deliveryTypePolicy=="1" || val_deliveryTypePolicy =="4") && var_realstock==0){ 
		            	alert('[' + itemKey + '] - 아직 입고확인되지 않은 상품입니다. 텐바이텐 배송일 경우, 입고확인 후 오픈예약이 가능합니다. 확인 후 다시 시도해주세요');
		                return;
		            }
		          
		           eval("opener.document.all.sellreserve_"+itemKey).innerHTML = "오픈예약:"+document.all.dSR.value;
		           
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
'### 오픈예약 조건(텐배상품 재고수량) 체크 
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
		<td>상품오픈예약(<%=iCnt%>건)<hr width="100%"></td>
	</tr> 
	 <tr> 
			<td style="padding:3px">
				<input type="checkbox" name="chkSR" value="Y" onClick="jsChkSellReserve();"> 상품오픈예약: 
				<input type="text" name="dSR" value="" size="10" class="input"   onClick="jsPopCal('dSR');"> 
				<input type="image" name="imgSR" src="/images/admin_calendar.png" onClick="jsPopCal('dSR');"  > 
				  <div style="padding:3px;">사용안함 상태일 경우 예약된 시간에 오픈이 되지 않습니다. 
			   텐바이텐 배송일 경우, 입고 확인 후 오픈예약이 가능합니다.  <br><br> 
<font color="blue">오픈날짜 지정 후, [선택상품저장]을 눌러주셔야 설정이 완료됩니다.</font>
			   </div>
			</td> 
	</tR> 
	<tr>
		<td align="center"><input type="button" class="button" value="취소" onclick="javascript:jsCancel();"> <input type="button" class="button" value="확인" onClick="jsSetSellDate();"></td>
	</tr>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->