<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/event/v5/popup/pop_app_event_item_regist.asp
' Description :  ������ �̺�Ʈ ��ǰ ��� ����
' History : 2023.01.09 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/appDedicatedEventCls.asp"-->
<%
'��������
dim evt_code : evt_code = requestCheckvar(request("evt_code"),10)
Dim oAppDedicated, arrList, intLoop

set oAppDedicated = new AppEventCls
oAppDedicated.FRectEventCode = evt_code
arrList = oAppDedicated.fnGetAppDedicatedItemList
%>
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<link rel="stylesheet" href="http://code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">

<style type="text/css">
div.btmLine {background:url(/images/partner/admin_grade.png) left bottom repeat-x; padding-bottom:5px;}    
.tab {position:relative; z-index:50;}
.tab ul {_zoom:1; border-left:1px solid #ccc; border-bottom:1px solid #ccc; list-style:none; margin:0; padding:0;}
.tab ul:after {content:""; display:block; height:0; clear:both; visibility:hidden;}
.tab ul li {float:left; text-align:center;height:23px;padding-top:7px; border:1px solid #ccc; margin:0 0 -1px -1px; cursor:pointer;  background-color:#fff; }
.tab ul li.selected {background-color:#FAECC5; position:relative; font-weight:bold;}
.col11 {width:15% !important;}

select {font-size:12px; vertical-align:top;}
input[type=button], input[type=text] {vertical-align:top;}
</style>
<script>
$(function(){
    $("#datepicker").datepicker({
        changeMonth: true, 
        changeYear: true,
        minDate: '-50y',
        nextText: '���� ��',
        prevText: '���� ��',
        yearRange: 'c-50:c+20',
        showButtonPanel: true, 
        currentText: '���� ��¥',
        closeText: '�ݱ�',
        dateFormat: "yy-mm-dd",
        showAnim: "slide",
        showMonthAfterYear: true, 
        dayNamesMin: ['��', 'ȭ', '��', '��', '��', '��', '��'],
        monthNamesShort: ['1��','2��','3��','4��','5��','6��','7��','8��','9��','10��','11��','12��']
    });
    $("#datepicker2").datepicker({
        changeMonth: true, 
        changeYear: true,
        minDate: '-50y',
        nextText: '���� ��',
        prevText: '���� ��',
        yearRange: 'c-50:c+20',
        showButtonPanel: true, 
        currentText: '���� ��¥',
        closeText: '�ݱ�',
        dateFormat: "yy-mm-dd",
        showAnim: "slide",
        showMonthAfterYear: true, 
        dayNamesMin: ['��', 'ȭ', '��', '��', '��', '��', '��'],
        monthNamesShort: ['1��','2��','3��','4��','5��','6��','7��','8��','9��','10��','11��','12��']
    });
    $("#datepicker3").datepicker({
        changeMonth: true, 
        changeYear: true,
        minDate: '-50y',
        nextText: '���� ��',
        prevText: '���� ��',
        yearRange: 'c-50:c+20',
        showButtonPanel: true, 
        currentText: '���� ��¥',
        closeText: '�ݱ�',
        dateFormat: "yy-mm-dd",
        showAnim: "slide",
        showMonthAfterYear: true, 
        dayNamesMin: ['��', 'ȭ', '��', '��', '��', '��', '��'],
        monthNamesShort: ['1��','2��','3��','4��','5��','6��','7��','8��','9��','10��','11��','12��']
    });
});
function fnAddItem(frm){
    if(frm.itemid.value==""){
        alert("��ǰ�ڵ带 �Է��ϼ���.");
    }else if(frm.startdate.value==""){
        alert("�������� �Է��ϼ���.");
    }else if(frm.startdate.value==""){
        alert("�������� �Է��ϼ���.");
    }else{
        frm.submit();
    }
}
//����
function jsDel(sType, iValue){	
	var frm;		
	var sValue;		
	frm = document.fitem;
	sValue = "";
	
	if (sType ==0) {
		if(!frm.chkI) return;
		
		if (frm.chkI.length > 1){
		for (var i=0;i<frm.chkI.length;i++){
			if(frm.chkI[i].checked){
				if (sValue==""){
					sValue = frm.chkI[i].value;		
				}else{
					sValue =sValue+","+frm.chkI[i].value;		
				}	
			}
		}	
		}else{
			if(frm.chkI.checked){
				sValue = frm.chkI.value;
			}	
		}
	
		if (sValue == "") {
			alert('���� ��ǰ�� �����ϴ�.');
			return;
		}
		document.frmDel.itemidarr.value = sValue;
	}else{
		document.frmDel.itemidarr.value = iValue;
	}	
		
	if(confirm("�����Ͻ� ��ǰ�� �����Ͻðڽ��ϱ�?")){
		document.frmDel.submit();
	}
}
//��ü����
var ichk;
ichk = 1;
	
function jsChkAll(){
	var frm, blnChk;
	frm = document.fitem;
	
	if(!frm.chkI) return;
	if ( ichk == 1 ){
		blnChk = true;
		ichk = 0;
	}else{
		blnChk = false;
		ichk = 1;
	}
	for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA		
		if ((e.type=="checkbox")) {				
			e.checked = blnChk ;
		}
	}
}
function fnEventPrizeAdd(episode,itemid){
	var winpop = window.open('/admin/eventmanage/event/v5/popup/pop_appDedicatedEvent_PrizeSet.asp?evt_code=<%=evt_code%>&episode='+episode+'&itemid='+itemid,'winPrize','width=600,height=400,scrollbars=yes,resizable=yes');
}
</script>
<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a"  style="padding-top:10px">
	<tr><!-- �˻�--->
		<td>     
		    <table cellspacing="5"  bgcolor="FAECC5" width="100%" class="a" cellpadding="0">
		        <tr>
		            <td bgcolor="#FFFFFF"> 
            			<form name="frmitemAdd" method="post" action="appDedicatedItem_process.asp"> 
                        <input type="hidden" name="evt_code" value="<%=evt_code%>">
                        <input type="hidden" name="mode" value="add">
            			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
            				<tr align="center" >
            					<td  width="100" bgcolor="<%= adminColor("tabletop") %>">ȸ��(��ǰ) �߰�</td>
            					<td align="left"  bgcolor="#ffffff">  	 
            						<table border="0" cellpadding="1" cellspacing="1" class="a">
            						<tr>
            							<td style="white-space:nowrap;padding-left:10px;">ȸ��: <select name="episode"><option value="1">1</option><option value="2">2</option><option value="3">3</option><option value="4">4</option><option value="5">5</option><option value="6">6</option><option value="7">7</option><option value="8">8</option><option value="9">9</option></select></td>  
            							<td style="white-space:nowrap;padding-left:10px;">��ǰ�ڵ�: <input type="text" class="text" name="itemid" size="15" maxlength="10" /></td>
										<td style="white-space:nowrap;padding-left:10px;">��÷��: <input type="text" class="text" name="prize_count" size="4" maxlength="2" /></td>
										<td style="white-space:nowrap;padding-left:10px;">��÷���÷�: <input type="text" class="text" name="prize_count_color" size="15" maxlength="10" /></td>
                                        <td style="white-space:nowrap;padding-left:10px;">������: <input type="text" class="text" name="startdate" id="datepicker" size="15" maxlength="10" /></td>
                                        <td style="white-space:nowrap;padding-left:10px;">������: <input type="text" class="text" name="enddate" id="datepicker2" size="15" maxlength="10" /></td>
                                        <td style="white-space:nowrap;padding-left:10px;">��÷��ǥ��: <input type="text" class="text" name="prizedate" id="datepicker3" size="15" maxlength="10" />
											<select name="prizetime">
												<option value="0">0��</option>
												<option value="1">1��</option>
												<option value="2">2��</option>
												<option value="3">3��</option>
												<option value="4">4��</option>
												<option value="5">5��</option>
												<option value="6">6��</option>
												<option value="7">7��</option>
												<option value="8">8��</option>
												<option value="9">9��</option>
												<option value="10">10��</option>
												<option value="11">11��</option>
												<option value="12">12��</option>
												<option value="13">13��</option>
												<option value="14">14��</option>
												<option value="15">15��</option>
												<option value="16">16��</option>
												<option value="17">17��</option>
												<option value="18">18��</option>
												<option value="19">19��</option>
												<option value="20">20��</option>
												<option value="21">21��</option>
												<option value="22">22��</option>
												<option value="23">23��</option>
												<option value="24">24��</option>
											</select>
										</td>
                                        <td style="white-space:nowrap;padding-left:10px;"><input type="button" class="button_s" value="���" onClick="javascript:fnAddItem(this.form);"></td>
            						</tr>
            			        	</table>
            			        </td>
            			    </tr> 
            			</table>
            			</form>
            		</td>
            	</tR> <!-- �˻�--->
            	
            	<tr>
            		<td style="padding-top:10px;" valign="top" >  
            		  <div id="divPC">
            		 <form name="fitem" method="post">
                     <input type="hidden" name="mode" value="">
                     <input type="hidden" name="evt_code" value="<%=evt_code%>">
            		  <table width="100%" border="0" align="center" cellpadding="0"  class="a" cellspacing="0"  >	  
            		     <tr>
            		       <td>
                                <input type="button" value="���û���" onclick="jsDel(0,'');" class="button">
                      	    </td>
                          </tr>  
                          <tr>
                      	    <td> 
                      			<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>"> 
                      			    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
                      			    	<td><input type="checkbox" name="chkA" onClick="jsChkAll();"></td>    				    	
                       			    	<td>ȸ��</td>
										<td>��ǰID</td>
                      					<td>�̹���</td>
										<td>��ǰ��</td>
                      					<td>�ǸŰ�</td>
										<td>��÷��</td>
                      					<td>������</td>
                      					<td>������</td>
                                        <td>��÷��ǥ��</td>
                                        <td>��÷��ǥ</td>
                      				</tr>
									<tbody>
                      			    <%
									IF isArray(arrList) THEN 
                      			    	For intLoop = 0 To UBound(arrList,2)
                      			    %>
                      			    <tr align="center" bgcolor="#FFFFFF">    
                      			    	<td><input type="checkbox" name="chkI" value="<%=arrList(0,intLoop)%>"></td>    				    	
                      			    	<td><%=arrList(1,intLoop)%></td>
										<td><%=arrList(2,intLoop)%></td>
                                        <td>
                                            <% if (Not IsNull(arrList(7,intLoop))) and (arrList(7,intLoop)<>"") then %>
                                                <img src="http://webimage.10x10.co.kr/image/list/<%=GetImageSubFolderByItemid(arrList(2,intLoop))%>/<%=arrList(7,intLoop)%>" width="50%">
                                            <%end if%>
                                        </td>
										<td><%=db2html(arrList(5,intLoop))%></td>
                                        <td><%=formatnumber(arrList(6,intLoop),0)%>��</td>
                      			    	<td><%=arrList(10,intLoop)%><br><%=arrList(11,intLoop)%></td>
										<td><%=arrList(3,intLoop)%></td>
                                        <td><%=arrList(4,intLoop)%></td>
                                        <td><%=arrList(9,intLoop)%></td>
                                        <td>
                                            <% if arrList(8,intLoop)="N" then %>
                                                <input type="button" value="��÷��ǥ" onclick="fnEventPrizeAdd(<%=arrList(1,intLoop)%>,<%=arrList(2,intLoop)%>);" class="button">
                                            <% else %>
                                                <input type="button" value="��÷�ں���" onclick="fnEventPrizeAdd(<%=arrList(1,intLoop)%>,<%=arrList(2,intLoop)%>);" class="button">
                                            <% end if %>
                                        </td>
                      			    </tr>   
                      				<%	
										Next
                      				ELSE
                      				%>
                      			   	<tr align="center" bgcolor="#FFFFFF">
                      			   		<td colspan="19">��ϵ� ������ �����ϴ�.</td>
                      			   	</tr>	
                      			   	<%END IF%>
									</tbody>
                      			</table>
                              </td>
                             </tr>
                          </table>
                          </form>    
                      </div> 
            		</td>
            	</tR> 
            </table>
        </tD>
    </tr> 
</table>
<%
	set oAppDedicated = nothing
%>	
<iframe name="ifrmProc" src="about:blank;" frameborder="0" width="0" height="0"></iframe>
<!-- ���� ����--->
<form name="frmDel" method="post" action="appDedicatedItem_process.asp">
<input type="hidden" name="mode" value="del">
<input type="hidden" name="evt_code" value="<%=evt_code%>">
<input type="hidden" name="itemidarr">
</form>
<!-- ǥ �ϴܹ� ��-->		
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->