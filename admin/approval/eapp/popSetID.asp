<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������� ���
' History : 2011.03.16 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp"-->
<%
	Dim part_sn, ireportidx
	Dim iMode, ilastApprovalID,iAuthPosition,sjob_name

	part_sn 		= requestCheckvar(Request("part_sn"),10)
	ireportidx 		= requestCheckvar(Request("iridx"),10)
	iMode			= requestCheckvar(Request("iM"),1)
	ilastApprovalID	= requestCheckvar(Request("iLAI"),10)
	iAuthPosition	= requestCheckvar(Request("iAP"),10)
  	sjob_name		= requestCheckvar(Request("sjn"),30)
	IF part_sn = "" THEN part_sn =0
	'// �������� ����Ʈ
	dim oMember, arrList,intLoop
	Set oMember = new CTenByTenMember
	oMember.Fpart_sn 		= part_sn
	arrList = oMember.fnGetPartUserList
	set oMember = nothing

%>
<script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script>
<script language="javascript">
<!--
$(document).ready(function(){
	$("#part_sn").change(function(){
		var part_sn = $("#part_sn").val();
		 var url="ajaxPartUserList.asp";
		 var params = "part_sn="+part_sn;

		 $.ajax({
		 	type:"POST",
		 	url:url,
		 	data:params,
		 	success:function(args){
		 		$("#devU").html(args);
		 	},

		 	error:function(e){
		 		alert("�����ͷε��� ������ ������ϴ�. �ý������� �������ּ���");
		 		//alert(e.responseText);
		 	}
		 });
	});
});

var mode = "<%=iMode%>";
var ilastApprovalID = "<%=ilastApprovalID%>";
var iAuthPosition ="<%=iAuthPosition%>";
$(function(){
	//�߰���ư Ŭ���� �̺�Ʈ
	$("#btnAdd").click(function(){
	 var sValue;
	 var sText;
	if( $("#selUL option:selected").size() < 1){return}; //���ð� ���� ��� return ó��
	if ($("#selUL option:selected").size()> 1) {
	  alert("���� �Ѱ��� �������ּ���");
	  return;
	 	}

	sValue = $("#selUL option:selected").val();  //���ð�
	sText  = $("#selUL option:selected").text(); //���ð��� �ؽ�Ʈ


	for(j=0; j<$("#selUC option").size();j++){

		if($("#selUC option:eq("+j+")").val()==sValue){
			alert("�̹� ��ϵǾ��ֽ��ϴ�.");
			return;
		}
	}
 if (mode !=2){
	if ($("#selUC option").size() > 0){
		alert("�����ڴ� �Ѹ� ��ϰ����մϴ�.");
		return;
	}
}
	$("#selUC").append("<option value='"+sValue+"'>"+sText+"</option>"); //�߰�ó��
	});

	//����
	$("#btnDel").click(function(){
		 $("#selUC option:selected").remove();
	});
});


//opener ������� �߰�
	function jsSetId(){
		var strMsg = "";
		var strUser ="";
		if(mode==1){
			strUser = "������";
		}else if(mode==3){
			strUser = "�����";
		}else if(mode==4){
			strUser = "������";
		}else{
			strUser = "������";
		}

		if(document.frm.selUC.length==0){
 			alert(strUser+"�� �߰����ּ���");
		 return;
	 	}
		var arrValue =  document.frm.selUC[0].value.split("-");
		if (mode==1){ //����ǰ�Ǽ� ���缱 ���
			if( arrValue[1]<=ilastApprovalID && arrValue[1] > 0){	//�������� ���޼��ý�
				if(confirm("�������� �����Դϴ�. ���������ڷ� �����Ͻðڽ��ϱ�?\n\n���ǰ� �ʿ��� ��� �����ڸ� �߰� ����Ͻñ� �ٶ��ϴ�.")){
					//������ div null ó��
					strMsg = "<table width='100%' cellpadding='5' cellspacing='0' class='a'>"
							+ "<tr><td align='Center' bgcolor='#E6E6E6'>&nbsp;</td></tr>"
							+ "<tr><td align='Center'>&nbsp;</td></tr>"
							+ "</table>"
					opener.eval("document.all.dAP"+iAuthPosition).innerHTML = strMsg;

					//�������� div�� ������
					strMsg = "<table width='100%'  cellpadding='5' cellspacing='0' class='a'>"
							+ "<tr><td align='Center' bgcolor='#E6E6E6'>����������</td></tr>"
							+ "<tr><td align='Center'><input type='text' name='sASD' style='border:0;text-align:center;' value='���δ��'></td></tr>"
							+ "<tr><td align='Center'><input type='text' name='sALN' id='sALN' value='"+document.frm.selUC[0].text+"' style='border:0;text-align:center;' readonly size='20'>"
							+ "<input type='hidden' name='hidAJ' id='hidAJ' value='"+ arrValue[1]+"'></td></tr>"
							+ "<tr><td align='Center'><input type='text' name='sADD' value='' style='border:0;text-align:center;'></td></tr>"
							+ "<tr><td align='Center'><input type='button' class='button' value='������ ���' onClick='jsRegID(1,0);'><br><input type='checkbox' value='1' name='chkSms' checked> SMS����</td></tr>"
							+ "</table> "
					opener.document.all.dAP0.innerHTML = strMsg;
					opener.document.frm.hidAI.value = arrValue[0] ;
					opener.document.frm.hidPS.value = document.frm.part_sn.value;
					opener.document.frm.blnL.value = 1;

                    //������ �߰� (���� ���� ���ýø� ����)
					strMsg = "<table width='100%'  cellpadding='5' cellspacing='0' class='a'>"
					       + "<tr><td align='Center' bgcolor='#E6E6E6'>����</td></tr>"
					       + "<tr><td align='Center'>&nbsp;</td></tr>"
						   + "<tr><td align='Center'><input type='button' class='button' value='������ ���' onClick='jsRegID_H(4);'></td></tr>"
						   + "</table>"

					opener.document.all.dAP_H.innerHTML = strMsg;
                    opener.document.frm.hidAI_H.value = '';//�ʱ�ȭ
			        opener.document.frm.hidPS_H.value = '';
					self.close();
				}
			}else{   //�Ϲݰ����� ���ý�
				if(opener.document.frm.blnL.value==1){ //������ ���������� ���� �ϰ� ������ ��� �������� div null ó��
			 	strMsg = "<table width='100%' cellpadding='5' cellspacing='0' class='a'>"
							+ "<tr><td align='Center' bgcolor='#E6E6E6'>����������</td></tr>"
							+ "<tr><td align='Center'>&nbsp;</td></tr>"
							+ "<tr><td align='Center'><%=sjob_name%></td></tr>"
							+ "</table>"
				opener.document.all.dAP0.innerHTML = strMsg;
				}
				//���� �ֱ�
				strMsg = "<table width='100%'  cellpadding='5' cellspacing='0' class='a'>"
							+ "<tr><td align='Center' bgcolor='#E6E6E6'>"+iAuthPosition+"�� ����</td></tr>"
							+ "<tr><td align='Center'><input type='text' name='sASD' style='border:0;text-align:center;' value='���δ��'></td></tr>"
							+ "<tr><td align='Center'><input type='text' name='sALN' id='sALN' value='"+document.frm.selUC[0].text+"' style='border:0;text-align:center;' readonly size='20'>"
							+ "<input type='hidden' name='hidAJ' id='hidAJ' value='"+ arrValue[1]+"'></td></tr>"
							+ "<tr><td align='Center'><input type='text' name='sADD' value='' style='border:0;text-align:center;'></td></tr>"
							+ "<tr><td align='Center'><input type='button' class='button' value='������ ���' onClick='jsRegID(1,"+iAuthPosition+");'><br><input type='checkbox' value='1' name='chkSms' checked> SMS����</td></tr>"
							+ "</table> "
				opener.eval("document.all.dAP"+iAuthPosition).innerHTML = strMsg;
				opener.document.frm.hidAI.value = arrValue[0] ;
				opener.document.frm.hidPS.value = document.frm.part_sn.value;
				opener.document.frm.blnL.value = 0;
				self.close();
			}
		}else if(mode==3){	//�繫ȸ�� �����
			opener.document.frm.hidAI.value=arrValue[0];
			opener.document.frm.sALN.value=document.frm.selUC[0].text;
			opener.document.frm.hidAJ.value= arrValue[1];
			self.close();
        }else if(mode==4){	//������
            //������ �߰� (���� ���� ���ýø� ����)
            strMsg = "<table width='100%'  cellpadding='5' cellspacing='0' class='a'>"
							+ "<tr><td align='Center' bgcolor='#E6E6E6'>����</td></tr>"
							+ "<tr><td align='Center'><input type='text' name='sASD_H' style='border:0;text-align:center;' value='���δ��'></td></tr>"
							+ "<tr><td align='Center'><input type='text' name='sALN_H' id='sALN_H' value='"+document.frm.selUC[0].text+"' style='border:0;text-align:center;' readonly size='20'>"
							+ "<input type='hidden' name='hidAJ_H' id='hidAJ_H' value='"+ arrValue[1]+"'></td></tr>"
							+ "<tr><td align='Center'><input type='text' name='sADD_H' value='' style='border:0;text-align:center;'></td></tr>"
							+ "<tr><td align='Center'><input type='button' class='button' value='������ ���' onClick='jsRegID_H(4);'><br><input type='checkbox' value='1' name='chkSms_H' checked> SMS����</td></tr>"
							+ "</table> "
			opener.document.all.dAP_H.innerHTML = strMsg;
			opener.document.frm.hidAI_H.value = arrValue[0] ;
			opener.document.frm.hidPS_H.value = document.frm.part_sn.value;
			self.close();
		}else{ //������ ���
			for(i=0;i<frm.selUC.length;i++){
				if(i==0){
				opener.document.frm.sRfN.value = document.frm.selUC[i].text ;
				opener.document.frm.hidRfI.value =arrValue[0] ;
				}else{
				opener.document.frm.sRfN.value = opener.document.frm.sRfN.value +"," + document.frm.selUC[i].text ;
				opener.document.frm.hidRfI.value =opener.document.frm.hidRfI.value +","+ document.frm.selUC[i].value.split("-")[0] ;
				}
			}
			opener.document.frm.hidPS.value = document.frm.part_sn.value;
			self.close();
		}
	}

$(window).load(function(){ //������ �ε��

	if ((mode!=2)&&(mode!=4)){
		if( ($("#hidAI",window.opener.document).val() != "") && ($("#hidAI",window.opener.document).val() !=  undefined )){ //���� ���ð� ���� ���
		var sText = $("#sALN",window.opener.document).val();
		var sValue = $("#hidAI",window.opener.document).val()+"-"+$("#hidAJ",window.opener.document).val();

		$("#selUC").append("<option value='"+sValue+"'>"+sText+"</option>"); //�ɼǰ� �߰�
		}
	}else if (mode==4){ //������
	    if( ($("#hidAI_H",window.opener.document).val() != "") && ($("#hidAI_H",window.opener.document).val() !=  undefined )){ //���� ���ð� ���� ���
		var sText = $("#sALN_H",window.opener.document).val();
		var sValue = $("#hidAI_H",window.opener.document).val()+"-"+$("#hidAJ_H",window.opener.document).val();

		$("#selUC").append("<option value='"+sValue+"'>"+sText+"</option>"); //�ɼǰ� �߰�
		}
	}else{ //2 : ����
		if($("#hidRfI",window.opener.document).val() != "" && typeof($("#hidRfI",window.opener.document).val()) !=  undefined){ //���� ���ð� ���� ���
    		var sText = $("#sRfN",window.opener.document).val();
    		var sValue = $("#hidRfI",window.opener.document).val();
            if (sValue!=undefined){
        		var arrN = sText.split(",");
        		var arrI = sValue.split(",");
        		for(i=0;i<arrI.length;i++){
        		$("#selUC").append("<option value='"+arrI[i]+"'>"+arrN[i]+"</option>"); //�ɼǰ� �߰�
        		}
    		}
		}
	}
});


//-->
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
<form name="frm" method="post">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" border="0">
		<tr>
			<td  align="center"> <%=printPartOptionAddEtc("part_sn", part_sn, "id=part_sn")%></td>
		</tr>
		<tr>
			<td align="center">
			<div id="devU">
				<select  name="selUL" id="selUL" multiple size="20" style="width:200px">
				<%IF isArray(arrList) THEN
					For intLoop = 0 To UBOund(arrList,2)
				%>
					<option value="<%=arrList(2,intLoop)&"-"&arrList(4,intLoop)%>"><%=arrList(1,intLoop)%>&nbsp;<%=arrList(7,intLoop)%> <%=arrList(2,intLoop)%>   </option>
				<%	Next
				END IF%>
				</select>
			</div>
			</td>
		</tr>
		</table>
	</td>
	<td>
		<input type="button" id="btnAdd" value="�߰���" class="button"> <br><br>
		<input type="button" id="btnDel" value="������" class="button">
	</td>
	<td  valign="bottom">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
		<tr>
			<td>
				<select  name="selUC" id="selUC" multiple size="20" style="width:200px">

				</select>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" colspan="3"><input type="button" class="button" value="���" onClick="jsSetId();"></td>
</tr>

</form>
</table>
