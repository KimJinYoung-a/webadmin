<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� �ֹ����� ����
' Hieditor : 2015.06.18 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/incsessionadmin.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->

<%
dim voutmallorderseq, voutmallorderserial, vxSiteOrdercount, i
	voutmallorderseq = requestcheckvar(request("outmallorderseq"),10)

if voutmallorderseq="" then
	response.write "<script type='text/javascript'>alert('��ĪŰ�� �����ϴ�'); self.close();</script>"
	dbget.close()	:	response.end
end if

dim oorderedit
set oorderedit = new CxSiteTempOrder
    oorderedit.FPageSize=1
    oorderedit.FCurrPage=1
	oorderedit.frectoutmallorderseq = voutmallorderseq
	oorderedit.frectoutmallorderserial = voutmallorderserial
	oorderedit.fxsiteorderedit()

if oorderedit.ftotalcount=0 then
	response.write "<script type='text/javascript'>alert('�ֹ����� �����ϴ�.'); self.close();</script>"
	dbget.close()	:	response.end
end if

if oorderedit.foneitem.fsellsite<>"cjmall" and oorderedit.foneitem.fsellsite<>"lotteimall" then
	response.write "<script type='text/javascript'>alert('CJMALL/Lotteimall �ֹ��Ǹ� ������ �����մϴ�.'); self.close();</script>"
	dbget.close()	:	response.end
end if

if oorderedit.foneitem.fOrderSerial<>"" then
	response.write "<script type='text/javascript'>alert('�� �Է� �ֹ���.'); self.close();</script>"
	dbget.close()	:	response.end
end if

voutmallorderserial = oorderedit.foneitem.fOutMallOrderSerial

vxSiteOrdercount = getxSiteDuppReceiverCheck(voutmallorderserial)
if vxSiteOrdercount=0 then
	response.write "<script type='text/javascript'>alert('�ش�Ǵ� �ֹ����� �����ϴ�.'); self.close();</script>"
	dbget.close()	:	response.end
end if

if vxSiteOrdercount>1 then
	If session("ssBctID")<>"kjy8517" Then
		response.write "<script type='text/javascript'>alert('�ټ����� üũ ���.'); self.close();</script>"
		dbget.close()	:	response.end		
	End If
end if
%>
<script type='text/javascript'>

	function chorderedit(){
	    var frm=document.frm;
/*
	    alert(frm.Org_OrderName.value.length);
	    alert(frm.OrderName.value.length);
*/
		if(frm.Org_OrderName.value.length <= 2 ){
		    if(frm.Org_OrderName.value.substring(0,1)!=frm.OrderName.value.substring(0,1)){
		        alert('���� �ֹ��ΰ� ��ġ���� �ʽ��ϴ�.');
		        frm.OrderName.focus();
		        return;
		    }
		}else{
		    if(frm.Org_OrderName.value.substring(0,2)!=frm.OrderName.value.substring(0,2)){
		        alert('���� �ֹ��ΰ� ��ġ���� �ʽ��ϴ�.');
		        frm.OrderName.focus();
		        return;
		    }
		}

	    if(frm.Org_ReceiveName.value.length <= 2 ){
		    if(frm.Org_ReceiveName.value.substring(0,1)!=frm.ReceiveName.value.substring(0,1)){
		        alert('���� �����ΰ� ��ġ���� �ʽ��ϴ�.');
		        frm.ReceiveName.focus();
		        return;
		    }
	    }else{
		    if(frm.Org_ReceiveName.value.substring(0,2)!=frm.ReceiveName.value.substring(0,2)){
		        alert('���� �����ΰ� ��ġ���� �ʽ��ϴ�.');
		        frm.ReceiveName.focus();
		        return;
		    }
		}
/*
		if(frm.Org_OrderName.value.length !=  frm.OrderName.value.length){
			alert('���� �ֹ��ΰ� ������ �̸��� ���̰� �ٸ��ϴ�');
			frm.OrderName.focus();
			return;
		}

		if(frm.Org_ReceiveName.value.length !=  frm.ReceiveName.value.length){
			alert('���� �����ΰ� ������ �̸��� ���̰� �ٸ��ϴ�');
			frm.ReceiveName.focus();
			return;
		}
*/

		if(confirm("[�����ֹ���ȣ : <%= voutmallorderserial %>]\n���� �ֹ����� <%= vxSiteOrdercount %>�� ���� �մϴ�. �����Ͻðڽ��ϱ�?")){
			frm.mode.value="orderedit"
			frm.submit();
		}
	}
	
</script>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			�� ���� �ֹ����� �� <%= vxSiteOrdercount %>�� �ֽ��ϴ�.
		</td>
		<td align="right">		
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<form name="frm" method="post" action="/admin/etc/orderinput/xSiteOrderprocess.asp">
<input type="hidden" name="mode">
<input type="hidden" name="outmallorderseq" value="<%=voutmallorderseq%>">
<input type="hidden" name="outmallorderserial" value="<%=voutmallorderserial%>">
<tr bgcolor="#FFFFFF">
	<td align="center">�Ǹż��θ�</td>
	<td colspan="2" >
		<%= oorderedit.foneitem.fsellsite %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">�����ֹ���ȣ</td>
	<td colspan="2"  >
		<%= oorderedit.foneitem.fOutMallOrderSerial %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center"></td>
	<td>�ֹ���</td>
	<td>������</td>	
	
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">����</td>
	<td><%= oorderedit.foneitem.fOrderName %></td>
	<td><%= oorderedit.foneitem.fReceiveName %></td>	
	
</tr>
<tr bgcolor="#FFFFFF">
	<td align="center">����</td>
	<td>
	    <input type="hidden" name="Org_OrderName" value="<%= oorderedit.foneitem.fOrderName %>">
	    <input type="text" name="OrderName" value="<%= oorderedit.foneitem.fOrderName %>" size="10">
	</td>
	<td>
	    <input type="hidden" name="Org_ReceiveName" value="<%= oorderedit.foneitem.fReceiveName %>">
	    <input type="text" name="ReceiveName" value="<%= oorderedit.foneitem.fReceiveName %>" size="10">
	</td>	
	
</tr>


<tr bgcolor="#FFFFFF">
	<td align="center" colspan="3">
		<input type="button" value="����" onclick="chorderedit();" class="button" >
	</td>
</tr>
</form>
</table>

<%
set oorderedit = nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->