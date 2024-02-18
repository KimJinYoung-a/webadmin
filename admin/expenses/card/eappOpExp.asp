<%@ language="VBScript" %>
<% option explicit %>
 
<%
'###########################################################
' Description : ������  ����
' History : 2011.05.30 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpCardCls.asp"-->
<%
 Dim clsOpExp, arrList, intLoop
 Dim iPartTypeIdx,iOpExpPartIdx,dyyyymm,sOpExpPartName, sPartTypeName
 Dim iOpExpidx, mLastMonthExp, mInExp, mOutExp, mTotExp 
 Dim returnValue,objCmd
 
 dyyyymm =  requestCheckvar(Request("dyyyymm"),7) 
 iOpExpPartIdx =  requestCheckvar(Request("hidP"),10)
  iPartTypeIdx=  requestCheckvar(Request("hidPT"),10)
Set clsOpExp = new OpExp 
'��� ����Ʈ	
	clsOpExp.FYYYYMM	= dyyyymm 
	clsOpExp.FPartTypeIdx	=iPartTypeIdx
	clsOpExp.FOpExpPartIdx	=iOpExpPartIdx
	clsOpExp.Farap_cd	=0 
	clsOpExp.fnGetOpExpMonthlyData
	iOpExpidx = clsOpExp.FOpExpidx 
	mOutExp	= clsOpExp.FOutExp 
	sOpExpPartName = clsOpExp.FOpExpPartName
	sPartTypeName = clsOpExp.FPartTypeName
	arrList = clsOpExp.fnGetOpExpDailySumList 
Set clsOpExp = nothing	 


'������ �������������� ���� 
Set objCmd = Server.CreateObject("ADODB.COMMAND")   
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText  		
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_OpExpMonthlyCard_setConfirm]("&iOpExpIdx&",5)}"							 
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With	
		    returnValue = objCmd(0).Value	  
Set objCmd = nothing	
%>  
<!-- #include virtual="/lib/db/dbclose.asp" -->  


<div id="divEapp" style="display:none;">
<table width="500" align="center" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td><%=dyyyymm%>&nbsp;<%=sPartTypeName%> > <%=sOpExpPartName%>&nbsp;����ī���볻�� </td>
</tr>  
<tr>
	<td> 
		<table width="500" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<td>����</td>
			<td>����</td>
			<td>�Ǽ�</td>
		</tr>
		<%
		Dim sumExp, sumCnt
		sumExp = 0
		sumCnt = 0
		IF isArray(arrList) THEN
			For intLoop = 0 To UBound(arrList,2) 
			%>
		<tr  align="center" bgcolor="#FFFFFF">
			<td><%=arrList(6,intLoop)%></td>
			<td><%=formatnumber(arrList(0,intLoop),0)%></td>
			<td><%=formatnumber(arrList(4,intLoop),0)%></td>
		</tr>
		<%  sumExp = sumExp + arrList(0,intLoop)
			sumCnt	= sumCnt + 	arrList(4,intLoop) 
			Next
		END IF%>
		<tr align="center" bgcolor="#FFFFFF">
			<td>�հ�</td>
			<td><%=formatnumber(sumExp,0)%></td>
			<td><%=formatnumber(sumCnt,0)%></td>
		</tr>
		</table>
	</td>
</tr>
</table>
</div>
<form name="frmEapp" method="post" action="/admin/approval/eapp/regeapp.asp">
<input type="hidden" name="iSL" value="<%=iOpExpidx%>">
<input type="hidden" name="tC" value="">
<input type="hidden" name="mRP" value="<%=sumExp%>">
<input type="hidden" name="ieidx" value="10"> <!-- ������ȣ ����!! -->
</form>

<script language="javascript">
<!--
	//���ڰ��� ǰ�Ǽ� ��� 
	  	document.frmEapp.tC.value = document.all.divEapp.innerHTML.replace(/\r|\n/g,"");  
	  	document.frmEapp.submit(); 
	//--> 
</script>
 