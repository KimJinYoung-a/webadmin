<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->  
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"--> 
<!-- #include virtual="/lib/classes/linkedERP/bizSectionCls.asp"-->
<!-- #include virtual="/lib/classes/analysis/bizMonthProfitCls.asp"-->
<%    
Dim intY, intM, sYear, sMonth
Dim icolType, irowType, sGRP_YN,accgroupcd
Dim oldacc, sumM, intI, intC
Dim accType, oldgrpcd,oldgrp, oldup, oldupcd, sumgrp, sumup,intSum
Dim dSDate, dEDate
Dim sumG2, arr210(), arr220(), arr230(), arr240(), arr250() 
Dim oldBizType
'������ �ޱ�
Dim bizSecCd : bizSecCd=requestCheckvar(request("bizSecCd"),16)
Dim accusecd  : accusecd=requestCheckvar(request("accusecd"),16)
Dim isINTrans  : isINTrans=requestCheckvar(request("isINTrans"),10) ''���ΰŷ� 
sYear = requestCheckvar(request("selY"),4)  
sMonth = requestCheckvar(request("selM"),2)  
icolType = requestCheckvar(request("rdoC"),1)  '����ɼ� ����(1:����κ�,2:����)
irowType = requestCheckvar(request("rdoR"),1)  '����ɼ� ����(1:�����׷�,2:�����з�,3:��������)
accgroupcd	=  requestCheckvar(request("rdoAGC"),3)

'�ʱⰪ ����
IF sYear ="" THEN sYear = year(date)
IF sMonth ="" THEN sMonth = month(date) 
	dSDate = dateserial(sYear,sMonth,"1")
 	dEDate = dateadd("d",-1,dateadd("m",1,dSDate))
IF icolType = "" THEN icolType =1  
IF irowType = "" THEN irowType =1
IF icolType ="1" THEN  
	sGRP_YN="Y"
ELSE
	sGRP_YN="N"
END IF	
IF accgroupcd ="" THEN accgroupcd = 0	
	
''����ι�
Dim clsBS, arrBizList, intL 
Set clsBS = new CBizSection   
	clsBS.FGRP_YN = sGRP_YN 
	clsBS.FYYYYMM = sYear&"-"& format00(2,sMonth) 
	arrBizList = clsBS.fnGetBizMonthProftist  
Set clsBS = nothing

'���ͺ�������Ʈ			
Dim clsBP, arrList, intLoop
Set clsBP = new CBizProfit
	clsBP.FYYYYMM =  sYear&"-"& format00(2,sMonth) 
	clsBP.Faccusecd = accusecd
	clsBP.FbizType = isINTrans
	clsBP.FAccGrpCd	=  accgroupcd  
	clsBP.FcolType = icolType
	clsBP.FrowType = irowType
	arrList = clsBP.fnGetBizMonthProfitList  
Set clsBP = nothing


 
'�����(arrPBiz), ����(arrBiz) ���� ������ ���� �� ���� �ҷ�����  
Dim arrPBiz(),arrBiz(),intB,intP, intChk, oldPCD, sBizcd,ichkNull,arrgrp(),arrup() , intSY, intSN
	intB = 0
	intP = 0
	intChk = 0
	intSY = 0 
	intSN = 0 
 IF isArray(arrBizList) THEN 
		For intLoop = 0 To UBound(arrBizList,2)   
		IF icolType ="1" THEN
			redim preserve arrBiz(2,intLoop) 
			arrBiz(1,intLoop) =  arrBizList(1,intLoop) 
			arrBiz(0,intLoop) =  arrBizList(0,intLoop)
			arrBiz(2,intLoop) =  arrBizList(4,intLoop)
			intP = intLoop 
		ELSE	
			IF oldPCD <> arrBizList(2,intLoop) THEN
				intP = intP + 1
				redim preserve arrPBiz(2,intP)
				arrPBiz(1,intP) =  arrBizList(3,intLoop)
				arrPBiz(2,intP) =  arrBizList(2,intLoop)
				IF intP> 1 THEN 
				arrPBiz(0,intP-1) = intChk
				END IF
				intChk =0
			 END IF
			 
				intChk = intChk + 1		'���� �� ������ ���� ���� ����μ� colspan ���� üũ
				
				redim preserve arrBiz(2,intLoop)
				arrBiz(1,intLoop) = arrBizList(1,intLoop)  
				arrBiz(0,intLoop) = arrBizList(0,intLoop)
				arrBiz(2,intLoop) =  arrBizList(4,intLoop)
			IF intLoop = UBound(arrBizList,2)    THEN
					arrPBiz(0,intP) = intChk
			END IF
			oldPCD  = arrBizList(2,intLoop)
		END IF
		if  arrBizList(4,intLoop)  then
			intSY = intSY + 1
		ELSE 
			intSN = intSN + 1
		end if	
		Next 
	END IF 

	 
%>
 <script type="text/javascript">
 	
 	function jsSearch(){
 		document.frm.submit();
 	}
 	
 	function jsFillCal(val1, val2){   
 		for(i=0;i<document.all.selY.length;i++){
	    if(document.all.selY.options[i].value == val1){ 
	    	document.all.selY.options[i].selected = true;
	    }  
   }
  
    if(document.all.selM.options[parseInt(val2)-1].value == val2){ 
    	document.all.selM.options[parseInt(val2)-1].selected = true;
    }
}


function showProfitDetail(bizSecCd,accusecd,biztype){ 
    var iURI ;
    if( biztype=="2"){//���ΰŷ�
    	iURI = "/admin/approval/innerorder/innerOrderList.asp?research=on&page=&yyyy1=<%=sYear%>&mm1=<%=sMonth%>&yyyy2=<%=sYear%>&mm2=<%=sMonth%>&bizsection_cd="+bizSecCd;
  	}else{
  		iURI = "/admin/analysis/popBizProfitDetail.asp?dSDate=<%=dSDate%>&dEDate=<%=dEDate%>&bizSecCd="+bizSecCd+"&accusecd="+accusecd+"&isINTrans=<%=isINTrans%>";
  	}
    var popwin = window.open(iURI,'showProfitDetail','scrollbars=yes,resizable=yes,width=900,height=600');
    popwin.focus();
}
function jsUpdateMPReport(){
		document.frmU.submit();
}
 </script>
 <!-- update ó��-->
 <form name="frmU" method="post" action="procBizMonthProfit.asp">
 	<input type="hidden" name="hidM" value="R">
	<input type="hidden" name="hidYM" value="<%= sYear&"-"& format00(2,sMonth) %>"> 
</form>		
<!-- // update ó��-->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
	<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="page" value=""> 
			<tr align="center" bgcolor="#FFFFFF" >
				<td  width="100" height="30" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
				<td align="left"> 
				��¥:
					<select name="selY" id="selY">
						<%For intY = Year(date) To 2012 STEP -1%>
						<option value="<%=intY%>" <%IF Cstr(sYear) = Cstr(intY) THEN%>selected<%END IF%>><%=intY%></option>
						<%Next%>
					</select>
						<select name="selM">
						<%For intM = 1 To 12%>
						<option value="<%=intM%>" <%IF Cstr(sMonth) = Cstr(intM) THEN%>selected<%END IF%>><%=intM%></option>
						<%Next%>
					</select> 
					<input type="button" value="������" class="button" onClick="jsFillCal('<%= Year(DateAdd("m",-2,now()))%>','<%= Month(DateAdd("m",-2,now()))%>')";>
					<input type="button" value="����" class="button" onClick="jsFillCal('<%= Year(DateAdd("m",-1,now()))%>','<%= Month(DateAdd("m",-1,now()))%>')";>
					<input type="button" value="�̹���" class="button" onClick="jsFillCal('<%= Year(DateAdd("m",0,now()))%>','<%= Month(DateAdd("m",0,now()))%>')";>
					
					&nbsp;&nbsp;
					<input type="checkbox" name="isINTrans" value="2" <%= ChkIIF(isINTrans="2","checked","") %> > ���ΰŷ���
					&nbsp;&nbsp;
					  ���������ڵ�:
					<input type="text" name="accusecd" value="<%=accusecd%>" size="15"> 
				</td>
				<td    width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="jsSearch();">
				</td>
			</tr> 
			<tr align="center" bgcolor="#FFFFFF" >
				<td  width="100" height="30" bgcolor="<%= adminColor("gray") %>">���� �ɼ�</td>
				<td align="left" colspan="2">  
					����: <input type="radio" name="rdoC" value="1" <%IF icolType=1 THEN%>checked<%END IF%>  onClick="jsSearch();">����κ� <input type="radio" name="rdoC" value="2" <%IF icolType=2 THEN%>checked<%END IF%>  onClick="jsSearch();">���� &nbsp;&nbsp;&nbsp;
					����: <input type="radio" name="rdoR" value="1" <%IF irowType=1 THEN%>checked<%END IF%>  onClick="jsSearch();">�����׷�  <input type="radio" name="rdoR" value="2" <%IF irowType=2 THEN%>checked<%END IF%> onClick="jsSearch();">�����з� <input type="radio" name="rdoR" value="3" <%IF irowType=3 THEN%>checked<%END IF%> onClick="jsSearch();">��������
					 &nbsp;&nbsp;&nbsp;
					�����׷�:
					<input type="radio" name="rdoAGC" value="100" <%IF accgroupcd=100 THEN%>checked<%END IF%>  onClick="jsSearch();">�ڻ�
					<input type="radio" name="rdoAGC" value="200" <%IF accgroupcd=200 THEN%>checked<%END IF%>  onClick="jsSearch();">100����
					<input type="radio" name="rdoAGC" value="300" <%IF accgroupcd=300 THEN%>checked<%END IF%>  onClick="jsSearch();">����
					<input type="radio" name="rdoAGC" value="0" <%IF accgroupcd=0 THEN%>checked<%END IF%>  onClick="jsSearch();">��ü
					
					==&gt;������� (A:�ڻ�,B:��ä,C:�ں�,D:����) <!-- tbl_TMS_SL_ACC_CD_GRP -->
				</td> 
			</tr>  
			</form>
		</table>
	</td>
</tr> 
<tr>
	<td><input type="button" class="button" value="update" onClick="jsUpdateMPReport();"> : �˻������� ��¥�� �ش��ϴ� �����͸� ������Ʈ �˴ϴ�.</td>
</tr> 
<tr>
	<td>
		<!-- ��� �� ���� -->
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" border="0">   
	<%	IF icolType="1" THEN '--1-1.����:����κ����� -�̸� ������ �������� �ҷ�����%>
		<tr  bgcolor="<%= adminColor("tabletop") %>"  align="center"> 
			<%IF irowType ="1" THEN%>
				<td rowspan="2">�����׷��ڵ�</td> 
				<td colspan="5"  rowspan="2">�����׷��</td> 
			<%ELSEIF irowType="2" THEN%>
				<td  rowspan="2">�����׷��ڵ�</td> 
				<td  rowspan="2">�����׷��</td> 
				<td  rowspan="2">�����з��ڵ�</td> 
				<td  colspan="3"  rowspan="2">�����з���</td>  
			<%ELSE%>
				<td  rowspan="2">�����׷��ڵ�</td> 
				<td  rowspan="2">�����׷��</td> 
				<td  rowspan="2">�����з��ڵ�</td> 
				<td  rowspan="2">�����ڵ�</td> 
				<td  rowspan="2">���������ڵ�</td> 
				<td  rowspan="2">������</td>  
			<%END IF%>
				<td  rowspan="2">NULL </td>
			<% IF isArray(arrBizList) THEN
			 
			 %> 
				<td <%if intSY>1 then%>colspan="<%=intSY%>" <%end if%>>���ͺμ�</td>  
				<td  <%if intSN>1 then%>colspan="<%=intSN%>" <%end if%>>�����μ�</td> 
			<% 
		END IF
			%>
			<td rowspan="2">�հ�</td> 
		</tr>
		<tr bgcolor="<%= adminColor("tabletop") %>"  align="center">			
			<% IF isArray(arrBizList) THEN
			For intLoop=0 To UBound(arrBizList,2) %> 
				<td><%=arrBiz(1,intLoop)%></td> 
			<%Next
		END IF
			%> 
			
		</tr> 
	<%	ELSE '--1-2����:�������� %> 
		<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			<%IF irowType ="1" THEN%>
				<td rowspan="3">�����׷��ڵ�</td> 
				<td rowspan="3" colspan="5">�����׷��</td> 
			<%ELSEIF irowType="2" THEN%>
				<td rowspan="3">�����׷��ڵ�</td> 
				<td rowspan="3">�����׷��</td> 
				<td rowspan="3">�����з��ڵ�</td>  
				<td rowspan="3" colspan="3">�����з���</td>  
			<%ELSE%>
				<td rowspan="3">�����׷��ڵ�</td> 
				<td rowspan="3">�����׷��</td> 
				<td rowspan="3">�����з��ڵ�</td> 
				<td rowspan="3">�����ڵ�</td> 
				<td rowspan="3">���������ڵ�</td> 
				<td rowspan="3">������</td>  
			<%END IF%>
				<td rowspan="3">NULL </td>
			<% IF isArray(arrBizList) THEN
			 
			%> 
				<td <%if intSY>1 then%>colspan="<%=intSY%>" <%end if%>>���ͺμ�</td>  
				<td  <%if intSN>1 then%>colspan="<%=intSN%>" <%end if%>>�����μ�</td> 
			<% 
				END IF
			%>
				<td rowspan="3" >�հ�</td>
		</tr>   
		<tr  bgcolor="<%= adminColor("tabletop") %>"  align="center">
			<%IF isArray(arrPBiz) THEN
				For intLoop = 0 To intP%>
				<td colspan="<%=arrPBiz(0,intLoop)%>"><%=arrPBiz(1,intLoop)%></td>
			<%Next
				END IF
			%>  
		</tr>
			<tr  bgcolor="<%= adminColor("tabletop") %>"  align="center">   	
			<%IF isArray(arrBizList) THEN
			For intLoop=0 To UBound(arrBizList,2) %> 
				<td><%=arrBiz(1,intLoop)%></td> 
			<%Next
		END IF
			%> 
		</tr> 
	<%	END IF %> 
	<% 	intI = 0	
	sumgrp = 0
	sumup = 0 
	IF isArray(arrBizList) THEN 
		redim arrgrp(UBound(arrBizList,2)+1)
		redim arrup(UBound(arrBizList,2)+1)
		redim arr210(UBound(arrBizList,2)+1)
		redim arr220(UBound(arrBizList,2)+1)
		redim arr230(UBound(arrBizList,2)+1)
		redim arr240(UBound(arrBizList,2)+1)
		redim arr250(UBound(arrBizList,2)+1)
	END IF
	Dim chkUpCd
	IF isArray(arrList) THEN 
			For intLoop = 0 To UBound(arrList,2)
				ichkNull =0
				'//���� �ɼǿ� ���� ���̺� �ٹٲ� ��� ����
				IF  irowType = "1" THEN
					accType =arrList(0,intLoop) '�����׷�
					chkUpCd =""
				ELSEIF irowType = "2" THEN
					accType =arrList(7,intLoop) '�����з�
					chkUpCd =arrList(6,intLoop)
				ELSE
					accType =arrList(8,intLoop)	'��������
					chkUpCd =arrList(6,intLoop)
				END IF 
				IF oldacc <> accType or (irowType="2" and oldupcd <> chkUpCd ) or oldbizType <> arrList(5,intLoop) THEN	'���������� Ʋ������ table�ٹٲ�   
					IF intLoop > 0 THEN
						For intC= intI To UBound(arrBizList,2) '
				%>
				<td align="right">0</td>
			<%		Next
					intI = 0	
			%>
				<td  align="right"><%=formatnumber(sumM,0)%></td>
			</tr>
			<% IF irowType > "2"  then 
						if oldupcd <> arrList(6,intLoop) THEN
							IF oldupcd <> "" THEN 
							sumup = 0 
				%>
				<tr bgcolor="#EEEEEE"> 
				<td colspan="6" align="center"><%=oldup%></td> 
					<%For intSum = 0 To UBound(arrBizList,2)+1
						sumup = sumup + arrup(intSum)
				%>
				<td  align="right"><%=formatnumber(arrup(intSum),0)%></td> 
				<%Next%>
					<td  align="right"><%=formatnumber(sumup,0)%></td> 
				</tr>
				<%	 	
						END IF
						redim arrup(UBound(arrBizList,2)+1)
						END IF
					END IF 
				%>	
				<% IF oldgrpcd <> arrList(0,intLoop) THEN  
						IF irowType > "1" THEN  
							sumgrp = 0 
				%>
				<tr bgcolor="#DDFFDD"> 
				<td colspan="6" align="center" style="border-bottom:2px solid <%=adminColor("tablebg")%>;"><%=oldgrp%></td> 
				<%For intSum = 0 To UBound(arrBizList,2)+1
						sumgrp = sumgrp + arrgrp(intSum)
				%>
				<td  align="right"  style="border-bottom:2px solid <%=adminColor("tablebg")%>;"><%=formatnumber(arrgrp(intSum),0)%></td> 
				<%Next%>
					<td  align="right"  style="border-bottom:2px solid <%=adminColor("tablebg")%>;"><%=formatnumber(sumgrp,0)%></td> 
				</tr>  
				<%	redim arrgrp(UBound(arrBizList,2)+1)
				END IF%> 
				<% sumG2 = 0
				IF oldgrpcd = "220" THEN %>
				<tr bgcolor="#FFDDDD">
					<td colspan="6" align="center">�������</td>
					<%For intSum = 0 To UBound(arrBizList,2)+1 
							sumG2 = sumG2+arr210(intSum)+arr220(intSum)
					%>
					<td align="right"><%=formatnumber(arr210(intSum)+arr220(intSum),0)%></td> 
					<%Next%>
					<td align="right"><%=formatnumber(sumG2,0)%></td>
				</tr>
				<%	
					 ELSEIF oldgrpcd = "230" THEN 
				%>
				<tr bgcolor="#FFDDDD">
					<td  colspan="6"  align="center">��������</td>
					<%For intSum = 0 To UBound(arrBizList,2)+1 
							sumG2 = sumG2+arr230(intSum)
					%>
					<td align="right"><%=formatnumber(arr210(intSum)+arr220(intSum)+arr230(intSum),0)%></td> 
					<%Next%>
							<td align="right"><%=formatnumber(sumG2,0)%></td>
				</tr>
			<%  ELSEIF oldgrpcd = "250" THEN 
				%>
				<tr bgcolor="#FFDDDD">
					<td colspan="6" align="center">������</td>
					<%For intSum = 0 To UBound(arrBizList,2)+1 
						sumG2 = sumG2+arr230(intSum)+arr240(intSum)+arr250(intSum)
					%>
					<td align="right"><%=formatnumber(arr210(intSum)+arr220(intSum)+arr230(intSum)+arr240(intSum)+arr250(intSum),0)%></td> 
					<%Next%>
					<td align="right"><%=formatnumber(sumG2,0)%></td>
				</tr>
				<% 
				END IF
				%>	
				<% 
				END IF
				%>	
			
			<% 	END IF%>
			<tr align="center" bgcolor="#FFFFFF">
				<!--��������-->
				<%IF irowType = "1" THEN%>
				<td><%=arrList(0,intLoop)%></td>
				<td colspan="5"><%=arrList(1,intLoop)%></td> 
				<%ELSEIF irowType = "2" THEN%>
				<td><%=arrList(0,intLoop)%></td>
				<td><%=arrList(1,intLoop)%></td> 
				<td><%=arrList(6,intLoop)%></td> 
				<td  colspan="3"><%=arrList(7,intLoop)%></td> 
				<%ELSE%>
				<td><%=arrList(0,intLoop)%></td>
				<td><%=arrList(1,intLoop)%></td>
				<td><%=arrList(6,intLoop)%></td> 
				<td><%=arrList(8,intLoop)%></td>
				<td><%=arrList(9,intLoop)%></td>
				<td><%=arrList(10,intLoop)%><%IF arrList(5,intLoop) ="2" THEN%><font color="blue">(���ΰŷ�)</font><%END IF%></td> 
				<%END IF%>
				<!--/��������-->
			<%	sumM = 0 
					IF isNull(arrList(2,intLoop)) or arrList(2,intLoop) ="" THEN	'����μ� Null �϶�
					sumM = sumM + arrList(4,intLoop)-arrList(3,intLoop)
					 arrgrp(0) = arrgrp(0) + arrList(4,intLoop)-arrList(3,intLoop)
					 	arrup(0) = arrup(0)+ arrList(4,intLoop)-arrList(3,intLoop)
					 	IF arrList(0,intLoop) = "210" THEN 
					 		arr210(0) = arrgrp(0)
					 	ELSEIF arrList(0,intLoop) = "220" THEN 
					 		arr220(0) = arrgrp(0)
					 	ELSEIF arrList(0,intLoop) = "230" THEN 
					 		arr230(0) = arrgrp(0)
					 	ELSEIF arrList(0,intLoop) = "240" THEN
					 		 arr240(0) = arrgrp(0)
					 	ELSEIF arrList(0,intLoop) = "250" THEN 
					 		arr250(0) = arrgrp(0)
						END IF
			%>	
				<td  align="right"><%IF irowType="3" THEN%><a href="javascript:showProfitDetail('','<%=arrList(9,intLoop)%>','<%=arrList(5,intLoop)%>');"><%END IF%>
				<%=formatnumber(arrList(4,intLoop)-arrList(3,intLoop),0)%>
				</a></td>		
		<%   	ichkNull = 1
				ELSE %>
				<td align="right">0</td>
		<%	END IF
			END IF 
			IF ichkNull = 0 THEN
			For intC = intI To  UBound(arrBizList,2)  '����μ� ����ŭ ����
					IF arrBiz(0,intC)  = arrList(2,intLoop)  THEN '����μ��� ���͵�� �μ��� �����Ҷ� �� �����ش�. 
					 	sumM = sumM + arrList(4,intLoop)-arrList(3,intLoop)
					 	arrgrp(intC+1) = arrgrp(intC+1) + arrList(4,intLoop)-arrList(3,intLoop) 
					 	arrup(intC+1) = arrup(intC+1)+ arrList(4,intLoop)-arrList(3,intLoop)
					 	IF arrList(0,intLoop) = "210" THEN
					 		 arr210(intC+1) = arrList(4,intLoop)-arrList(3,intLoop)  
					 	ELSEIF arrList(0,intLoop) = "220" THEN 
					 		arr220(intC+1) = arrList(4,intLoop)-arrList(3,intLoop) 
					 	ELSEIF arrList(0,intLoop) = "230" THEN 
					 		arr230(intC+1) = arrList(4,intLoop)-arrList(3,intLoop) 
					 	ELSEIF arrList(0,intLoop) = "240" THEN 
					 		arr240(intC+1) = arrList(4,intLoop)-arrList(3,intLoop) 
					 	ELSEIF arrList(0,intLoop) = "250" THEN 
					 		arr250(intC+1) = arrList(4,intLoop)-arrList(3,intLoop) 
						END IF
				%>
				<td align="right"><%IF irowType="3" THEN%><a href="javascript:showProfitDetail('<%=arrList(2,intLoop)%>','<%=arrList(9,intLoop)%>','<%=arrList(5,intLoop)%>');"><%END IF%><%=formatnumber(arrList(4,intLoop)-arrList(3,intLoop),0)%></a></td>
				<%	 intI = intC+1
						Exit For  
					ELSE	'���͵�� �� ������ 0 �ѷ��ش�.
				%>
				<td align="right">0</td>
				<%		
					END IF
			Next%> 
		<%END IF	
				oldBizType = arrList(5,intLoop)
				oldacc = accType 
				oldgrp = arrList(1,intLoop)
				oldgrpcd = arrList(0,intLoop)
				if irowtype > 1 then	
					oldupcd = arrList(6,intLoop) 
					oldup = arrList(7,intLoop)	 
				end if
			Next  
			For intC= intI To UBound(arrBizList,2) 
		%>
				<td align="right">0</td>
		<%Next %>  
				<td  align="right"><%=formatnumber(sumM,0)%></td>
			</tr>
			<%  
			IF irowType > "2"  then
						if  oldup <> "" THEN
							sumup = 0
				%>
				<tr bgcolor="#EEEEEE"> 
				<td colspan="6" align="center"><%=oldup%></td> 
					<%For intSum = 0 To UBound(arrBizList,2)+1
						sumup = sumup + arrup(intSum)
				%>
				<td  align="right"><%=formatnumber(arrup(intSum),0)%></td> 
				<%Next%>
					<td  align="right"><%=formatnumber(sumup,0)%></td> 
				</tr>
				<%	
				END IF
					END IF
				%>	
				<% sumgrp = 0
				 IF irowType >  "1"  THEN 
				%>
				<tr bgcolor="#DDFFDD"> 
				<td colspan="6" align="center" style="border-bottom:2px solid <%=adminColor("tablebg")%>;"><%=oldgrp%></td> 
				<%For intSum = 0 To UBound(arrBizList,2)+1
						sumgrp = sumgrp + arrgrp(intSum)
				%>
				<td  align="right"  style="border-bottom:2px solid <%=adminColor("tablebg")%>;"><%=formatnumber(arrgrp(intSum),0)%></td> 
				<%Next%>
					<td  align="right"  style="border-bottom:2px solid <%=adminColor("tablebg")%>;"><%=formatnumber(sumgrp,0)%></td> 
				</tr>  
				<%	
				END IF
				%>	
				<% sumG2 = 0
				IF oldgrpcd = "220" THEN %>
				<tr bgcolor="#FFDDDD">
					<td colspan="6" align="center">�������</td>
					<%For intSum = 0 To UBound(arrBizList,2)+1 
							sumG2 = sumG2+arr210(intSum)+arr220(intSum)
					%>
					<td align="right"><%=formatnumber(arr210(intSum)+arr220(intSum),0)%></td> 
					<%Next%>
					<td align="right"><%=formatnumber(sumG2,0)%></td>
				</tr>
				<%	
					 ELSEIF oldgrpcd = "230" THEN
				%>
				<tr bgcolor="#FFDDDD">
					<td colspan="6" align="center">��������</td>
					<%For intSum = 0 To UBound(arrBizList,2)+1  
					 		sumG2 = sumG2+arr230(intSum)
					%>
					<td align="right"><%=formatnumber(arr210(intSum)+arr220(intSum)+arr230(intSum),0)%></td> 
					<%Next%>
							<td align="right"><%=formatnumber(sumG2,0)%></td>
				</tr>
			<%  ELSEIF oldgrpcd = "250" THEN 
				%>
				<tr bgcolor="#FFDDDD">
					<td colspan="6" align="center">������</td>
					<%For intSum = 0 To UBound(arrBizList,2)+1
					sumG2 = sumG2+arr230(intSum)+arr240(intSum)+arr250(intSum)
					 %>
					<td align="right"><%=formatnumber(arr210(intSum)+arr220(intSum)+arr230(intSum)+arr240(intSum)+arr250(intSum),0)%></td> 
					<%Next%>
					<td align="right"><%=formatnumber(sumG2,0)%></td>
				</tr>
				<% 
				END IF
				%>	
		<%ELSE%>
		<tR>
			<td bgcolor="#FFFFFF" colspan="35" align="center">��ϵ� ������ �����ϴ�.</td>
		</tr>	
		<%END IF%>
		</table>
	</td>
</tr>
</table>	