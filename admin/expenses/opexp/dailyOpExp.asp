<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������ ��   ����Ʈ
' History : 2011.06.03 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"--> 
<!-- #include virtual="/lib/classes/expenses/OpExpCls.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpPartCls.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpAccountCls.asp"-->
<!-- #include virtual="/lib/classes/approval/partMoneyCls.asp"-->
<%
Dim clsPart,clsOpExp, arrPart, arrList, arrType, intLoop, clsPartMoney
Dim clsAccount, arrAccount ,iarap_cd
Dim  arrUsePart ,sOpExpPartName, sPartTypeName
Dim dYear, dMonth, iPartTypeIdx, iOpExpPartIdx	,sbizsection_cd,sbizsection_nm
Dim intY, intM
Dim iTotCnt,iPageSize, iTotalPage,iCurrPage
Dim blnAdmin, blnWorker, blnReg 
 
	iPageSize = 20
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1
		 
 	iPartTypeIdx	= requestCheckvar(Request("selPT"),10)
 	iOpExpPartIdx	= requestCheckvar(Request("selP"),10)
 	dYear			=  requestCheckvar(Request("selY"),4)
 	dMonth			=  requestCheckvar(Request("selM"),2)
 	iarap_cd		= requestCheckvar(Request("selA"),10)
 	sbizsection_nm=requestCheckvar(Request("sBiznm"),100)
 	IF dYear = "" THEN dYear = year(date())
 	IF dMonth = "" THEN dMonth = month(date())	
 	IF iPartTypeIdx = "" THEN iPartTypeIdx = 0
 		
 	'�����ʱⰪ ����-------------- 
 	blnWorker = 0 '�����
 	blnReg = 0 	'��ϱ���
 	 
  	blnAdmin = fnChkAdminAuth(session("ssAdminLsn"),session("ssAdminPsn"))  '���α���	
  	
  	IF blnAdmin THEN blnReg = 1 '���α��� ���� ��� ���ó�� �׻� ���� 
 	 
 '������ �� ���� ����Ʈ		
Set clsPart = new COpExpPart
	arrType = clsPart.fnGetOpExpPartTypeList 
	IF iPartTypeIdx > 0 THEN
	clsPart.FPartTypeidx 	= iPartTypeIdx  
	arrPart = clsPart.fnGetOpExppartAllList  
	END IF
	
	IF iOpExpPartIdx > 0 THEN
		clsPart.FOpExpPartidx = iOpExpPartIdx
		clsPart.fnGetOpExpPartName
		sOpExpPartName =clsPart.FOpExpPartName
		sPartTypeName  =clsPart.FPartTypeName 
	END IF
Set clsPart = nothing

'���� ����Ʈ
set clsAccount = new COpExpAccount
	arrAccount = clsAccount.fnGetAccountAll
set clsAccount = nothing  

	
'��� ����Ʈ	
Set clsOpExp = new OpExp
	clsOpExp.FYYYYMM 	= dYear&"-"&Format00(2,dMonth)
	clsOpExp.FPartTypeIdx = iPartTypeIdx 
	clsOpExp.FOpExpPartIdx = iOpExpPartIdx 
	clsOpExp.Farap_cd = iarap_cd
	clsOpExp.FBizsection_nm = sbizsection_nm
	clsOpExp.FCurrPage 	= iCurrPage
	clsOpExp.FPageSize 	= iPageSize
	arrList = clsOpExp.fnGetOpExpDailyList
	iTotCnt = clsOpExp.FTotCnt 
	'����üũ----------------------------
	 
	IF iOpExpPartIdx > 0 THEN
		clsOpExp.Fyyyymm		= dYear&"-"&Format00(2,dMonth) 
		clsOpExp.FOpExpPartIdx	= iOpExpPartIdx 
		clsOpExp.FMode			= "I"
		clsOpExp.FadminID 		= session("ssBctId") 
		blnWorker = clsOpExp.fnGetOpExpAuth   
 
	 	IF blnWorker = 1   THEN	 blnReg =1 '������̰ų� ���α����� ���� ��� ���ó�� ���� 
	END IF
	'/����üũ---------------------------- 
Set clsOpExp = nothing	 

iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
 
%> 
 <script type="text/javascript" src="/js/ajax.js"></script>
<script language="javascript">
<!--
 	//�� ���� 
	// ajax =========================================================================================================
    initializeReturnFunction("processAjax()");
    initializeErrorFunction("onErrorAjax()");
    
    var _divName = "divP";
    
    function processAjax(){
        var reTxt = xmlHttp.responseText;   
        eval("document.all."+_divName).innerHTML = reTxt; 
    }
    
    function onErrorAjax() {
            alert("ERROR : " + xmlHttp.status);
    }
    
     
    function jsChangePart(iValue){   
        initializeURL('/admin/expenses/part/ajaxPart.asp?iPTIdx='+iValue);
    	startRequest();
    	
    }
	   
//���ε��
function jsNewReg(){
	var winNew = window.open("about:blank","popNew","width=1500,height=600,resizable=yes, scrollbars=yes");
	document.frm.target = "popNew";
	document.frm.action = "regOpExp.asp";
	document.frm.submit();
	winNew.focus();
}  
 
//����
function jsModOpExp(iOED){
	var winNew = window.open("regOpExp.asp?selY=<%=dyear%>&selM=<%=dmonth%>&selPT=<%=iPartTypeIdx%>&selP=<%=iOpExpPartIdx%>&hidOED="+iOED,"popNew","width=1500,height=600,resizable=yes, scrollbars=yes");
	winNew.focus(); 
} 

//����
 	function jsDelOpExp(idx){
 		if(confirm("�����Ͻðڽ��ϱ�?")){
 			document.frmDel.hidOED.value = idx;
 			document.frmDel.submit();
 		}
 	}
 	
 //�������̵�	
 	function jsGoPage(iP){
		document.frm.iCP.value = iP;
		document.frm.submit();
	}
	
	//�˻�
	function jsSearch(){
			document.frm.target = "_self";
	document.frm.action = "dailyOpExp.asp";
		document.frm.submit();
	}
	
	//����Ʈ�� �̵�
	function jsGoList(sPage){
		location.href = sPage+".asp?selSY=<%=dyear%>&selSM=<%=dmonth%>&selPT=<%=iPartTypeIdx%>&selP=<%=iOpExpPartIdx%>&menupos=<%=menupos%>";
	}
	
	//����Ʈ
	function jsPrint(){
		var winP = window.open("printDailyOpExp.asp?selY=<%=dyear%>&selM=<%=dmonth%>&selPT=<%=iPartTypeIdx%>&selP=<%=iOpExpPartIdx%>","popP","width=1024, height=600,resizable=yes, scrollbars=yes");
		winP.focus();
	}
	
 
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">  
<form name="frmDel" method="post" action="procOpExp.asp">
<input type="hidden" name="hidM" value="D">  
<input type="hidden" name="hidOED" value="">
<input type="hidden" name="selY" value="<%=dYear%>">
<input type="hidden" name="selM" value="<%=dMonth%>">
<input type="hidden" name="selPT" value="<%=ipartTypeIdx%>">
<input type="hidden" name="selP" value="<%=iOpExpPartIdx%>">
<input type="hidden" name="menupos" value="<%=menupos%>"> 
</form>
<tr>
	<td>+ <a href="javascript:jsGoList('index');">������ ���� �� ����Ʈ</a> > <a href="javascript:jsGoList('dailySumOpExp');">���� �� ����Ʈ</a> > �Ϻ� �� ����Ʈ</td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="iCP" value=""> 
			<tr align="center" bgcolor="#FFFFFF" >
				<td rowspan="2" width="100" height="50" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
				<td align="left">
				��¥:
				<select name="selY">
				<%For intY = Year(date()) To 2011 STEP -1%>
				<option value="<%=intY%>" <%IF Cstr(intY) = Cstr(dYear) THEN%>selected<%END IF%>><%=intY%></option>
				<%Next%>
				</select>��
				 <select name="selM">
				<%For intM = 1 To 12%>
				<option value="<%=intM%>" <%IF Cstr(intM) = Cstr(dMonth) THEN%>selected<%END IF%>><%=intM%></option>
				<%Next%>
				</select>��
				&nbsp;&nbsp;
				�����ó:
				<select name="selPT" onChange="jsChangePart(this.value);">
				<option value="0">--��ü--</option>
				<% sbOptPartType arrType,ipartTypeIdx%>
				</select>
				<span id="divP">
				<select name="selP" onChange="jsSearch();">
				<option value="0">--��ü--</option>
				<% sbOptPart arrPart,iOpExpPartIdx%>
				</select> 
				</span> 
				&nbsp;&nbsp;
				�����׸�:
				<select name="selA">
				<option value="0">--��ü--</option>
				<% sbOptAccount arrAccount, iarap_cd%>
				</select>
				&nbsp;&nbsp;
				���μ�:
				<input type="text" name="sBiznm" value="<%=sbizsection_nm%>" size="20"> 
				</td> 
				<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="jsSearch();">
				</td>
			</tr>
			</form>
		</table>
	</td>
</tr> 
<!-- #include virtual="/lib/db/dbclose.asp" -->  
<tr>
	<td>
		<table border="0" cellpadding="5" cellspacing="0" width="100%">
		<tr>
		<%IF blnReg = 1 THEN%>
			<td><input type="button" class="button" value="���󼼳��� ���" onClick="jsNewReg();"></td>
		<%END IF%>
		<%IF iOpExpPartIdx > 0 THEN%>
			<td align="right"><input type="button" class="button" value="����Ʈ" onClick="jsPrint();"></td>
		<%END IF%>	
		</tr>
		</table>
	</td>
</tr> 
<tr>
	<td>
		<!-- ��� �� ���� -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
		<tr>
			<td>
				<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
				<form name="frmReg" method="post" action="regOpExp.asp">
				<input type="hidden" name="menupos" value="<%=menupos%>">
				<input type="hidden" name="selY" value="<%=dYear%>">
				<input type="hidden" name="selM" value="<%=dMonth%>">
				<input type="hidden" name="selPT" value="<%=ipartTypeIdx%>">
				<input type="hidden" name="selP" value="<%=iOpExpPartIdx%>">
				<input type="hidden" name="iCP" value="">
					<tr align="center" bgcolor="<%= adminColor("tabletop") %>">  
						<td width="50">����</td>
						<td width="50">��¥(��)</td>  
						<td>�����ó</td>   
						<td>�����׸�</td>
						<td>��ü��</td>  
						<td>����(�󼼳���)</td>   
						<td>���޾�</td>  
						<td>����</td>   
						<td>���ް���</td> 
						<td>�ΰ���</td> 
						<td>���ι�ȣ</td> 
						<td>���μ�</td>  
						<td width="100">ó��</td>     
					</tr>
					<%   Dim totInExp, totOutExp,sumInExp,sumOutExp, iNum, sumSupExp, sumVatExp, totSupExp, totVatExp
					totInExp = 0
					totOutExp = 0
					sumInExp=0
					sumOutExp=0
					sumSupExp=0
					sumVatExp=0
					totSupExp=0
					totVatExp=0
					iNum = 1
					IF isArray(arrList) THEN
						For intLoop = 0 To UBound(arrList,2)  
					 %>  
					<tr height=30 bgcolor="#FFFFFF">	 
						<td align="center"><%=iNum%></td>
						<td align="center"><%=day(arrList(1,intLoop))%></td>
						<td align="center"><%=arrList(12,intLoop)%> > <%=arrList(11,intLoop)%></td>
						<td align="center"><%=arrList(3,intLoop)%></td>
						<td><%=arrList(6,intLoop)%></td>
						<td><%=arrList(7,intLoop)%></td> 
						<td align="right"><%=formatnumber(arrList(4,intLoop),0)%></td>
						<td align="right"><%=formatnumber(arrList(5,intLoop),0)%></td> 
						<td align="right"><%=formatnumber(arrList(8,intLoop),0)%></td> 
						<td align="right"><%=formatnumber(arrList(9,intLoop),0)%></td> 
						<td align="center"><%=arrList(10,intLoop)%></td> 
						<td align="center"><%=arrList(13,intLoop)%></td>  
						<td align="center">
						<% if IsNULL(arrList(20,intLoop)) then %>
						<%IF blnReg = 1 THEN%>
							<input type="button" class="button" value="����" onClick="jsModOpExp(<%=arrList(0,intLoop)%>);">
							<input type="button" class="button" value="����" onClick="jsDelOpExp(<%=arrList(0,intLoop)%>)">
						<%END IF%>
						<% else %>
						    <%= arrList(20,intLoop) %>
						<% end if %>
						</td>
					</tr>	
					<%
					  totInExp = totInExp + arrList(4,intLoop)
					  totOutExp = totOutExp + arrList(5,intLoop)
					  totSupExp = totSupExp + arrList(8,intLoop)
					  totVatExp = totVatExp + arrList(9,intLoop)
					  	
					  sumInExp = sumInExp +  arrList(4,intLoop)
					  sumOutExp = sumOutExp +  arrList(5,intLoop)	 
					  sumSupExp = sumSupExp +  arrList(8,intLoop)
					  sumVatExp = sumVatExp +  arrList(9,intLoop)	
					  
					  iNum = iNum + 1
				IF intLoop  < UBound(arrList,2)  THEN
					IF Cstr(arrList(2,intLoop)) <> Cstr(arrList(2,intLoop+1)) THEN%>
				   <tr height=30 align="center" bgcolor="#FFFFFF"> 
				   	<td colspan="6"><b><%=arrList(3,intLoop)%></b></td>
				   	<td align="right"><b><%=formatnumber(sumInExp,0)%></b></td>
				   	<td align="right"><b><%=formatnumber(sumOutExp,0)%></b></td>
				   	<td align="right"><%=formatnumber(sumSupExp,0)%></td>
				   	<td align="right"><%=formatnumber(sumVatExp,0)%></td>
				    <td colspan="4"></td> 
				</tr>
				<%	sumInExp = 0
					sumOutExp = 0
					sumSupExp = 0
					sumVatExp = 0
					iNum = 1
					END IF
				END IF
					Next  %>
					<tr  height=30 align="center" bgcolor="#FFFFFF"> 
				   	<td colspan="6"><b><%=arrList(3,intLoop-1)%></b></td>
				   	<td align="right"><b><%=formatnumber(sumInExp,0)%></b></td>
				   	<td align="right"><b><%=formatnumber(sumOutExp,0)%></b></td>
				   	<td align="right"><%=formatnumber(sumSupExp,0)%></td>
				   	<td align="right"><%=formatnumber(sumVatExp,0)%></td>
				   	<td colspan="4"></td> 
					</tr>
					<%
					ELSE%>
					<tr height="30" align="center" bgcolor="#FFFFFF">				
						<td colspan="14">��ϵ� ������ �����ϴ�.</td>	
					</tr>
					<%END IF%>
				<%IF iOpExpPartIdx > 0 THEN	 %>
				 <tr  height=30 align="center" bgcolor="#DDDDFF"> 
				   	<td colspan="6">����</td>
				   	<td align="right"><%=formatnumber(totInExp,0)%></td>
				   	<td align="right"><%=formatnumber(totOutExp,0)%></td>
				   	<td align="right"><%=formatnumber(totSupExp,0)%></td>
				   	<td align="right"><%=formatnumber(totVatExp,0)%></td>
				   	<td colspan="4"></td> 
				</tr>
			 <%END IF%>
				</table>	
			</td>
		</tR>	
		<!-- ������ ���� -->
		<%
		IF iOpExpPartIdx = 0 THEN
		Dim iStartPage,iEndPage,iX,iPerCnt
		iPerCnt = 10
		
		iStartPage = (Int((iCurrPage-1)/iPerCnt)*iPerCnt) + 1
		
		If (iCurrPage mod iPerCnt) = 0 Then
			iEndPage = iCurrPage
		Else
			iEndPage = iStartPage + (iPerCnt-1)
		End If
		%>
			<tr height="25" >
				<td colspan="15" align="center">
					<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
					    <tr valign="bottom" height="25">        
					        <td valign="bottom" align="center">
					         <% if (iStartPage-1 )> 0 then %><a href="javascript:jsGoPage(<%= iStartPage-1 %>)" onfocus="this.blur();">[pre]</a>
							<% else %>[pre]<% end if %>
					        <%
								for ix = iStartPage  to iEndPage
									if (ix > iTotalPage) then Exit for
									if Cint(ix) = Cint(iCurrPage) then
							%>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();"><font color="00abdf"><strong>[<%=ix%>]</strong></font></a>
							<%		else %>
								<a href="javascript:jsGoPage(<%= ix %>)" class="menu_link3" onfocus="this.blur();">[<%=ix%>]</a>
							<%
									end if
								next
							%>
					    	<% if Cint(iTotalPage) > Cint(iEndPage)  then %><a href="javascript:jsGoPage(<%= ix %>)" onfocus="this.blur();">[next]</a>
							<% else %>[next]<% end if %>
					        </td>        
					    </tr>        
					</table>
				</td>
			</tr>
		<%END IF%>
			</table>
	</td> 
</tr>  
</table> 
</body>
</html> 



	