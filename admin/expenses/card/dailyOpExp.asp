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
<!-- #include virtual="/lib/classes/expenses/OpExpCardCls.asp"-->
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
Dim blnAdmin, blnWorker, blnReg ,ipartsn,sadminid
 
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
 	IF 	iPartTypeIdx = "" THEN iPartTypeIdx = 0
 	IF iOpExpPartIdx ="" THEN iOpExpPartIdx =0
 	'�����ʱⰪ ����-------------- 
 	blnWorker = 0 '�����
 	blnReg = 0 	'��ϱ���
  	blnAdmin = fnChkAdminAuth(session("ssAdminLsn"),session("ssAdminPsn"))  '���α���	
  	
  	IF blnAdmin THEN blnReg = 1 '���α��� ���� ��� ���ó�� �׻� ���� 
 	 
 '������ �� ���� ����Ʈ		
Set clsPart = new COpExpPart 
	IF not blnAdmin THEN  '����Ʈ ������ ���� ����� �����ϰ� ����ڿ� ���μ�  view ����
		ipartsn  =  session("ssAdminPsn")
 		sadminid = 	session("ssBctId")
 	END IF	
	clsPart.FRectPartsn = ipartsn
	clsPart.FRectUserid = sadminid  
	arrType = clsPart.fnGetOpExpPartTypeCardList 
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
Set clsOpExp = nothing	 

iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
 
%> 
<script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script> 
<script language="javascript">
<!--
 //�� ���� 
// =========================================================================================================
$(document).ready(function(){
	$("#selPT").change(function(){
		var iValue = $("#selPT").val();
		var url="/admin/expenses/part/ajaxPart.asp";
		 var params = "iPTIdx="+iValue;  
		  	 
		 $.ajax({
		 	type:"POST",
		 	url:url,
		 	data:params,
		 	success:function(args){   
		 		$("#divP").html(args);   
		 	},
		 	 
		 	error:function(e){ 
		 		alert("�����ͷε��� ������ ������ϴ�. �ý������� �������ּ���");
		 		//alert(e.responseText);
		 	}
		 }); 
	}); 
});
     
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
	var winNew = window.open("regOpExp.asp?selY=<%=dyear%>&selM=<%=dmonth%>&selPT=<%=iPartTypeIdx%>&selP=<%=iOpExpPartIdx%>&hidOED="+iOED,"popNew","resizable=yes, scrollbars=yes");
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
	
	//����Ÿ�Ժ���
 function jsSetDeduct(idx,iType){
 		document.frmDeduct.hidOED.value = idx;
 		document.frmDeduct.rdoD.value = iType;
 		document.frmDeduct.submit();
}
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">  
<form name="frmDel" method="post" action="procOpExp.asp">
<input type="hidden" name="hidM" value="D">  
<input type="hidden" name="hidOED" value="">
<input type="hidden" name="selY" value="<%=dYear%>">
<input type="hidden" name="selM" value="<%=dMonth%>"> 
<input type="hidden" name="selPT" value="<%=iPartTypeIdx%>">
<input type="hidden" name="selP" value="<%=iOpExpPartIdx%>">
<input type="hidden" name="menupos" value="<%=menupos%>"> 
<input type="hidden" name="hidRU" value="dailyOpExp.asp">
</form>
<form name="frmDeduct" method="post" action="procOpExp.asp">
<input type="hidden" name="hidM" value="T">  
<input type="hidden" name="rdoD" value="">
<input type="hidden" name="hidOED" value="">
<input type="hidden" name="selY" value="<%=dYear%>">
<input type="hidden" name="selM" value="<%=dMonth%>"> 
<input type="hidden" name="selPT" value="<%=iPartTypeIdx%>">
<input type="hidden" name="selP" value="<%=iOpExpPartIdx%>">
<input type="hidden" name="menupos" value="<%=menupos%>"> 
<input type="hidden" name="hidRU" value="dailyOpExp.asp">
</form>
<tr>
	<td>+ <a href="javascript:jsGoList('index');">����ī����� ���� ����Ʈ</a> > <a href="javascript:jsGoList('dailySumOpExp');">���� �� ����Ʈ</a> > �Ϻ� �� ����Ʈ</td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="">
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="iCP" value=""> 
			<tr align="center" bgcolor="#FFFFFF" >
				<td  rowspan="2" width="100" height="50" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
				<td align="left">
				��¥(û����):
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
				<select name="selPT"  id="selPT"   class="select">
					<option value="0">--����--</option>
					<% sbOptPartType arrType,ipartTypeIdx%>
					</select>
					<span id="divP">
					<select name="selP"  id="selP" class="select">
					<option value="0">--����--</option>	
					<% sbOptPart arrPart,iOpExpPartIdx%>
					</select> 
					</span>	 
					&nbsp;&nbsp;
				</td>
					<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="jsSearch();">
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFFF"> 
				�����׸�:
				<select name="selA">
				<option value="0">--��ü--</option>
				<% sbOptAccount arrAccount, iarap_cd%>
				</select>
				&nbsp;&nbsp;
				���μ�:
				<input type="text" name="sBiznm" value="<%=sbizsection_nm%>" size="20"> 
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
		<%IF iOpExpPartIdx > 0 THEN%>
			<!--td align="right"><input type="button" class="button" value="����Ʈ" onClick="jsPrint();"></td-->
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
				<input type="hidden" name="selP" value="<%=iOpExpPartIdx%>"> 
				<input type="hidden" name="selPT" value="<%=iPartTypeIdx%>"> 
				<input type="hidden" name="hidRU" value="dailyOpExp.asp">
				<input type="hidden" name="iCP" value="">
					<tr align="center" bgcolor="<%= adminColor("tabletop") %>">  
						<td width="50">����</td>
						<td width="50">������</td>  
						<td>�����ó</td>   
						<td>�����׸�</td>
						<td>��ü��</td>  
						<td>����(�󼼳���)</td>    
						<td>����</td>   
						<td>���ް���</td> 
						<td>�ΰ���</td> 
						<td>�����</td> 
						<td>���ι�ȣ</td> 
						<td>��������</td>  
						<td>����/��</td> 
						<td>���μ�</td> 
						<td>��������</td>  
						<td width="100">ó��</td>     
					</tr> 
					<% 
					Dim   iNum  
					iNum = 1
					IF isArray(arrList) THEN
						For intLoop = 0 To UBound(arrList,2)  
					 %>  
					<tr height=30 bgcolor="#FFFFFF">	 
						<td align="center"><%=iNum%></td>
						<td align="center"><%=formatdate(arrList(2,intLoop),"0000-00-00")%></td>
						<td align="center"><%=arrList(15,intLoop)%></td>
						<td align="center"><%=arrList(5,intLoop)%></td>
						<td><%=arrList(11,intLoop)%></td>
						<td><%=arrList(12,intLoop)%></td> 
						<td align="right"><%=formatnumber(arrList(6,intLoop),0)%></td>
						<td align="right"><%=formatnumber(arrList(7,intLoop),0)%></td> 
						<td align="right"><%=formatnumber(arrList(8,intLoop),0)%></td> 
						<td align="right"><%=formatnumber(arrList(9,intLoop),0)%></td> 
						<td align="center"><%=arrList(10,intLoop)%></td>  
						<td align="center"><%=arrList(16,intLoop)%></td>  
						<td align="center"><%IF arrList(19,intLoop)=1 THEN%>����<%ELSE%>����<%END IF%></td> 
						<td align="center"><%=arrList(14,intLoop)%></td> 
						<td align="center"><%IF blnReg = 1 THEN%><a href="javascript:jsSetDeduct(<%=arrList(0,intLoop)%>,'<%IF arrList(17,intLoop) THEN%>0<%ELSE%>1<%END IF%>');"><img src="/images/icon_arrow_link.gif" align="absmiddle" border="0"> <%END if%><%IF arrList(17,intLoop) THEN%><font color="red">Y</font><%ELSE%><font color="blue">N</font><%END IF%></a></td> 
						<td align="center">
						<% if IsNULL(arrList(23,intLoop)) then %>
						<%IF blnReg = 1 THEN%>
							<input type="button" class="button" value="����" onClick="jsModOpExp(<%=arrList(0,intLoop)%>);">
							<input type="button" class="button" value="����" onClick="jsDelOpExp(<%=arrList(0,intLoop)%>)">
						<%END IF%>
						<% else %>
						    <%= arrList(23,intLoop) %>
						<% end if %>
						</td>
					</tr>	
					<%  
					  iNum = iNum + 1 
					Next  %> 
					<%
					ELSE%>
					<tr height="30" align="center" bgcolor="#FFFFFF">				
						<td colspan="16">��ϵ� ������ �����ϴ�.</td>	
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



	