<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������    ����Ʈ
' History : 2011.06.03 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpAccountCls.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpCardCls.asp"-->
<!-- #include virtual="/lib/classes/expenses/OpExpPartCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
Dim isUseSerp : isUseSerp = true

Dim clsPart, arrType,arrPart,clsOpExp
Dim arrList, intLoop
Dim dSYear, dSMonth,dEYear, dEMonth, iPartTypeIdx ,iOpExpPartIdx
Dim intY, intM
Dim blnAdmin, blnWorker ,blnReg, blnSearch, sadminid, ipartsn, department_id
Dim iState
dim sBankName,sBankAccNo,arrIK,intK

''// ===========================================================================
''������ = �����ͱ��� or �濵������
''
''�����, ���μ�, ������ : ��ȸ����
''�����, ������ : �ۼ�����
''// ===========================================================================

 	dSYear			= requestCheckvar(Request("selSY"),4)
 	dSMonth			= requestCheckvar(Request("selSM"),2)
 	dEYear			= requestCheckvar(Request("selEY"),4)
 	dEMonth			= requestCheckvar(Request("selEM"),2)
 	iPartTypeIdx	= requestCheckvar(Request("selPT"),10)
 	iOpExpPartIdx	= requestCheckvar(Request("iPS"),10)
 	iState			= requestCheckvar(Request("selSt"),1)

 	IF dSYear = "" THEN dSYear = year(dateadd("m",-1,date()))
 	IF dSMonth = "" THEN dSMonth = month(dateadd("m",-1,date()))
 	IF dEYear = "" THEN dEYear = year(date())
 	IF dEMonth = "" THEN dEMonth = month(date())
 	IF iPartTypeIdx = "" THEN iPartTypeIdx = 0
 	IF iOpExpPartIdx ="" THEN iOpExpPartIdx = 0

 	'�����ʱⰪ ����--------------
 	blnWorker = 0 '�����
 	blnReg = 0 	'��ϱ���
  	blnAdmin = fnChkAdminAuth(session("ssAdminLsn"),session("ssAdminPsn"))  '���α���

  	IF blnAdmin THEN blnReg = 1 '���α��� ���� ��� ���ó�� �׻� ����

 '������ �� ���� ����Ʈ
Set clsPart = new COpExpPart
	IF not blnAdmin THEN  '����Ʈ ������ ���� ����� �����ϰ� ����ڿ� ���μ�  view ����
		ipartsn  =  session("ssAdminPsn")
		department_id = GetUserDepartmentID("",session("ssBctID"))
 		sadminid = 	session("ssBctId")
 	END IF
	''clsPart.FRectPartsn = ipartsn
	clsPart.FRectDepartmentID = department_id
	clsPart.FRectUserid = sadminid
	arrType = clsPart.fnGetOpExpPartTypeCardListNew
	IF iPartTypeIdx > 0 THEN
	clsPart.FPartTypeidx 	= iPartTypeIdx
	arrPart = clsPart.fnGetOpExppartAllListNew
	END IF
Set clsPart = nothing


Set clsOpExp = new OpExp
	'��� ����Ʈ
	''clsOpExp.FRectPartsn = ipartsn
	clsOpExp.FRectDepartmentID = department_id
	clsOpExp.FRectUserid = sadminid
	clsOpExp.FSYYYYMM	= dSYear&"-"&Format00(2,dSMonth)
	clsOpExp.FEYYYYMM	= dEYear&"-"&Format00(2,dEMonth)
	clsOpExp.FPartTypeIdx	=iPartTypeIdx
	clsOpExp.FOpExpPartIdx	=iOpExpPartIdx
	clsOpExp.FState = iState
	arrList = clsOpExp.fnGetOpExpMonthlyList

    IF isArray(arrList) THEN
        sBankName = arrList(14,0)
        sBankAccNo = replace(arrList(15,0),"-","")
    END IF

    IF  sBankName <> "" and    sBankAccNo <>"" THEN
    clsOpExp.FRectBankNM    = sBankName
    clsOpExp.FRectBankAccNo = sBankAccNo
    arrIK = clsOpExp.fnGetIpkumList
    END IF
    
	'����üũ------------------------
	IF iOpExpPartIdx > 0  THEN	'��� ���ó ���а� ���� ��쿡�� üũ
	clsOpExp.FOpExpPartIdx	= iOpExpPartIdx
	clsOpExp.FadminID 		= session("ssBctId")
	blnWorker = clsOpExp.fnGetOpExpPartAuth '����� ���� Ȯ��
 
	IF  blnWorker =1  THEN	blnReg =1 '������̰ų� ���α����� ���� ��� ���ó�� ����
	END IF
	'/����üũ------------------------
	
Set clsOpExp = nothing

%>
<script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script>
<script language="javascript">
<!--
//�� ����
// =========================================================================================================
$(document).ready(function(){
	$("#selPT").change(function(){
		var iValue = $("#selPT").val();
		var url="/admin/expenses/part/ajaxDepartment.asp";
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

//���ε��
function jsNewReg(){
	var winNew = window.open("about:blank","popNew","width=1500,height=600,resizable=yes, scrollbars=yes");
	document.frm.target = "popNew";
	document.frm.action = "regOpExp.asp";
	document.frm.submit();
	winNew.focus();
}

//���ϵ��
function jsNewRegFile(){
			var winF = window.open('/admin/expenses/card/popRegFile.asp?selP=<%=iOpExpPartIdx%>','popP','width=600, height=500, resizable=yes, scrollbars=yes');
			winF.focus();
	}

 //�󼼺���
 function jsDetail(sPage, dyear, dmonth, ipartypeidx, iopexppartidx){
 	location.href = sPage +".asp?selY="+dyear+"&selM="+dmonth+"&selPT="+ipartypeidx+"&selP="+iopexppartidx+"&menupos=<%=menupos%>";
 }

 	//���ڰ��� ǰ�Ǽ� ���
	function jsRegEapp(dyyyymm, iOpexpPartidx, iPartTypeIdx){
		var winEapp = window.open("eappOpExp.asp?dyyyymm="+dyyyymm+"&hidP="+iOpexpPartidx+"&hidPT="+iPartTypeIdx,"popE","width=1200,height=600,scrollbars=yes,resizable=yes");
		winEapp.focus();
	}

	//���ڰ��� ǰ�Ǽ� ���뺸��
	function jsViewEapp(reportidx,reportstate){
		var winEapp = window.open("/admin/approval/eapp/modeapp.asp?blnP=1&iRS="+reportstate+"&iridx="+reportidx,"popE","");
		winEapp.focus();
	}

	//���º���ó��
	function jsOpExpConfirm(strMsg,sY,sM,iOpExp,istate){
		if(confirm(strMsg)){
		document.frmC.hidOE.value = iOpExp;
		document.frmC.hidS.value = istate;
		document.frmC.selY.value = sY;
		document.frmC.selM.value = sM;
		document.frmC.submit();
		}
		}

	//�˻�
	function jsSearch(){
		document.frm.target = "_self";
		document.frm.action = "index.asp";
		document.frm.iPS.value = $("#selP").val();
		document.frm.submit();
	}

function jsLinkERP(frm){
    var ischecked =false;

    for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    ischecked = e.checked;
			if (ischecked) break;
		}
	}

	if (!ischecked){
	    alert('���� ������ �����ϴ�.');
	    return;
	}
	//alert('�۾���.. 10/17�� ���� �۾��ϰ���.');
	//return;
	if (confirm('���� ������ ERP�� �����Ͻðڽ��ϱ�?')){
	    frm.LTp.value="D";
	    frm.action="/admin/approval/payreqList/erpLink_Process.asp";
	    frm.submit();
	}
}

function jsLink_SERP_unlock(frm){
    for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    e.disabled=false;
		}
	}
}

function jsLinkERP_sERP(frm){
    alert('��������޴�');
    return;
    var ischecked =false;

    for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    ischecked = e.checked;
			if (ischecked) break;
		}
	}

	if (!ischecked){
	    alert('���� ������ �����ϴ�.');
	    return;
	}

	if (confirm('���� ������ ERP�� �����Ͻðڽ��ϱ�?')){
	    frm.LTp.value="D";
	    frm.action="/admin/approval/payreqList/S_erpLink_Process.asp";
	    frm.submit();
	}
}


function jsMakeMonth(frm){
    var cstr = frm.selY.value+'-'+frm.selM.value+' ���� �����͸� �����Ͻðڽ��ϱ�?'
    if (confirm(cstr)){
        frm.submit();
    }
}

function CkeckAll(comp){
    var frm = comp.form;
    var bool =comp.checked;
	for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    if (e.disabled) continue;
			e.checked=bool;
			AnCheckClick(e)
		}
	}
}

function checkThis(comp){
    AnCheckClick(comp)
}
//-->
</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td>+ ����ī����� ���� ����Ʈ </td>
</tr>
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<form name="frm" method="get" action="index.asp" >
			<input type="hidden" name="menupos" value="<%= menupos %>">
			<input type="hidden" name="iCP" value="">
			<input type="hidden" name="iPS" value="">
			<tr align="center" bgcolor="#FFFFFF" >
				<td width="100" height="50" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
				<td align="left">
					û����:
					<select name="selSY" class="select">
					<%For intY = Year(date()) To 2011 STEP -1%>
					<option value="<%=intY%>" <%IF Cstr(intY) = Cstr(dSYear) THEN%>selected<%END IF%>><%=intY%></option>
					<%Next%>
					</select>��
					 <select name="selSM"  class="select">
					<%For intM = 1 To 12%>
					<option value="<%=intM%>" <%IF Cstr(intM) = Cstr(dSMonth) THEN%>selected<%END IF%>><%=intM%></option>
					<%Next%>
					</select>��
					-
					<select name="selEY" class="select">
					<%For intY = Year(date()) To 2011 STEP -1%>
					<option value="<%=intY%>" <%IF Cstr(intY) = Cstr(dEYear) THEN%>selected<%END IF%>><%=intY%></option>
					<%Next%>
					</select>��
					 <select name="selEM" class="select">
					<%For intM = 1 To 12%>
					<option value="<%=intM%>" <%IF Cstr(intM) = Cstr(dEMonth) THEN%>selected<%END IF%>><%=intM%></option>
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
					����:
					<select name="selSt" id="selSt" class="select">
					<% SbOptState iState%>
					</select>
				</td>
				<td width="50" bgcolor="<%= adminColor("gray") %>">
					<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSearch();">
				</td>
			</tr>
			</form>
		</table>
	</td>
</tr>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%IF    blnReg =1    THEN %>
<tr>
    <td> + ���� ������ ����ݳ���
        <table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
            <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
                <td>IDX</td>
            	<td>�����</td>
            	<td>���¹�ȣ</td>
            	<td>�������������</td>
            	<td>���ó</td>
            	<td>�ŷ�����</td>
              	<td>�Աݱݾ�</td>
            	<td>��ݱݾ�</td>
            	<td>�ܾ�</td>
            	<td>������Ʈ�ð�</td> 
            </tr>
            
            <%IF isArray(arrIK) THEN
                For intK = 0 To UBound(arrIK,2)
                %>
             <tr height="30" align="center" bgcolor="#FFFFFF">
                <td><%=arrIK(0,intK)%></td>
            	<td><%=arrIK(1,intK)%></td>
            	<td><%=arrIK(2,intK)%></td>
            	<td><%=arrIK(3,intK)%></td>
            	<td><%=arrIK(4,intK)%></td>
            	<td><%=arrIK(5,intK)%></td>
              	<td><%IF arrIK(6,intK) = 2 THEN%><%=formatnumber(arrIK(7,intK),0)%><%ELSE%>0<%END IF%></td>
            	<td><%IF arrIK(6,intK) = 1 THEN%><%=formatnumber(arrIK(7,intK),0)%><%ELSE%>0<%END IF%></td>
            	<td><%=formatnumber(arrIK(8,intK),0)%></td>
            	<td><%=arrIK(9,intK)%></td> 
            </tr>    
            <%  Next 
            ELSE%>
            <tr  height="30" align="center" bgcolor="#FFFFFF">
                <td colspan="10">��ϵ� ������ �������� �ʽ��ϴ�.</td>
            </tr>
            <%END IF%>
        </table>
    </td>
</tr>
<%END IF%>
<tr>
	<td>
	    <table width="100%" cellspacing="0" cellpadding="0">
	    <tr>
	    	<%IF  FALSE and blnReg =1    THEN%>
	    	<td>
					<input type="button" class="button" value="���󼼳��� �űԵ��" onClick="jsNewReg();">
					<input type="button" class="button" value="���ϵ��" onClick="jsNewRegFile();">
			</td>
	    	<%END IF%>
	    	<% IF (blnAdmin) THEN %>
	    	<td align="left" ><input type="button" class="button" value="������������(<%=dSYear%>-<%=dSMonth%>)" onClick="jsMakeMonth(frmMnAct);"></td>
	        <td align="right" >
	            <% if (isUseSerp) then %>
	            <!-- ������.
	                <input type="button" value="sERP ����" onClick="jsLinkERP_sERP(frmAct)"> 
	             --> 
	            <% else %>
	            <input type="button" class="button" value="ERP ����" onClick="jsLinkERP(frmAct);">
	            
    	        <% if session("ssBctID")="icommang" or session("ssBctID")="ju1209XXX" then %>
    	            <font color=red>sERP[</font> 
    	            <input type="button" value="unlock" onClick="jsLink_SERP_unlock(frmAct)">
                    <input type="button" value="sERP ����" onClick="jsLinkERP_sERP(frmAct)"> 
                    <font color=red>]</font>
                <% end if %>
                <% end if %>
	        </td>
	        
	        
            
	      <% END IF %>
	    </tr>
	    </table>
	</td>
</tr>
<tr>
	<td>
		<!-- ��� �� ���� -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		    <Form name="frmAct" method="post" action="/admin/approval/payreqList/erpLink_Process.asp">
		    <input type="hidden" name="LTp" value="C">
			<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
			    <% IF (blnAdmin) THEN %>
			    <td width="20"><input type="checkbox" name="chkAll" onClick="CkeckAll(this)"></td>
			    <% END IF %>
				<td>��¥</td>
				<td>����</td>
				<td>�����ó</td>
				<td>�������</td>
				<td>����</td>
				<%IF blnReg=1  THEN%>
				<td>ó��</td>
				<%END IF%>
				<td>�濵������<br>����Ȯ��</td>
				<% IF (blnAdmin) THEN %><td>ERP<br>��������</td>  <% END IF %>
				<td>�󼼳�������</td>
			</tr>
			<%   dim dRectY, dRectM
			IF isArray(arrList) THEN
				For intLoop = 0 To UBound(arrList,2)
				dRectY = year(arrList(1,intLoop))
				dRectM = month(arrList(1,intLoop))
			 %>
			<tr height=30 align="center" bgcolor="#FFFFFF">
			    <% IF (blnAdmin) THEN %>
			    <td><input type="checkbox" name="chk" value="<%=arrList(0,intLoop)%>" onClick="checkThis(this)" <%= CHKIIF((arrList(10,intLoop)="9") and (arrList(13,intLoop)=10) and (arrList(1,intLoop)>="2012-09") or (TRUE),"","disabled") %> ></td>
			    <% END IF %>
				<td><%=arrList(1,intLoop)%></td>
				<td><%=fnGetPartTypeDesc(arrList(13,intLoop))%></td>
				<td><%=arrList(7,intLoop)%> </td>
				<td><%=formatnumber(arrList(3,intLoop),0)%></td>
				<td><%=fnGetStateDesc(arrList(10,intLoop))%></td>
				<%IF ( blnReg=1 ) THEN%>
				<td>
				    <%=arrList(10,intLoop)%>
				    <% IF (arrList(10,intLoop)>9) THEN %>
				    �����Ұ� <input type="button" class="button" value="ǰ�Ǽ����� >" onClick="jsViewEapp('<%=arrList(8,intLoop)%>','<%=arrList(9,intLoop)%>')">
				    <% ELSE %>
    					<%IF (arrList(10,intLoop) = 1 and blnWorker = 1) OR (arrList(10,intLoop) >0 and arrList(10,intLoop) < 9 and blnAdmin ) THEN %>
    					<input type="button" class="button" style="color:gray;" value="< �ۼ���" onClick="jsOpExpConfirm('�ۼ��� ���·� �����Ͻðڽ��ϱ�?',<%=year(arrList(1,intLoop))%>,<%=month(arrList(1,intLoop))%>,'<%=arrList(0,intLoop)%>',0)">
    					<%END IF%>
    					<%IF isNull(arrList(8,intLoop)) and  (arrList(10,intLoop) = 1 or arrList(10,intLoop) = 5) THEN %>
    						<input type="button" class="button"   value="ǰ�Ǽ��ۼ� >" onClick="jsRegEapp('<%=arrList(1,intLoop)%>','<%=arrList(2,intLoop)%>','<%=arrList(13,intLoop)%>')">
    					<%ELSEIF not isNull(arrList(8,intLoop))  THEN%>
    						<input type="button" class="button" value="ǰ�Ǽ����� >" onClick="jsViewEapp('<%=arrList(8,intLoop)%>','<%=arrList(9,intLoop)%>')">
    					<%ELSE%>
    						<%IF blnAdmin THEN%>
    						<input type="button" class="button" value="�ۼ��Ϸ� >" onClick="jsOpExpConfirm('�ۼ��Ϸ��Ͻðڽ��ϱ�?',<%=year(arrList(1,intLoop))%>,<%=month(arrList(1,intLoop))%>,'<%=arrList(0,intLoop)%>',1)">
    						<%eND IF%>
    					<%END IF%>
					<% END IF %>
				</td>
				<%END IF%>
				<td>
					<%if  blnAdmin  and  (arrList(10,intLoop) >=7 ) and  (arrList(10,intLoop) <10 ) then%>
					<input type="radio" name="rdoC<%=arrList(0,intLoop)%>" value="1" <%IF arrList(10,intLoop) = 9 THEN%>checked<%END IF%> onClick="jsOpExpConfirm('����Ȯ�λ��·� �����Ͻðڽ��ϱ�?',<%=year(arrList(1,intLoop))%>,<%=month(arrList(1,intLoop))%>,<%=arrList(0,intLoop)%>,9)"><font color="blue">Y</font>
					<input type="radio" name="rdoC<%=arrList(0,intLoop)%>" value="0" <%IF arrList(10,intLoop) <> 9 THEN%>checked<%END IF%>  onClick="jsOpExpConfirm('����Ȯ���� ����Ͻðڽ��ϱ�?',<%=year(arrList(1,intLoop))%>,<%=month(arrList(1,intLoop))%>,<%=arrList(0,intLoop)%>,7)"><font color="red">N</font>
					<%else%>
						<%IF arrList(10,intLoop) >= 9 THEN %>
							<font color="blue">Y</font></a>
						<%ELSE%>
								<font color="red">N</font></a>
						<%END IF%>
					<%end if%>
				</td>
				<% IF (blnAdmin) THEN %>
				<td>
				    <% if Not IsNULL(arrList(12,intLoop)) then %>
				    [<%= arrList(11,intLoop) %>]<%= arrList(12,intLoop) %>
	                <% end if %>
  				</td>
				<% END IF %>
				<td>
					<a href="javascript:jsDetail('dailySumOpExp','<%=dRectY%>','<%=dRectM%>','<%=arrList(13,intLoop)%>','<%=arrList(2,intLoop)%>')">[������]</a>
					<a href="javascript:jsDetail('dailyOpExp','<%=dRectY%>','<%=dRectM%>','<%=arrList(13,intLoop)%>','<%=arrList(2,intLoop)%>')">[�Ϻ���]</a>
				</td>
			</tr>
		<%
			Next
			ELSE%>
			<tr height="30" align="center" bgcolor="#FFFFFF">
				<td colspan="13">��ϵ� ������ �����ϴ�.</td>
			</tr>
			<%END IF%>
			</form>
		</table>
	</td>
</tr>
</table>
<form name="frmC" method="post" action="procOpExp.asp">
<input type="hidden" name="hidM" value="C">
<input type="hidden" name="hidOE" value="">
<input type="hidden" name="hidS" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="selSY" value="<%= dSYear %>">
<input type="hidden" name="selSM" value="<%= dSMonth %>">
<input type="hidden" name="selY" value="" >
<input type="hidden" name="selM" value="">
<input type="hidden" name="selP" value="<%= iOpExpPartIdx %>">
<input type="hidden" name="selPT" value="<%= iPartTypeIdx %>">
</form>
<form name="frmMnAct" method="post" action="procOpExp.asp">
<input type="hidden" name="hidM" value="M">
<input type="hidden" name="hidOE" value="">
<input type="hidden" name="hidS" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="selY" value="<%= dSYear %>">
<input type="hidden" name="selM" value="<%= dSMonth %>">
<input type="hidden" name="selP" value="<%= iOpExpPartIdx %>">
<input type="hidden" name="selPT" value="<%= iPartTypeIdx %>">
</form>
</body>
</html>
