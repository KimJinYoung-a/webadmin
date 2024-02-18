<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������ ��   ����Ʈ
' History : 2011.06.03 ������ ����
'			2019.04.05 �ѿ�� ����(ǥ���ڵ��� �ƴ϶� ���� ��ã��. ǥ���ڵ�� �ٽ� �ڵ�)
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
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<%
Dim isUseSerp : isUseSerp = True

Dim clsPart,clsOpExp, arrPart, arrList, arrType, intLoop, clsPartMoney
Dim clsAccount, arrAccount ,iarap_cd
Dim  arrUsePart ,sOpExpPartName, sPartTypeName
Dim dYear, dMonth,dSYear, dSMonth, dEYear, 	dEMonth, dDate, iPartTypeIdx, iOpExpPartIdx	,sbizsection_cd,sbizsection_nm
Dim intY, intM
Dim iTotCnt,iPageSize, iTotalPage,iCurrPage
Dim blnAdmin, blnWorker, blnReg ,blnYYYYMM
Dim ipartsn,sadminid, department_id
Dim dedTp,bizNo
''// ===========================================================================
''������ = �����ͱ��� or �濵������
''
''�����, ���μ�, ������ : ��ȸ����
''�����, ������ : �ۼ�����
''// ===========================================================================

	iPageSize = 20
	iCurrPage = requestCheckvar(Request("iCP"),10)
	if iCurrPage="" then iCurrPage=1
	dDate = dateadd("m",-1,date())
  dYear = year(dDate)
  dMonth = month(dDate)

  iPartTypeIdx	= requestCheckvar(Request("selPT"),10)
 	IF iPartTypeIdx = "" THEN iPartTypeIdx = 0
 	iOpExpPartIdx	= requestCheckvar(Request("selP"),10)

 	IF iOpExpPartIdx = "" THEN iOpExpPartIdx = 0
 	dSYear			=  requestCheckvar(Request("selSY"),4)
 	dSMonth			=  requestCheckvar(Request("selSM"),2)
 	IF dSYear = "" THEN dSYear = year(dDate)
 	IF dSMonth = "" THEN dSMonth = month(dDate)
 	dEYear			=  requestCheckvar(Request("selEY"),4)
 	dEMonth			=  requestCheckvar(Request("selEM"),2)
 	IF dEYear = "" THEN dEYear = year(date())
 	IF dEMonth = "" THEN dEMonth = month(date())

 	iarap_cd		= requestCheckvar(Request("selA"),10)
 	sbizsection_nm=requestCheckvar(Request("sBiznm"),100)


 	blnYYYYMM = requestCheckvar(Request("chkD"),1)

 	dedTp  = requestCheckvar(Request("dedTp"),10)
 	bizNo  = requestCheckvar(Request("bizNo"),10)

 	'�����ʱⰪ ����--------------
 	blnWorker = 0 '�����
 	blnReg = 0 	'��ϱ���
  'blnAdmin = fnChkAdminAuth(session("ssAdminLsn"),session("ssAdminPsn"))  '���α���
  ' �繫�� or ������
  blnAdmin = C_MngPart or C_ADMIN_AUTH

  IF blnAdmin THEN blnReg = 1 '���α��� ���� ��� ���ó�� �׻� ����

 '������ �� ���� ����Ʈ
Set clsPart = new COpExpPart
	clsPart.FPartTypeidx 	= 4
	IF Not blnAdmin THEN  '����Ʈ ������ ���� ����� �����ϰ� ����ڿ� ���μ�  view ����
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
	clsOpExp.FSAuthDate 	= dSYear&"-"&Format00(2,dSMonth)
	clsOpExp.FEAuthDate 	= dEYear&"-"&Format00(2,dEMonth)
	clsOpExp.FPartTypeIdx = iPartTypeIdx
	clsOpExp.FOpExpPartIdx = iOpExpPartIdx
	clsOpExp.Farap_cd = iarap_cd
	clsOpExp.FBizsection_nm = sbizsection_nm
	clsOpExp.FisYYYYMM	= blnYYYYMM
	''clsOpExp.FRectPartsn = ipartsn
	clsOpExp.FRectDepartmentID = department_id
	clsOpExp.FRectUserid = sadminid
	clsOpExp.FCurrPage 	= iCurrPage
	clsOpExp.FPageSize 	= iPageSize
	arrList = clsOpExp.fnGetOpExpDailyNoSetList
	iTotCnt = clsOpExp.FTotCnt
	'/����üũ----------------------------
  IF iOpExpPartIdx > 0 THEN
  		clsOpExp.FadminID = session("ssBctId")
  	  blnWorker = clsOpExp.fnGetOpExpPartAuth
  	  IF blnWorker = 1 THEN blnReg = 1
	END IF
Set clsOpExp = nothing
iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��

'####### ERP������ �������� ���� üũ ####### => ���ϸ���
Dim erpdatacountCls, vERPDataCount
'Set erpdatacountCls = new OpExp
'erpdatacountCls.fERPdataCount()
'vERPDataCount = erpdatacountCls.FTotCnt
'SEt erpdatacountCls = Nothing
'####### ERP������ �������� ���� üũ #######
%>
<script type="text/javascript" src="/js/jquery-1.6.2.min.js"> </script>
<script type="text/javascript">
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

//����
function jsModOpExp(sYear, sMonth,iOED,iOpExpPartIdx){
	var winNew = window.open("preRegOpExp.asp?selY="+sYear+"&selM="+sMonth+"&selPT=<%=iPartTypeIdx%>&selP="+iOpExpPartIdx+"&hidOED="+iOED,"popNew","width=1500,height=600,resizable=yes, scrollbars=yes");
	winNew.focus();
}

//����
 	function jsDelOpExp(idx){
 		if(confirm("�����Ͻðڽ��ϱ�?")){
 			document.frmDel.hidOED.value = idx;
 			document.frmDel.hidM.value = 'D';
 			document.frmDel.submit();
 		}
 	}
//����
    function jsLiveOpExp(idx){
 		if(confirm("�����Ͻðڽ��ϱ�?")){
 			document.frmDel.hidOED.value = idx;
 			document.frmDel.hidM.value = 'R';
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
		document.frm.action = "preDailyOpExp.asp";
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

//û���� ���õ��
	function jsSelReg(){
		var ischecked =false;

        for (var i=0;i<frmReg.elements.length;i++){
    		//check optioon
    		var e = frmReg.elements[i];

    		//check itemEA
    		if ((e.type=="checkbox")&&(e.name=="chk")) {
    		    ischecked = e.checked;
    			if (ischecked) break;
    		}
    	}

    	if (!ischecked){
    	    alert('���� ������ �����ϴ�.');
    	    return;
    	}

     	var sYear = document.frmReg.selY.options[document.frmReg.selY.selectedIndex].value;
    	var sMonth = document.frmReg.selM.options[document.frmReg.selM.selectedIndex].value;
     	if (confirm('���� ������ û���� '+sYear+'�� '+sMonth+'���� ����Ͻðڽ��ϱ�?')){
     	    frmReg.action="procOpExp.asp";
     	    frmReg.submit();
     	}

	}

// ERP �����ڷ� ����(ī��)
	function jsSelReg2(){
		var ischecked =false;

        for (var i=0;i<frmReg.elements.length;i++){
    		//check optioon
    		var e = frmReg.elements[i];

    		//check itemEA
    		if ((e.type=="checkbox")&&(e.name=="chk2")) {
    		    ischecked = e.checked;
    			if (ischecked) break;
    		}
    	}

    	if (!ischecked){
    	    alert('���� ������ �����ϴ�.');
    	    return;
    	}

     	if (confirm('���� ������ ERP ī�� �����ڷ�� ���� �Ͻðڽ��ϱ�?')){
     	    frmReg.mode.value="regCardMeaip";
     	    frmReg.action="/admin/tax/eTax_process.asp";
     	    frmReg.submit();
     	}

	}

// sERP �����ڷ� ����(ī��)
    function jsSelReg2_unlock(){
        for (var i=0;i<frmReg.elements.length;i++){
    		//check optioon
    		var e = frmReg.elements[i];

    		//check itemEA
    		if ((e.type=="checkbox")&&(e.name=="chk2")) {
    		    e.disabled=false;
    		}
    	}
    }

    function jsSelReg2_sERP(){
		var ischecked =false;

        for (var i=0;i<frmReg.elements.length;i++){
    		//check optioon
    		var e = frmReg.elements[i];

    		//check itemEA
    		if ((e.type=="checkbox")&&(e.name=="chk2")) {
    		    ischecked = e.checked;
    			if (ischecked) break;
    		}
    	}

    	if (!ischecked){
    	    alert('���� ������ �����ϴ�.');
    	    return;
    	}

     	if (confirm('���� ���� sERP ī�� ���� ������ ��� �Ͻðڽ��ϱ�?')){
     	    frmReg.mode.value="regCardUp";
     	    frmReg.action="/admin/tax/eTax_sERP_process.asp";
     	    frmReg.submit();
     	}

	}

	function jsSelReg2_sERP_OLD(){
		var ischecked =false;

        for (var i=0;i<frmReg.elements.length;i++){
    		//check optioon
    		var e = frmReg.elements[i];

    		//check itemEA
    		if ((e.type=="checkbox")&&(e.name=="chk2")) {
    		    ischecked = e.checked;
    			if (ischecked) break;
    		}
    	}

    	if (!ischecked){
    	    alert('���� ������ �����ϴ�.');
    	    return;
    	}

     	if (confirm('���� ������ ERP ī�� �����ڷ�� ���� �Ͻðڽ��ϱ�?')){
     	    frmReg.mode.value="regCardMeaip";
     	    frmReg.action="/admin/tax/eTax_sERP_process.asp";
     	    frmReg.submit();
     	}

	}
	function CkeckAll(comp){
        var frm = comp.form;
        var bool =comp.checked;
    	for (var i=0;i<frm.elements.length;i++){
    		//check optioon
    		var e = frm.elements[i];

    		//check itemEA
    		if ((e.type=="checkbox")&&(e.name=="chk")) {
    		    if (e.disabled) continue;
    			e.checked=bool;
    			AnCheckClick(e)
    		}
    	}

    }

    function CkeckAll2(comp){
        var frm = comp.form;
        var bool =comp.checked;
    	for (var i=0;i<frm.elements.length;i++){
    		//check optioon
    		var e = frm.elements[i];

    		//check itemEA
    		if ((e.type=="checkbox")&&(e.name=="chk2")) {
    		    if (e.disabled) continue;
    			e.checked=bool;
    			AnCheckClick(e)
    		}
    	}
	}

	function checkThis(comp){
        AnCheckClick(comp)
    }

    //���ϵ��
    function jsNewRegFile(){
			var winF = window.open('/admin/expenses/card/popRegFile.asp?selPT=<%=iPartTypeIdx%>&selP=<%=iOpExpPartIdx%>','popP','width=600, height=500, resizable=yes, scrollbars=yes');
			winF.focus();
	}

    function jsGetERPdata(){
    	document.erpdataproc.location.href = "procOpExp_db_direct.asp";
    }

    function jsGetERPdata_sERP(){
       // alert('�۾���');
       // return;
    	document.erpdataproc.location.href = "procOpExp_db_direct.asp";
    }

function popXL() {
	var popwin = window.open("preDailyOpExp_xl_down.asp?startYYYYMM=<%= dSYear&"-"&Format00(2,dSMonth) %>&endYYYYMM=<%= dEYear&"-"&Format00(2,dEMonth) %>","popXL","width=200,height=100 scrollbars=yes resizable=yes");
	popwin.focus();
}

function chselPval(vselP) {
	$("input[name='selP']").val(vselP)
}

//-->
</script>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="selP" value="<%= iOpExpPartIdx %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="iCP" value="">
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<tr>
	<td>
		<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="#FFFFFF" >
			<td  rowspan="2" width="100" height="50" bgcolor="<%= adminColor("gray") %>">�˻� ����</td>
			<td align="left">
			������ :
			<select name="selSY">
			<%For intY = Year(date()) To 2011 STEP -1%>
			<option value="<%=intY%>" <%IF Cstr(intY) = Cstr(dSYear) THEN%>selected<%END IF%>><%=intY%></option>
			<%Next%>
			</select>��
				<select name="selSM">
			<%For intM = 1 To 12%>
			<option value="<%=intM%>" <%IF Cstr(intM) = Cstr(dSMonth) THEN%>selected<%END IF%>><%=intM%></option>
			<%Next%>
			</select>��
			-
			<select name="selEY">
			<%For intY = Year(date()) To 2011 STEP -1%>
			<option value="<%=intY%>" <%IF Cstr(intY) = Cstr(dEYear) THEN%>selected<%END IF%>><%=intY%></option>
			<%Next%>
			</select>��
				<select name="selEM">
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
				<select name="selPtemp"  id="selP" class="select" onchange="chselPval(this.value);" >
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
				&nbsp;&nbsp;
				��������:
				<select name="dedTp">
				<option value="">--��ü--</option>
				<option value="0" <%=CHKIIF(dedTp="0","selected","")%>>����N</option>
				<option value="1" <%=CHKIIF(dedTp="0","selected","")%>>����Y</option>
				</select>

				&nbsp;&nbsp;
				����ڹ�ȣ:
				<input type="text" name="bizNo" value="<%=bizNo%>" size="12" maxlength="10">

				&nbsp;&nbsp;
				<input type="checkbox" name="chkD" value="Y" <%IF blnYYYYMM ="Y" THEN%>checked<%END IF%>> û���������� ������
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
<form name="frmReg" method="post" action="procOpExp.asp" style="margin:0px;">
<input type="hidden" name="hidM" value="S">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="selPT" value="<%=iPartTypeIdx%>">
<input type="hidden" name="selP" value="<%=iOpExpPartIdx%>">
<input type="hidden" name="hidRU" value="preDailyOpExp.asp">
<input type="hidden" name="iCP" value="">
<input type="hidden" name="mode" value="">
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a">
<%IF  blnAdmin  THEN%>
<Tr>
	<td>
		<table class="a" border="0">
		<tr>
			<td><input type="button" class="button" value="���ϵ��" onClick="jsNewRegFile();"></td>
			<td style="padding-left:10px;">
			    <% if (isUseSerp) then %>
			    <input type="button" class="button" value="sERP������ ��������" onClick="jsGetERPdata_sERP();">
			    <% else %>
				<input type="button" class="button" value="ERP������ ��������" onClick="jsGetERPdata();">
			    <% end if %>
				<span id="erpprocmessage"><!--* <%=vERPDataCount%>���� ������.--></span>
			</td>
			<td><span id="reflashbutton" style="display:none;"><input type="button" class="button" value="������ ���ΰ�ħ" onClick="document.location.reload();"></span></td>
		</tr>
		</table>
	</td>
</tr>
<%END IF%>
<!-- #include virtual="/lib/db/dbclose.asp" -->

<%IF blnAdmin THEN '���α��� �����ڸ� ��ϰ���%>
<tr>
	<td><hr width="100%"><br> û����
		<select name="selY" class="select">
			<%For intY = Year(date()) To 2011 STEP -1%>
			<option value="<%=intY%>" <%IF Cstr(intY) = Cstr(dYear) THEN%>selected<%END IF%>><%=intY%></option>
			<%Next%>
		</select>��
		<select name="selM"  class="select">
			<%For intM = 1 To 12%>
			<option value="<%=intM%>" <%IF Cstr(intM) = Cstr(dMonth) THEN%>selected<%END IF%>><%=intM%></option>
			<%Next%>
		</select>��
		&nbsp;&nbsp;
		<input type="button" name="btnReg" value="���õ��" class="button" onClick="jsSelReg();" >
		(�����׸�,�󼼳���,���μ� �Է� �Ϸ� �� û���� ��ϰ����մϴ�.)

		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;|

		<% if (isUseSerp) then %>
		    <input type="button" value="���� sERP ī�� �������� ���" onClick="jsSelReg2_sERP()" class="button" >
	    <% else %>
		<input type="button" name="btnReg" value="���� ERP ī�� �����ڷ� ���" class="button" onClick="jsSelReg2();" >

		<% if session("ssBctID")="icommang" or session("ssBctID")="ju1209" then %>
		    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;|
		    <font color=red>sERP[</font>
		<input type="button" value="unlock" onClick="jsSelReg2_unlock()" class="button" >
        <input type="button" value="���� sERP ī�� �����ڷ� ���" onClick="jsSelReg2_sERP()" class="button" >
        <font color=red>]</font>
        <% end if %>
        <% end if %>
	</td>
</tr>
<% END IF%>
<tr>
	<td>
		<div align="right">��: <%=formatnumber(iTotCnt,0)%>�� &nbsp; <input type="button" class="button" value="�����ޱ�" onclick="popXL();"></div>
		<!-- ��� �� ���� -->
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a">
		<tr>
			<td>
				<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
					<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
						<td width="20"><input type="checkbox" name="chkAll" onClick="CkeckAll(this)"></td>
						<td width="20"><input type="checkbox" name="chkAll2" onClick="CkeckAll2(this)"></td>
						<td width="50">û����</td>
						<td width="70">�����Ͻ�</td>
						<td>�����ó</td>
						<td>�����׸�</td>
						<td>��ü��</td>
						<td>����ڹ�ȣ</td>
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
					Dim  totOutExp, sumOutExp, iNum, sumSupExp, sumVatExp, sumSevExp, totSupExp, totVatExp, totSevExp
					totOutExp = 0
					sumOutExp=0
					sumSupExp=0
					sumVatExp=0
					sumSevExp=0
					totSupExp=0
					totVatExp=0
					totSevExp=0
					iNum = 1
					IF isArray(arrList) THEN
						For intLoop = 0 To UBound(arrList,2)
					 %>
					<tr height=30 bgcolor="<%= CHKIIF(arrList(22,intLoop)=0,"#CCCCCC","#FFFFFF") %>">
						<td align="center"><input type="checkbox" name="chk" value="<%=arrList(0,intLoop)%>" onClick="checkThis(this)" <%IF arrList(1,intLoop) <> "" or isNull(arrList(5,intLoop)) or  arrList(5,intLoop) ="" or isnull(arrList(12,intLoop)) or arrList(12,intLoop)="" or isNull(arrList(14,intLoop)) or arrList(14,intLoop)="" or arrList(22,intLoop)=0 THEN%>disabled<%END IF%>></td>
						<td align="center"><input type="checkbox" name="chk2" value="<%=arrList(0,intLoop)%>" onClick="checkThis2(this)" <%IF arrList(1,intLoop) = "" or isNull(arrList(5,intLoop)) or  arrList(5,intLoop) ="" or isnull(arrList(12,intLoop)) or arrList(12,intLoop)="" or isNull(arrList(14,intLoop)) or arrList(14,intLoop)="" or arrList(22,intLoop)=0 or arrList(17,intLoop)=0 or (NOT IsNULL(arrList(21,intLoop))) THEN%>disabled<%END IF%>></td>

						<td align="center"><%IF arrList(1,intLoop) <> "" THEN%><%=formatdate(arrList(1,intLoop),"0000-00")%><%END IF%></td>
						<td align="center"><%=formatdate(arrList(2,intLoop),"0000-00-00 00:00:00")%></td>
						<td align="center"><%=arrList(15,intLoop)%></td>
						<td align="center"><%=arrList(5,intLoop)%></td>
						<td><%=arrList(11,intLoop)%></td>
						<td><%=arrList(18,intLoop)%></td>
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

						<% if IsNULL(arrList(21,intLoop)) then %>
						<%IF blnReg = 1 THEN%>
						    <% if (arrList(22,intLoop)<>0) then %>
							<input type="button" class="button" value="����" onClick="jsModOpExp('<%=Year(arrList(2,intLoop))%>','<%=month(arrList(2,intLoop))%>',<%=arrList(0,intLoop)%>,<%=arrList(3,intLoop)%>);">
							<% end if %>
						<%END IF%>
						<%IF blnAdmin THEN%>
						    <% if (arrList(22,intLoop)=0) then %>
						    <input type="button" class="button" value="����" onClick="jsLiveOpExp(<%=arrList(0,intLoop)%>)">
						    <% else %>
							<input type="button" class="button" value="����" onClick="jsDelOpExp(<%=arrList(0,intLoop)%>)">
							<% end if %>
						<%END IF%>
						<% else %>
						    <%= arrList(21,intLoop) %>
						<% end if %>
						</td>
					</tr>
					<%
					  iNum = iNum + 1
					Next
					ELSE%>
					<tr height="30" align="center" bgcolor="#FFFFFF">
						<td colspan="20">��ϵ� ������ �����ϴ�.</td>
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
</form>
<form name="frmDel" method="post" action="procOpExp.asp" style="margin:0px;">
<input type="hidden" name="hidM" value="D">
<input type="hidden" name="hidOED" value="">
<input type="hidden" name="selY" value="<%=dYear%>">
<input type="hidden" name="selM" value="<%=dMonth%>">
<input type="hidden" name="selPT" value="<%=iPartTypeIdx%>">
<input type="hidden" name="selP" value="<%=iOpExpPartIdx%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="hidRU" value="preDailyOpExp.asp">
</form>
<form name="frmDeduct" method="post" action="procOpExp.asp" style="margin:0px;">
<input type="hidden" name="hidM" value="T">
<input type="hidden" name="rdoD" value="">
<input type="hidden" name="hidOED" value="">
<input type="hidden" name="selY" value="<%=dYear%>">
<input type="hidden" name="selM" value="<%=dMonth%>">
<input type="hidden" name="selPT" value="<%=iPartTypeIdx%>">
<input type="hidden" name="selP" value="<%=iOpExpPartIdx%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="hidRU" value="preDailyOpExp.asp">
<input type="hidden" name="selSY" value="<%=dSYear%>">
<input type="hidden" name="selSM" value="<%=dSMonth%>">
<input type="hidden" name="selEY" value="<%=dEYear%>">
<input type="hidden" name="selEM" value="<%=dEMonth%>">
<input type="hidden" name="dedTp" value="<%=dedTp%>">
<input type="hidden" name="bizNo" value="<%=bizNo%>">
</form>
<iframe src="about:blank" name="erpdataproc" width="110" height="110" frameborder=""></iframe>
</body>
</html>
