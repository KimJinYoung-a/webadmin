<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���ڰ��� �������� ���� ��� 
' History : 2011.03.10 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/approval/commCls.asp" -->
<!-- #include virtual="/lib/classes/approval/accountCls.asp" -->
<!-- #include virtual="/lib/classes/approval/edmsCls.asp" -->
<%
Dim clsAccount,clsComm, clsedms
Dim iaccountidx , iaccountkind, iedmsidx, saccountname, blnusing
Dim icateidx1, icateidx2 
Dim sMode,menupos
  
iaccountidx= requestCheckvar(Request("iaidx"),10)  
icateidx1= requestCheckvar(Request("selC1"),10)  
icateidx2= requestCheckvar(Request("selC2"),10)  
iedmsidx= requestCheckvar(Request("selC3"),10)  
sMode = "I"
if icateidx1 ="" THEN icateidx1 =0
if icateidx2 ="" THEN icateidx2 =0
if iedmsidx = "" THEN iedmsidx = 0	
Set clsAccount= new CAccount
IF iaccountidx <> "" THEN
	sMode ="U"
	clsAccount.Faccountidx = iaccountidx
	clsAccount.fnGetAccountData	
	  
	iaccountIdx 	 	= clsAccount.FaccountIdx 	  
	iaccountKind     	= clsAccount.FaccountKind     
	iedmsIdx         	= clsAccount.FedmsIdx         
	saccountName    	= clsAccount.FaccountName   
	icateidx1			= clsAccount.Fcateidx1 
	icateidx2			= clsAccount.Fcateidx2  
END IF
 Set clsAccount= nothing
 
 
%>  
 <script type="text/javascript" src="/js/ajax.js"></script>
<script language="javascript">
<!-- 
    // ī�װ� ajax =========================================================================================================
    initializeReturnFunction("processAjax()");
    initializeErrorFunction("onErrorAjax()");
    
    var _divName = "CM";
    
    function processAjax(){
        var reTxt = xmlHttp.responseText;  
        eval("document.all.div"+_divName).innerHTML = reTxt;
    }
    
    function onErrorAjax() {
            alert("ERROR : " + xmlHttp.status);
    }
    
    //������ ī�װ��� ���� ���� ī�װ� ����Ʈ �������� Ajax
    function jsSetCategory(sMode){	 
        var ipcidx  = document.frmReg.selC1.value;
        var icidx   = document.frmReg.selC2.value;  
        var ieidx	= document.frmReg.selC3.value;     
        _divName = sMode;   
        initializeURL('/admin/approval/edms/ajaxCategory.asp?sMode='+sMode+'&ipcidx='+ipcidx+'&icidx='+icidx+'&ieidx='+ieidx);
    	startRequest();
    	
    }
 
	//����� �ʵ� üũ
	function jsSubmit(){
	 if(document.frmReg.sAN.value==""){
	 alert("���������� �Է����ּ���");
	 document.frmReg.sAN.focus();
	 return false;
	 } 
	  
	 return true;
	}
//-->
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><strong>�������� ������</strong><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" align="center" class="a" bgcolor=#BABABA> 
		<form name="frmReg" method="post" action="procAccount.asp" OnSubmit="return jsSubmit();">
		<input type="hidden" name="hidM" value="<%=sMode%>">
		<input type="hidden" name="iaidx" value="<%=iaccountIdx%>">  
		<%IF sMode="U" THEN%>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">IDX</td>
			<td bgcolor="#FFFFFF"><%=iaccountIdx%> </td>
		</tr>	
		<%END IF%>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">��������</td>
			<td bgcolor="#FFFFFF"><input type="text" name="sAN" size="30" maxlength="32" value="<%=sAccountName%>"></td>
		</tr>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">��������</td>
			<td bgcolor="#FFFFFF"> 
			<select name="selAK">
			<option value="0">--����--</option>
			<% 	set clsComm = new CcommCode
				clsComm.Fparentkey = 1
				clsComm.Fcomm_cd = iaccountkind
				clsComm.sbOptCommCD
				Set clsComm = nothing 
			%>
			</select> 
			</td>
		</tr>
		
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center" rowspan="3">������</td>
			<td bgcolor="#FFFFFF">
			<%set clsedms = new Cedms%>
			<div id="divCL"> 
					��ī�װ�:
					<select name="selC1" onChange="jsSetCategory('CM');">
					<option value="0">--����--</option> 
					<%clsedms.sbGetOptedmsCategory 1,0,icateidx1 %>
					</select> 
			</div>	
			</td>
		</tr>
		<tr>
			<td bgcolor="#FFFFFF">	
			<div id="divCM"> 
					��ī�װ�:
					<select name="selC2" onChange="jsSetCategory('CD');">
					<option value="0">-- ���� --</option>
				<% 	IF icateidx1 > 0 THEN	'��ī�װ� ���� �� ��ī�װ� ���ð����ϰ�
						clsedms.sbGetOptedmsCategory 2,icateidx1,icateidx2 
					END IF
				%>
					</select> 
			</div>
			</td>
		</tr>
		<tr>
			<td bgcolor="#FFFFFF">
			<div id="divCD"> 
					������:
					<select name="selC3" >
					<option value="0">-- ���� --</option>
				 	<% 	IF icateidx1 > 0 and icateidx2>0 THEN	
				 		clsedms.FCateIdx1 = icateidx1
				 		clsedms.FCateIdx2 = icateidx2 
				 		clsedms.Fedmsidx = iedmsidx
						clsedms.sbOptPayEdmsList 
						END IF
					%>
					</select> 
			</div>
				<%set clsedms = nothing %>
			</td>
		</tr> 
		<%IF sMode="U" THEN%>
		<tr>
			<td  bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">�������</td>
			<td bgcolor="#FFFFFF" width="180"><input type="radio" name="blnU" value="1" checked>��� <input type="radio" name="blnU" value="0">������</td>
		</tr>	
		<%END IF%>
		</table>
</td>
</tr>
<tr>
	<td align="center"><input type="submit" value="���" class="button"></td>
</tr>
</form>
</table>
</body>
</html> 

<!-- #include virtual="/lib/db/dbclose.asp" --> 	
	