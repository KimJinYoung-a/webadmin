<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ü ��� ����
' Hieditor : 2016.02.15 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/classes/partners/contractclsNew.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%
dim makerid, regUserid,dispCate,arect,contractNo,uniqGroupID, reqCtrSearch,grpType,ctrType,crect
dim ContractState, reqCtr,notboru,jMonth
dim iCurrpage, iPageSize, iTotCnt, iTotalPage,iPerCnt
dim cCtrList, intLoop, arrList
dim contracttype
dim oldgroupid
dim nregUserid, ncontractNo ,nreqCtrSearch, nctrType,nContractState,nreqCtr,nnotboru  
dim strParm
dim selSP, nselSP 
dim arrgroupid, intG,iTotGCnt
dim regDefUserid

    iCurrpage = requestCheckVar(Request("iC"),10)	'���� ������ ��ȣ
	IF iCurrpage = "" THEN
		iCurrpage = 1
	END IF
	iPageSize = 100		'�� �������� �������� ���� ��
	iPerCnt = 10		'�������� ������ ����
	 
	makerid = requestCheckVar(request("makerid"),32) 
	dispCate = requestCheckvar(request("dispCate"),10)
	arect   = requestCheckVar(request("arect"),32)
 	crect   = requestCheckVar(request("crect"),32) 
 	
 	regDefUserid    = requestCheckVar(request("rDU"),32)
 	reguserid       = requestCheckVar(request("rU"),32)
 	contractNo      = requestCheckVar(request("contractNo"),20)
 	ContractState   = requestCheckVar(request("ContractState"),10)
    reqCtrSearch    = requestCheckvar(request("reqCtrSearch"),10)
    reqCtrSearch = "P"
    reqCtr          = requestCheckvar(request("reqCtr"),10) 
    notboru         = requestCheckvar(request("notboru"),10)
    ctrType         = requestCheckvar(request("ctrType"),10)
	selSP           = requestCheckvar(request("selSP"),10)
	
	nreguserid      = requestCheckVar(request("nrU"),32)
 	ncontractNo     = requestCheckVar(request("ncontractNo"),20)
 	nContractState  = requestCheckVar(request("nContractState"),10)
    nreqCtrSearch   = requestCheckvar(request("nreqCtrSearch"),10)
    nreqCtrSearch ="P"
    nreqCtr         = requestCheckvar(request("nreqCtr"),10) 
    nnotboru        = requestCheckvar(request("nnotboru"),10)
    nctrType        = requestCheckvar(request("nctrType"),10)
    nselSP           = requestCheckvar(request("nselSP"),10)
'	catecode = requestCheckvar(request("catecode"),10) 
'    grpType = requestCheckvar(request("grpType"),10) 

    arrgroupid = split(request("arrgid") ,",")
strParm = "makerid="&makerid&"&dispcate="&dispcate&"&arect="&arect&"&crect="&crect&"&rU="&reguserid&"&contractNo="&contractNo&"&ContractState="&ContractState&"&ctrType="&ctrType&"&nrU="&nreguserid&"&ncontractNo="&ncontractNo&"&nContractState="&nContractState&"&nctrType="&nctrType&"&iC="&iCurrpage&"&selSP="&selSP&"&nselSP="&nselSP&"&rDU="&regDefUserid
set cCtrList = new CCtrNew
		cCtrList.FCPage = iCurrpage		'����������
		cCtrList.FPSize = iPageSize		'���������� ���̴� ���ڵ尹��

 		cCtrList.FRectDispCateCode = dispCate
     	cCtrList.FRectMakerid = makerid
     	cCtrList.FRectCompanyName = arect
     	cCtrList.FRectGroupID = crect
     	
     	cCtrList.FRectregDefuserid    =   regDefUserid    
        cCtrList.FRectreguserid       =   reguserid      
        cCtrList.FRectcontractNo      =   contractNo    
        cCtrList.FRectContractState   =  ContractState 
        cCtrList.FRectreqCtrSearch    =  reqCtrSearch  
        cCtrList.FRectreqCtr          =  reqCtr        
        cCtrList.FRectnotboru         =  notboru       
        cCtrList.FRectctrType         =  ctrType       
        cCtrList.FRectselSP          =  selSP             
                                                     
        cCtrList.FRectnreguserid       =  nreguserid    
        cCtrList.FRectncontractNo      =  ncontractNo   
        cCtrList.FRectnContractState   =  nContractState
        cCtrList.FRectnreqCtrSearch    =  nreqCtrSearch 
        cCtrList.FRectnreqCtr          =  nreqCtr       
        cCtrList.FRectnnotboru         =  nnotboru      
        cCtrList.FRectnctrType         =  nctrType      
        cCtrList.FRectnselSP          =  nselSP   
        
		arrList = cCtrList.fnGetCtrList
		iTotCnt = cCtrList.FTotCnt
		iTotGCnt = cCtrList.FgroupCnt
set cCtrList = nothing		

iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '��ü ������ ��
%>
<script type="text/javascript">
function regContract(makerid,groupid){
    var popwin = window.open('ctrReg.asp?makerid=' + makerid + '&groupid=' + groupid,'contractReg','width=1124,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}

//��ü ����
function jsChkAll(){
var frm;
frm = document.frmList;
	if (frm.chkAll.checked){
	   if(typeof(frm.chkid) !="undefined"){
	   	   if(!frm.chkid.length){
	   	   	if(frm.chkid.disabled==false){
		   	 	frm.chkid.checked = true;
		   	}
		   }else{
				for(i=0;i<frm.chkid.length;i++){
					 	if(frm.chkid[i].disabled==false){
					frm.chkid[i].checked = true;
				}
			 	}
		   }
	   }
	} else {
	  if(typeof(frm.chkid) !="undefined"){
	  	if(!frm.chkid.length){
	   	 	frm.chkid.checked = false;
	   	}else{
			for(i=0;i<frm.chkid.length;i++){
				frm.chkid[i].checked = false;
			}
		}
	  }

	}

}

//�űԵ��
function jsCtrReg(){
    if (confirm("�����Ͻ� �׷��ڵ��� ��ü��༭�� ����Ͻðڽ��ϱ�?")){
    document.frmList.hidM.value ="I";
    document.frmList.submit();
}
}

//����&�߼�
function jsCtrOpen(){
     if (confirm("�����Ͻ� �׷��ڵ��� ��༭�� �����ϰ� ���� �߼��Ͻðڽ��ϱ�?")){
     document.frmList.hidM.value ="M";
    document.frmList.submit();
}
}

//������� 
function jsCtrClose(){
     if (confirm("�����Ͻ� �׷��ڵ��� ��༭�� �����Ͻðڽ��ϱ�?")){
     document.frmList.hidM.value ="D";
    document.frmList.submit();
}
}

//��� ��������
function jsDivCtrClose(ctridx){
     if (confirm("�����Ͻ� ��༭�� �����Ͻðڽ��ϱ�?")){
     document.frmList.hidM.value ="P";
     document.frmList.hidCI.value = ctridx;
    document.frmList.submit();
} 
}
</script>
<!-- �˻� ���� --> 
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>" border="0">
<form name="frmSearch" method="get" action=""> 
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="5" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left" colspan="2">
		�귣������ :
    		<% drawSelectBoxDesignerwithName "makerid",makerid %>&nbsp;
        &nbsp;&nbsp;
        ȸ���/����ڹ�ȣ : <input type="text" name="arect" value="<%= arect %>" Maxlength="32" size="16" <%=CHKIIF(reqCtrSearch="N","disabled","") %>>
		&nbsp;&nbsp;
	    �׷��ڵ� : <input type="text" name="crect" value="<%= crect %>" Maxlength="32" size="16"> 
	     &nbsp;&nbsp;
		<span style="white-space:nowrap;">����ī�װ� : <% CALL DrawSelectBoxDispCateLarge("dispCate",dispCate,"")%></span>
		<br>
	</td>
	<td rowspan="5" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frmSearch.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
    <td >�������</td>
	<td align="left">
	    �ŷ��⺻�������:<input type="Text" name="rDU" value="<%=regDefUserid%>" size="20" class="text">
	     &nbsp;&nbsp;
	    ��� �����:<input type="Text" name="rU" value="<%=regUserid%>" size="20" class="text">
	    &nbsp;&nbsp;
		��༭��ȣ : <input type="text" name="contractNo" value="<%= contractNo %>" Maxlength="32" size="16" <%=CHKIIF(reqCtrSearch="N","disabled","") %>> 
		&nbsp;&nbsp; 
    	   ��� ������� :
   <!-- <input type="radio" name="reqCtrSearch" id="reqCtrSearch1" value="P" <%=CHKIIF(reqCtrSearch="P","checked","") %> ><label for="reqCtrSearch1">���������</label>-->
	<select name="ContractState" <%=CHKIIF(reqCtrSearch<>"P","disabled","") %>>
	<option value="">��ü
	<option value="M" <% if ContractState="M" then response.write "selected" %> >�̿Ϸ���ü
	<option value="0" <% if ContractState="0" then response.write "selected" %> >������
	<option value="1" <% if ContractState="1" then response.write "selected" %> >��ü����
	<option value="3" <% if ContractState="3" then response.write "selected" %> >��üȮ��
	<option value="7" <% if ContractState="7" then response.write "selected" %> >���Ϸ�
	<option value="-1" <% if ContractState="-1" then response.write "selected" %> >����
	</select>
		&nbsp;&nbsp;
		��༭ ����
    	<select name="ctrType">
    	<option value="">��ü
    	<option value="8" <%=CHKIIF(ctrType="8","selected","") %> >�⺻��༭
    	<option value="9" <%=CHKIIF(ctrType="9","selected","") %> >������,��ǰ��������
    	<option value="10" <%=CHKIIF(ctrType="10","selected","") %> >��ǰ���ް�༭
    	</select>
    	&nbsp;&nbsp;
    	�Ǹ�ó��
    	<select name="selSP">
    	  <option value="">��ü</option>    
    	  <option value="on" <%=CHKIIF(selSP="on","selected","") %>>�¶���</option>  
    	  <option value="off" <%=CHKIIF(selSP="off","selected","") %>>��������</option>
    	</select>
	</td>
</tr>
<!--<tr align="center" bgcolor="#FFFFFF" >
    <td align="left" height="30">
 
	
  <!--  <input type="radio" name="reqCtrSearch" id="reqCtrSearch2" value="N" <%=CHKIIF(reqCtrSearch="N","checked","") %> ><label for="reqCtrSearch2">�̰��</label>
    <select name="reqCtr" <%=CHKIIF(reqCtrSearch<>"N","disabled","") %> >
    <option value="OJ" <% if reqCtr="OJ" then response.write "selected" %> >�¶��� <%=jMonth%>���� �������,�ǸŻ�ǰ>0
    <option value="OT" <% if reqCtr="OT" then response.write "selected" %> >�¶��� <%=jMonth%>���� �������,���������������,�ǸŻ�ǰ>0
    <option value="OJN" <% if reqCtr="OJN" then response.write "selected" %> >�¶��� <%=jMonth%>���� �������,�����0, �ǸŻ�ǰ>0
    <option value="OJNN" <% if reqCtr="OJNN" then response.write "selected" %> >�¶��� <%=jMonth%>���� �������,�����0, �ǸŻ�ǰ=0

    <option value="FJ" <% if reqCtr="FJ" then response.write "selected" %> >�������� <%=jMonth%>���� �������,�ǸŻ�ǰ>0
    <option value="FN" <% if reqCtr="FN" then response.write "selected" %> >�������� <%=jMonth%>���� �������,�¶����������,�ǸŻ�ǰ>0
    </select>
 
   <span id="dvBoru" style="display:<%=CHKIIF(reqCtrSearch="N","","none") %>"><input type="checkbox" name="notboru" <%=CHKIIF(notboru="on","checked","")%> >�����귣�� ǥ�þ���</span>-->
    
    </td>
</tr> 
<tr align="center" bgcolor="#FFFFFF" >
    <td  >�ű԰��</td>
	<td align="left">
	    ��� �����:<input type="Text" name="nrU" value="<%=nregUserid%>" size="20" class="text">
	    &nbsp;&nbsp;
		��༭��ȣ : <input type="text" name="ncontractNo" value="<%= ncontractNo %>" Maxlength="32" size="16" <%=CHKIIF(nreqCtrSearch="N","disabled","") %>>
 
		&nbsp;&nbsp;
		 ��� ������� :
 <!--   <input type="radio" name="nreqCtrSearch" id="nreqCtrSearch1" value="P" <%=CHKIIF(nreqCtrSearch="P","checked","") %> ><label for="nreqCtrSearch1">���������</label>-->
	<select name="nContractState" <%=CHKIIF(nreqCtrSearch<>"P","disabled","") %>>
	<option value="">��ü
	<option value="M" <% if nContractState="M" then response.write "selected" %> >�̿Ϸ���ü
	<option value="0" <% if nContractState="0" then response.write "selected" %> >������
	<option value="1" <% if nContractState="1" then response.write "selected" %> >��ü����
	<option value="3" <% if nContractState="3" then response.write "selected" %> >��üȮ��
	<option value="7" <% if nContractState="7" then response.write "selected" %> >���Ϸ�
	<option value="D" <% if nContractState="D" then response.write "selected" %> >�������
	<option value="-1" <% if nContractState="-1" then response.write "selected" %> >����
	</select>
	&nbsp;&nbsp;
		��༭ ����
    	<select name="nctrType">
    	<option value="">��ü
    	<option value="11" <%=CHKIIF(nctrType="11","selected","") %> >�ŷ��⺻��༭
    	<option value="12" <%=CHKIIF(nctrType="12","selected","") %> >�ŷ��⺻���μ����Ǽ�
    	<option value="13" <%=CHKIIF(nctrType="13","selected","") %> >�����԰�༭
    	<option value="14" <%=CHKIIF(nctrType="14","selected","") %> >�����Ժμ����Ǽ�
    	</select>
    	&nbsp;&nbsp;
    	<!--�Ǹ�ó��
    	<select name="nselSP">
    	  <option value="">��ü</option>    
    	  <option value="on" <%=CHKIIF(nselSP="on","selected","") %>>�¶���</option>  
    	  <option value="off" <%=CHKIIF(nselSP="off","selected","") %>>��������</option>
    	</select>
    	-->
	
	</td>
</tr>
<!--<tr align="center" bgcolor="#FFFFFF" >
    <td align="left" height="30">
   

 <!--  <input type="radio" name="nreqCtrSearch" id="nreqCtrSearch2" value="N" <%=CHKIIF(nreqCtrSearch="N","checked","") %> ><label for="nreqCtrSearch2">�̰��</label>
    <select name="nreqCtr" <%=CHKIIF(nreqCtrSearch<>"N","disabled","") %> >
    <option value="OJ" <% if nreqCtr="OJ" then response.write "selected" %> >�¶��� <%=jMonth%>���� �������,�ǸŻ�ǰ>0
    <option value="OT" <% if nreqCtr="OT" then response.write "selected" %> >�¶��� <%=jMonth%>���� �������,���������������,�ǸŻ�ǰ>0
    <option value="OJN" <% if nreqCtr="OJN" then response.write "selected" %> >�¶��� <%=jMonth%>���� �������,�����0, �ǸŻ�ǰ>0
    <option value="OJNN" <% if nreqCtr="OJNN" then response.write "selected" %> >�¶��� <%=jMonth%>���� �������,�����0, �ǸŻ�ǰ=0

    <option value="FJ" <% if nreqCtr="FJ" then response.write "selected" %> >�������� <%=jMonth%>���� �������,�ǸŻ�ǰ>0
    <option value="FN" <% if nreqCtr="FN" then response.write "selected" %> >�������� <%=jMonth%>���� �������,�¶����������,�ǸŻ�ǰ>0
    </select>

    <span id="ndvBoru" style="display:<%=CHKIIF(nreqCtrSearch="N","","none") %>"><input type="checkbox" name="nnotboru" <%=CHKIIF(nnotboru="on","checked","")%> >�����귣�� ǥ�þ���</span>
    
    </td>
</tr>-->
</form>
</table>
<!-- �˻� �� -->
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:5;padding-bottom:5;">  
    <tr>
		<td align="right">
		     <input type="button" style="color:blue;" value="�ű� ��� ���" onClick="jsCtrReg()" class="button">
		    <input type="button" value="��༭ ����&�߼�" onClick= "jsCtrOpen()" class="button">
		     &nbsp;&nbsp;&nbsp;
             <input type="button" value="��� ����" onClick="jsCtrClose()" class="button"> 
		    <!--<input type="button" value=" ��� �Ϸ�" onClick="regContract('<%=makerid%>','<%=uniqGroupID%>')" class="button">-->
	</tr> 
</table>
<form name="frmList" method="post" action="procNewCtr.asp?<%=strParm%>">
<input type="hidden" name="hidM" value=""> 
<input type="hidden" name="hidCI" value="">
<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="23" align="right">�׷�: <%=formatnumber(iTotGCnt,0)%>�� &nbsp;��:<%=formatnumber(iTotCnt,0)%>��  &nbsp; <%=formatnumber(iCurrPage,0)%>/<%=formatnumber(iTotalPage,0)%> page</td>
</tr>  
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td rowspan="2"><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
    <td rowspan="2">�׷��ڵ�</td>
    <td rowspan="2">��ü��</td>
    <td rowspan="2">�귣��ID</td>
    <td rowspan="2">����ī�װ�</td>
    <td rowspan="2">�Ǹ�ó</td>
    <td rowspan="2">�������</td>
    <td colspan="3">�Ǹ����λ�ǰ��</td>
    <td colspan="6">���� ��༭</td>
    <td colspan="6">�ű� ��༭</td>
    <td rowspan="2">���</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>����</td>
    <td>��Ź</td>
    <td>��ü</td>
    
    <td>��༭��</td>
    <td>�����</td>
    <td>����</td>
    <td>�����</td>
    <td>�߼���</td>
    <td>�Ϸ���</td>
    
    <td>��༭��</td>
    <td>�����</td>
    <td>����</td>
    <td>�����</td>
    <td>�߼���</td>
    <td>�Ϸ���</td>
 </tr>   
 <%IF isArray(arrList) THEN%>
 <%intG = 0
 oldgroupid =0
 %>
 <% For intLoop = 0 To UBound(arrList,2)
     if not ( nContractState = "M" and oldgroupid <> arrList(0,intLoop) and arrList(32,intLoop)=10) then 
 %>
 <tr align="center"  bgcolor="#ffffff"> 
    <td><% if oldgroupid <> arrList(0,intLoop) then%><input type="checkbox" name="chkid" value="<%= arrList(0,intLoop) %>" <%if ubound(arrgroupid) >= 0 then%><%if Cstr(trim(arrList(0,intLoop))) = Cstr(trim(arrgroupid(intG))) then%>checked<%if intG < ubound(arrgroupid) then intG = intG+1 end if%><%end if%><%end if%>><%end if%></td>
    <td><a href="javascript:regContract('<%=arrList(2,intLoop) %>','<%= arrList(0,intLoop) %>');"><%=arrList(0,intLoop)%></a></td>
    <td><%=arrList(1,intLoop)%></td>
    <td><%=arrList(2,intLoop)%></td>
    <td><%=arrList(9,intLoop)%></td>
    <td><%=arrList(3,intLoop)%></td>
    <td><%=fnMaeipdivName(arrList(4,intLoop))%></td>
    <td><%=arrList(5,intLoop)%></td>
    <td><%=arrList(6,intLoop)%></td>
    <td><%=arrList(7,intLoop)%></td>
    <td><%=arrList(10,intLoop)%><br>
         <%if not isNull(arrList(11,intLoop)) then%><font color="gray">[<%=arrList(11,intLoop)%>]</font><%end if%>  
    </td>
    <td><%if not isNull(arrList(13,intLoop)) then %><%=FormatDate(arrList(13,intLoop),"0000-00-00")%><%end if%></td>
    <td><%=GetContractStateName(arrList(12,intLoop))%></td>
    <td><%=arrList(15,intLoop)%></td>
    <td><%=arrList(17,intLoop)%></td>
    <td><%=arrList(19,intLoop)%></td>
    
    <td><%=arrList(28,intLoop)%><br>
         <%if not isNull(arrList(21,intLoop)) then%><font color="gray">[<%=arrList(21,intLoop)%>]</font><%end if%>
        </td>
    <td><%=arrList(23,intLoop)%></td>
    <td><%if nContractState="D" then%>�������<%else%><%=GetContractStateName(arrList(22,intLoop))%><%end if%></td>
    <td><%=arrList(29,intLoop)%></td>
    <td><%=arrList(30,intLoop)%></td>
    <td><%=arrList(31,intLoop)%></td>
    <td><input type="button" class="button" value="�������" onClick="jsDivCtrClose('<%=arrList(27,intLoop)%>');"><br/><font color="white" size="1"><%=arrList(27,intLoop)%></td>
</tr>
<% 
   oldgroupid = arrList(0,intLoop)
   end if
 Next%>
 <%ELSE%>
 <tr  align="center"  bgcolor="#ffffff">
    <td colspan="23" >��ϵ� ������ �����ϴ�.</td>
</tr>
 <%END IF%>
 
</table>
</form>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:5;padding-bottom:5;"> 
	<tr>
		<td align="right">
		   
		   
		    <input type="button" style="color:blue;" value="�ű� ��� ���" onClick="jsCtrReg()" class="button">
		    <input type="button" value="��༭ ����&�߼�" onClick= "jsCtrOpen()" class="button">
		     &nbsp;&nbsp;&nbsp;
             <input type="button" value="��� ����" onClick="jsCtrClose()" class="button"> 
		    <!--<input type="button" value=" ��� �Ϸ�" onClick="regContract('<%=makerid%>','<%=uniqGroupID%>')" class="button">-->
	</tr> 
</table>
<!-- ����¡ó�� --> 
<table width="100%" cellpadding="10">
	<tr>
		<td align="center">  
 			<%sbDisplayPaging "iC", iCurrPage, iTotCnt, iPageSize, 10,menupos %>
		</td>
	</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
