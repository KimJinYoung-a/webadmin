<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ü ��� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls2013.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/ecContractApi_function.asp"-->
<%
dim groupid : groupid= requestCheckvar(request("groupid"),10)

dim ocontract
set ocontract = new CPartnerContract
	ocontract.FPageSize=200
	ocontract.FCurrPage = 1
	ocontract.FRectGroupID = groupid
	'ocontract.FRectMakerid = makerid
	'ocontract.FRectContractno  = contractNo
	'ocontract.FRectContractState = ContractState
	ocontract.GetNewContractList

dim ogroupInfo
SET ogroupInfo = new CPartnerGroup
ogroupInfo.FRectGroupid = groupid
if (groupid<>"") then
    ogroupInfo.GetOneGroupInfo

    if (ogroupInfo.FResultCount<1) then
        response.write "�ش� ��ü�׷� ������ �����ϴ�. "&groupid
        dbget.close(): response.end
    end if
end if

''���� ���� ����Ʈ
dim oAddContractList
set oAddContractList = new CPartnerContract
oAddContractList.FPageSize=30
oAddContractList.FCurrPage = 1
oAddContractList.FRectGroupID = groupid
if (groupid<>"") then
    oAddContractList.GetCurrAddContractListCheckMargin
end if


dim i, noMatCnt, dsbleCnt
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="contract.js?v=1.00"></script>
<script language='javascript'>
$(document).on('change','input[name="chkAll"]',function() {
    $('.idChk').prop("checked" , this.checked);
});

function checkALL(comp){

}


function sendContract(inoMatCnt, dsbleCnt){
<% if session("ssbctId")<>"icommang" then %>
//alert('���� ������ �� �����ϴ�.');
//return;
<% end if %>
    var frm=document.frmOpen;

    var chkCnt = $('.idChk:checked').length;

    if (frm.dsbleCnt.value*1>0){
        alert('��� �Ұ��� ������ �ֽ��ϴ�.\n\nȸ�� �� ����');
        return;
    }

    if (chkCnt<1){
        alert('���õ� ��༭�� �����ϴ�.');
        frm.chkAll.focus();
        return;
    }

    if ((frm.noMatCnt.value*1>0)&&(!confirm('���� SCM ���� ������ ��ġ���� �ʴ� ������ �ֽ��ϴ�.\n\n��� �Ͻðڽ��ϱ�?\n\n��� �Ͻô°�� ��ึ������ ���� SCM������ ����˴ϴ�.'))){
        return;
    }

 var signtype="";
   for(var i=0;i<frm.chkCtr.length;i++){
   	 	if (frm.chkCtr[i].checked){
	 		if (signtype ==""){
	   	 		signtype = frm.hidst[i].value;
	   	}else{
	   	 	 if (frm.hidst[i].value != signtype){
	   	 	 	alert("��༭ Ÿ���� ������ ��༭�� �ϰ� �߼� �����մϴ�.���ڿ� ���� �Ǵ� DocuSign ��༭�� ���� �߼����ּ���");
	   	 	 	return;
	   	 	 }
	     	}
	   	  frm.signtype.value = signtype;
  	 	}    	
   }

    <%'' DocuSign �� ��� ctropen�� �ƴ� ctropendocusign ���� ctrReg_Process.asp �������� ������. %>
    if (frm.signtype.value=="3") {
        frm.mode.value="ctropendocusign";
        $("#submitButton").attr('disabled',true);
    }

    if (confirm('���� ��༭�� �߼�(����)�Ͻðڽ��ϱ�?')){
        frm.submit();
    }

}

function jsEcSubmit(ecCtrSeq, companyno){
	document.frmecView.cont_seq.value = ecCtrSeq;
	document.frmecView.corp_id.value = companyno;
	document.frmecView.submit();
}

</script>

<table width="100%" border="0" cellspacing="1" cellpadding="4" class="a" bgcolor="#BABABA">
<form name="frm" method="get" action="">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">

    ��ü�׷��ڵ� : <input type="text" class="text" name="groupid" value="<%= groupid %>" size="10" >
    <input type="button" class="button" value="Code�˻�" onclick="popSearchGroupID(this.form.name,'groupid');" >
    &nbsp;&nbsp;
    </td>
    <td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
</form>
</table>
<p>
<form name="frmecView" method="post" action="<%=FecUrl%>/w20/contractView.do" target="_blank"> 
	<input type="hidden" name="remote_id" value="<%=FecID%>" />  <!-- �ۼ��� LOGIN ID -->
	<input type="hidden" name="cont_seq" value="" />  <!-- ��༭ ��ȣ -->
	<input type="hidden" name="corp_id" value="" /> <!-- ����� ȭ���Ϸ��� ����ڹ�ȣ -->
 </form> 
<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmOpen" method="post" action="ctrReg_Process.asp">
<input type="hidden" name="groupid" value="<%=groupid%>">
<input type="hidden" name="mode" value="ctropen">
<input type="hidden" name="reguserid" value="<%=session("ssBctID")%>">
<input type="hidden" name="signtype" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20" ><input type="checkbox" name="chkAll"></td>
    <td width="50" >��༭Ÿ��</td>
    <td width="110" >��༭ ��</td>
    <td width="100" >��༭��ȣ</td>
    <td width="80" >��ü��</td>
    <td width="100" >�귣��ID</td>
    <td width="100" >�Ǹ�ó</td>
    <td width="100" >����</td>
    <td width="70" >�����</td>
    <td width="70" >����</td>
    <td width="80" >�����</td>
    <td width="70" >�����</td>
    <td width="90" >���</td>
</tr>
<% dim signtype: signtype = 0

 if ocontract.FResultCount>0 then %>
<% for i=0 to ocontract.FResultCount - 1 %>
<tr bgcolor="#FFFFFF">
    <td ><input type="checkbox" class="<%=CHKIIF(ocontract.FITemList(i).IsCtrOpenValidState,"idChk","")%>" name="chkCtr" value="<%= ocontract.FITemList(i).FctrKey %>" <%=CHKIIF(ocontract.FITemList(i).IsCtrOpenValidState,"","disabled")%> onClick="checkALL(this)"></td>
    <td><%if ocontract.FItemList(i).FecCtrSeq <> "" and not isNull(ocontract.FItemList(i).FecCtrSeq) and ocontract.FItemList(i).FecCtrSeq <> "0" then%>
    		����(<%=ocontract.FItemList(i).FecCtrSeq %>)
    		<%signtype =2 %>
    		<%else%>
                <%If ocontract.FItemList(i).FsignType = "D" Then %>
                    DocuSign
                    <%signtype =3 %>
                <% Else %>
    		        ����
                    <%signtype =1 %>                    
                <% End If %>
    		<%end if%>
    	<input type="hidden" name="hidst" value="<%=signtype%>">
    </td>
    <td><%= ocontract.FITemList(i).FContractName %></td>
    <td align="center"><a href="javascript:modiContract('<%= ocontract.FITemList(i).FctrKey %>');"><%= CHKIIF(isNULL(ocontract.FITemList(i).FctrNo) or ocontract.FITemList(i).FctrNo="","-",ocontract.FITemList(i).FctrNo) %></a></td>
    <td><%= ocontract.FITemList(i).FcompanyName %></td>
    <td><%= ocontract.FITemList(i).FMakerid %></td>
    <td align="center"><%= ocontract.FITemList(i).getMajorSellplaceName %></td>
    <td align="center"><%= ocontract.FITemList(i).getMajorMarginStr %></td>
    <td align="center"><%= ocontract.FITemList(i).FcontractDate %></td>
    <td align="center"><font color="<%= ocontract.FITemList(i).GetContractStateColor %>"><%= ocontract.FITemList(i).GetContractStateName %></font></td>
    <td align="center"><%= ocontract.FITemList(i).FRegUserName %></td>
    <td ><%= LEFT(ocontract.FITemList(i).FregDate,10) %></td>
    <td align="center">
    	<%if  ocontract.FITemList(i).FecCtrSeq <>"" and not  isNull(ocontract.FITemList(i).FecCtrSeq ) and ocontract.FITemList(i).FecCtrSeq <> "0" then%>
        <img src="/images/documents_icon.png" style="cursor:pointer;" onClick="jsEcSubmit('<%=ocontract.FITemList(i).FecCtrSeq%>','<%= replace(ocontract.FItemList(i).FcompanyNo,"-","")%>');" >
         
        <%end if%>
        <% If ocontract.FITemList(i).FsignType = "D" Then %>
            <img src="/images/browser_icon.png" style="cursor:pointer;" onClick="dnWebAdmDocu('<%=ocontract.FITemList(i).FctrKey %>');">
        <% Else %>
            <img src="/images/browser_icon.png" style="cursor:pointer;" onClick="dnWebAdm('<%=ocontract.FITemList(i).FctrKey %>');">
        <% End If %>
        <img src="/images/pdf_icon.png" style="cursor:pointer;" onClick="dnPdfAdm('<%=ocontract.FITemList(i).getPdfDownLinkUrlAdm %>');">
    </td>
</tr>
<% next %>
<% else %>
<tr bgcolor="#FFFFFF">
    <td colspan="13" align="center">������ ��༭�� �����ϴ�.</td>
</tr>
<% end if %>
</table>
<p>

<!-- ���� ���� --> 
<table width="100%" border="0" cellspacing="1" cellpadding="4" class="a" bgcolor="#BABABA">
<tr bgcolor="<%= adminColor("gray") %>" align="center">
    <td colspan="4">�������</td>
    <td colspan="4">SCM��������</td>
    <td colspan="4">��ǰ����</td>
    <td rowspan="2">���</td>
</tr>
<tr bgcolor="<%= adminColor("gray") %>" align="center">
    <td>�귣��ID</td>
    <td>�Ǹ�ó</td>
    <td>���Ա���</td>
    <td>��ึ��</td>


    <td>���Ա���</td>
    <td>��ึ��</td>

    <td>��ǥ���Ա���</td>
    <td>��ǥ��ึ��</td>

    <td>����</td>
    <td>����</td>
    <td>�Ǹż�</td>
    <td>����</td>

</tr> 
<%
noMatCnt=0
dsbleCnt=0
%>

<% for i=0 to oAddContractList.FresultCount-1 %>
<%
if (oAddContractList.FItemList(i).isreqCheckMargin) then
    noMatCnt=noMatCnt+1
end if

if (oAddContractList.FItemList(i).isDisabledMWMargin) then
    dsbleCnt=dsbleCnt+1
end if
%> 
<tr bgcolor="<%=CHKIIF(oAddContractList.FItemList(i).isDisabledMWMargin,"#CCCCCC","#FFFFFF")%>" align="center">

    <td bgcolor="#FFFFFF" ><%=oAddContractList.FItemList(i).FMakerid %></td>
    <td bgcolor="#FFFFFF" ><%=oAddContractList.FItemList(i).getSellplaceName %></td>
    <td <%=CHKIIF(oAddContractList.FItemList(i).isreqCheckMW,"bgcolor='#DD7777'","")%> ><%=oAddContractList.FItemList(i).getContractMwDivStr %></td>
    <td <%=CHKIIF(oAddContractList.FItemList(i).isreqCheckMargin,"bgcolor='#DD7777'","")%> ><%=oAddContractList.FItemList(i).getContractMarginStr %></td>

    <td><%=fnMaeipdivName(oAddContractList.FItemList(i).FMaeipdiv) %></td>
    <td><%=oAddContractList.FItemList(i).getSCMDefaultmargineStr %></td>

    <% if (oAddContractList.FItemList(i).Fsellplace="ON") then %>
    <td><% if (oAddContractList.FItemList(i).Fcontractmwdiv=oAddContractList.FItemList(i).FMjmaeipdiv) then %><%=fnMaeipdivName(oAddContractList.FItemList(i).FMjmaeipdiv) %><% end if %></td>
    <td><% if (oAddContractList.FItemList(i).Fcontractmwdiv=oAddContractList.FItemList(i).FMjmaeipdiv) then %><%=oAddContractList.FItemList(i).getMjContractMarginStr %><% end if %></td>
    <% else %>
    <td><%=fnMaeipdivName(oAddContractList.FItemList(i).FMjmaeipdiv) %></td>
    <td><%=oAddContractList.FItemList(i).getMjContractMarginStr %></td>
    <% end if %>

    <td><%=FormatNumber(oAddContractList.FItemList(i).FuseitemCnt,0) %></td>
    <td><%=CLNG(oAddContractList.FItemList(i).Fuseitemmargin*100)/100 %></td>
    <td><%=FormatNumber(oAddContractList.FItemList(i).FsellitemCnt,0) %></td>
    <td><%=CLNG(oAddContractList.FItemList(i).Fsellitemmargin*100)/100 %></td>
    <td>

    </td>
</tr> 
<% next %>
<input type="hidden" name="noMatCnt" value="<%=noMatCnt%>">
<input type="hidden" name="dsbleCnt" value="<%=dsbleCnt%>">
</table> 
<p>
<% if ocontract.FResultCount>0 then %>
<table width="100%" border="0" cellspacing="1" cellpadding="4" class="a" bgcolor="#FFFFFF">
<tr bgcolor="#FFFFFF">
    <td align="center">
        <table width="80%" border="0" cellspacing="1" cellpadding="4" class="a" bgcolor="#BABABA">
        <tr bgcolor="#FFFFFF">
            <td width="100">�������</td>
            <td><%=session("ssBctCName")%></td>
            <td></td>
            <td></td>
            <td></td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td>��������</td>
            <td><%=ogroupInfo.FOneItem.Fmanager_name%></td>
            <td><%=ogroupInfo.FOneItem.Fmanager_phone%></td>
            <td><input type="checkbox" name="ckHp" value="on" checked >�߼� <input type="text" name="mngHp" value="<%=ogroupInfo.FOneItem.Fmanager_hp%>" class="text" size="15"></td>
            <td><input type="checkbox" name="ckEmail" value="on" checked >�߼� <input type="text" name="mngEmail" value="<%=ogroupInfo.FOneItem.Fmanager_email%>" class="text" size="22"></td>
        </tr>
        </table>
    </td>
</tr>
<input type="hidden" name="mngName" value="<%=ogroupInfo.FOneItem.Fmanager_name%>">
</table>
<% end if %>
<p>
<table width="100%" border="0" cellspacing="1" cellpadding="4" class="a" bgcolor="#FFFFFF">
<tr bgcolor="#FFFFFF">
    <td align="center">
    <% if ocontract.FResultCount>0 then %>
    <input type="button" class="button" value="���� ��༭ �߼�" id="submitButton" name="submitButton" onClick="sendContract()">
    &nbsp
    <input type="button" class="button" value="�̸��� �̸�����" onClick="preViewSendContract('<%=groupid%>','1')">
    &nbsp
    <input type="button" class="button" value="���ڰ�� �̸��� �̸�����" onClick="preViewSendContract('<%=groupid%>','2')">
    <% else %>
    <input type="button" class="button" value="�ݱ�" onClick="window.close()">
    <% end if %>
    </td>
</tr>
</form>
</table>

<%
SET oAddContractList=Nothing
SET ogroupInfo=Nothing
SET ocontract=Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->