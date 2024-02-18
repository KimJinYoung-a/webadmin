<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �귣�� ��� ����
' Hieditor : 2010.05.25 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->

<%
dim contractType, contractID, mode , makerid ,i , sqlStr, marginRows
	makerid         = request("makerid")
	contractType    = request("contractType")
	contractID      = request("contractID")
	mode            = request("mode")

dim opartner
set opartner = new CPartnerUser
opartner.FRectDesignerID = makerid

if (makerid<>"") then
    opartner.GetOnePartnerNUser
end if

''�귣�� ��༭ ����Ʈ
dim ocontractList
set ocontractList = new CPartnerContract
ocontractList.FRectMakerid = makerid
if (makerid<>"") then
    ocontractList.GetContractList
end if

''������ ��༭ or �������� ��༭
dim ocontract, ocontractDetail
set ocontract = new CPartnerContract
ocontract.FRectContractID = ContractID
ocontract.FRectMakerID = makerid

if (ContractID<>"") then
    ocontract.GetOneContract
elseif (mode="") then
    ocontract.GetLastOneContract
end if

if ocontract.FResultCount>0 then
    ContractID = ocontract.FOneItem.FContractID
end if

set ocontractDetail = new CPartnerContract
ocontractDetail.FRectContractID = ContractID
if (ContractID<>"") then
    ocontractDetail.GetContractDetailList
end if

'' ���õ�(��������) ����� �ִ°��
dim CONTRACTING_EXISTS
CONTRACTING_EXISTS = ocontract.FresultCount>0

'' ������ ����� ���°�� : ��༭ ProtoType �⺻ Setting
if (Not CONTRACTING_EXISTS) and (opartner.FResultCount>0) and (ContractType="") then
    if opartner.FOneItem.Fmaeipdiv="U" then
        ContractType="5"
    elseif opartner.FOneItem.Fmaeipdiv="W" then
        ContractType="1"
    elseif opartner.FOneItem.Fmaeipdiv="M" then
        ContractType="2"
    end if
else
    if ocontract.FResultCount>0 then
        ContractType = ocontract.FOneItem.FContractType
    end if
end if

dim ocontractProtoType
set ocontractProtoType = new CPartnerContract
ocontractProtoType.FRectContractType = ContractType

if (Not CONTRACTING_EXISTS) and (ContractType<>"") then
    ocontractProtoType.getOneContractProtoType
end if

dim ocontractProtoTypeDetail
set ocontractProtoTypeDetail = new CPartnerContract
ocontractProtoTypeDetail.FRectContractType = ContractType

if (Not CONTRACTING_EXISTS) and (ContractType<>"") then
    ocontractProtoTypeDetail.getContractDetailProtoType
end if

sqlStr = "select mwdiv, (100-buycash/sellcash*100) ,count(itemid) as cnt"
sqlStr = sqlStr & " from [db_item].[dbo].tbl_item"
sqlStr = sqlStr & " where itemid<>0"
sqlStr = sqlStr & " and makerid='" & makerid & "'"
sqlStr = sqlStr & " and sellcash<>0"
sqlStr = sqlStr & " and sellyn='Y'"
sqlStr = sqlStr & " and isusing='Y'"
sqlStr = sqlStr & " group by mwdiv, (100-buycash/sellcash*100)"
if makerid<>"" then
    rsget.Open sqlStr,dbget,1
    if  not rsget.EOF  then
        marginRows = rsget.getRows()
    end if
    rsget.close
end if
%>

<script language='javascript'>

function ChangeBrand(comp){
    var frm = document.frmReSearch;
    frm.makerid.value = comp.value;
    frm.ContractType.value = "";
    frm.submit();
}

function ChangeContractID(v){
    var frm = document.frmReSearch;
    frm.ContractType.value = "";
    frm.ContractID.value = v;
    frm.submit();
}

function ChangeContractType(comp){
    var frm = document.frmReSearch;
    frm.ContractType.value = comp.value;
    frm.submit();
}

function NewContractReg(){
alert('��� ���� �޴� �Դϴ�. ��Ʈ�ʰ��� >> ��ü������ �� ����ϼ���')
return;
    var frm = document.frmReSearch;
    frm.ContractType.value = "";
    frm.mode.value = "RegContract";
    frm.submit();
}

function SaveContract(frm){

    if (frm.makerid.value.length<1){
        alert('�귣�� ���̵� �����ϼ���.');
        frm.makerid.focus();
        return;
    }

    if (frm.contractType.value.length<1){
        alert('��༭ ������ �����ϼ���.');
        frm.contractType.focus();
        return;
    }

    //�ӽ�
    if (frm.contractType.value=="2"){
        alert('���� ���԰�༭�� �������� �ʽ��ϴ�.');
        frm.contractType.focus();
        return;
    }

    for (var i=0;i<frm.elements.length;i++){

        if (frm.elements[i].type=="text"){
            if (frm.elements[i].value.length<1){
                alert('�ʼ� �Է� �����Դϴ�.');
                frm.elements[i].focus();
                return;
            }
        }
    }

    if (confirm('��༭�� ����Ͻðڽ��ϱ�?')){
		frm.action = 'contractReg_Process.asp';
        frm.submit();
    }
}

function preViewContract(ContractID){
    var popwin = window.open('preViewContract.asp?ContractID=' + ContractID,'preViewContract','width=900,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function DocDownloadContract(ContractID){
    var popwin = window.open('DocDownloadContract.asp?ContractID=' + ContractID,'DocDownloadContract','width=900,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function goNextState(CurrState,NextState,confirmMsg){
    if (!confirm(confirmMsg)) return;

    var frm = document.frmReg;
    frm.mode.value = "stateChange";
    frm.CurrState.value = CurrState;
    frm.NextState.value = NextState;

    frm.submit();
}
</script>

<table width="100%" border="0" cellspacing="1" cellpadding="2" class="a" bgcolor="#BABABA">
<form name="frmReg" method="post" action="contractReg_Process.asp">
<input type="hidden" name="contractID" value="<%= contractID %>">
<input type="hidden" name="CurrState" value="">
<input type="hidden" name="NextState" value="">

<% if ContractID<>"" then %>
<input type="hidden" name="mode" value="editContract">
<% else %>
<input type="hidden" name="mode" value="regContract">
<% end if %>
<tr bgcolor="#FFFFFF">
    <td width="120" bgcolor="#DDDDFF">�귣��</td>
    <td colspan="2"><%	drawSelectBoxDesignerWithName "makerid", makerid %> <input type="button" value="����" onClick="ChangeBrand(frmReg.makerid);" class="button"></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">(��������)</td>
    <td colspan="2">
    <% if opartner.FResultCount>0 then %>
    <%= opartner.FOneItem.GetMWUName %> <%= opartner.FOneItem.Fdefaultmargine %> %
    &nbsp;&nbsp;
    <%= opartner.FOneItem.FRegDate %>
    <% end if %>

    <%
    dim rowcount
    rowcount =0
    if IsArray(marginRows) then
        rowcount = Ubound(marginRows,2) + 1
        for i=0 to rowcount-1
            response.write "<br>" & marginRows(0,i) & "," & marginRows(1,i) & "," & marginRows(2,i)
        next
    end if
    %>

    <br>
    <a href="javascript:PopBrandInfoEdit('<%= makerid %>');">[�귣����������]</a>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">���� ��ϳ���</td>
    <td colspan="2">
        <table width="100%" border="0" cellspacing="1" cellpadding="2" class="a" bgcolor="#BABABA">
        <tr bgcolor="#FFDDDD">
            <td colspan="4" align="right"><a href="javascript:NewContractReg();"><img src="/images/icon_new_registration.gif" width="75" border="0"></a></td>
        </tr>
        <tr bgcolor="#FFDDDD">
            <td width="100">��༭ ��ȣ</td>
            <td>��༭ ����</td>
            <td width="100">����</td>
            <td width="100">�����</td>
        </tr>
        <% for i=0 to ocontractList.FResultCount - 1 %>
        <tr bgcolor="#FFFFFF">
            <td><% if ocontractList.FItemList(i).FcontractID=contractID then response.write "<font color=red><b>&gt;&gt;</b></font>" %> <%= ocontractList.FItemList(i).FcontractID %></td>
            <td><a href="javascript:ChangeContractID('<%= ocontractList.FItemList(i).FcontractID %>')"><%= ocontractList.FItemList(i).FcontractName
 %></a></td>
            <td><%= ocontractList.FItemList(i).GetContractStateName %></td>
            <td><%= ocontractList.FItemList(i).Fregdate %></td>
        </tr>
        <% next %>
        </table>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">��༭����</td>
    <td><% drawSelectBoxContractTypeWithChangeEvent "contractType", contractType %></td>
    <td width="100"></td>
</tr>
<% if (CONTRACTING_EXISTS) then %>
<script language='javascript'>
    document.all.contractType.disabled=true;
</script>
<% end if %>
<%
'' ���õ�(��������) ����� �ִ°��
if ocontract.FResultCount>0 then
%>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="#DDDDFF"></td>
        <td >��༭��ȣ : <%= ocontract.FOneItem.FcontractNo %></td>
        <td>
        <a href="javascript:preViewContract('<%= ocontract.FOneItem.FcontractID %>');"><img src="/images/iexplorer.gif" width="21" border="0"></a>
        &nbsp;
        <a href="javascript:DocDownloadContract('<%= ocontract.FOneItem.FcontractID %>');"><img src="/images/btn_word.gif" width="70" border="0"></a></td>
        </td>
    </tr>
    <%
    for i=0 to ocontractDetail.FResultCount-1
    %>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="#DDDDFF"><!-- <%= ocontractDetail.FItemList(i).FDetailKey %> -->
        </td>
        <td>
        	<%
        	'//���屸��
    		if ocontractDetail.FItemList(i).FDetailKey = "$$A_STOREID$$" then
		    %>
		    	<% drawSelectOffShopmargin ocontractDetail.FItemList(i).FDetailKey , ocontractDetail.FItemList(i).FDetailValue %>
		    <%
		    '//���Ⱓ �������̳� ���Ⱓ �������� ���
		    elseif ocontractDetail.FItemList(i).FDetailKey = "$$STARTDATE$$" or ocontractDetail.FItemList(i).FDetailKey = "$$ENDDATE$$" then
		    %>
				<input type="text" name="<%=ocontractDetail.FItemList(i).FDetailKey%>" size=10 value="<%= ocontractDetail.FItemList(i).FDetailValue %>">
				<a href="javascript:calendarOpen3(frmReg.<%=ocontractDetail.FItemList(i).FDetailKey%>,'',frmReg.<%=ocontractDetail.FItemList(i).FDetailKey%>.value)">
				<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>
		    <%
		    else
		    %>
            	<input type="text" size="40" id="DetailKey" name="<%= ocontractDetail.FItemList(i).FDetailKey %>" value="<%= ocontractDetail.FItemList(i).FDetailValue %>" >
			<%
			end if
			%>
			&nbsp; <%= ocontractDetail.FItemList(i).FdetailDesc %>
        </td>
        <td><%= getDefaultContractValue(ocontractDetail.FItemList(i).FDetailKey,opartner) %></td>
    </tr>
    <%
    next
    %>
<% elseif ocontractProtoType.FResultCount>0 then %>
    <% for i=0 to ocontractProtoTypeDetail.FResultCount-1 %>
    <% if ocontractProtoTypeDetail.FItemList(i).FDetailKey<>"$$CONTRACT_NO$$" then %>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="#DDDDFF"><!-- <%= ocontractProtoTypeDetail.FItemList(i).FDetailKey %> -->
        </td>
        <td>
        	<%
        	'//���屸��
    		if ocontractProtoTypeDetail.FItemList(i).FDetailKey = "$$A_STOREID$$" then
		    %>
		    	<% drawSelectOffShopmargin ocontractProtoTypeDetail.FItemList(i).FDetailKey , "" %>
		    <%
		    '//���Ⱓ �������̳� ���Ⱓ �������� ���
		    elseif ocontractProtoTypeDetail.FItemList(i).FDetailKey = "$$STARTDATE$$" or ocontractProtoTypeDetail.FItemList(i).FDetailKey = "$$ENDDATE$$" then
		    %>
				<input type="text" name="<%=ocontractProtoTypeDetail.FItemList(i).FDetailKey%>" size=10>
				<a href="javascript:calendarOpen3(frmReg.<%=ocontractProtoTypeDetail.FItemList(i).FDetailKey%>,'',frmReg.<%=ocontractProtoTypeDetail.FItemList(i).FDetailKey%>.value)">
				<img src="/images/calicon.gif" width="21" border="0" align="middle"></a>
		    <%
		    else
		    %>
				<input type="text" size="40" id="DetailKey" name="<%= ocontractProtoTypeDetail.FItemList(i).FDetailKey %>" value="<%= getDefaultContractValue(ocontractProtoTypeDetail.FItemList(i).FDetailKey,opartner) %>" >
			<%
			end if
			%>
            &nbsp; <%= ocontractProtoTypeDetail.FItemList(i).FdetailDesc %>
        </td>
        <td>&nbsp;</td>
    </tr>
    <% end if %>
    <% next %>
<% else %>
<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF"></td>
    <td>-</td>
    <td>&nbsp;</td>
</tr>
<% end if %>

<tr bgcolor="#FFFFFF">
    <td bgcolor="#DDDDFF">��Ÿ��೻��</td>
    <% if ocontract.FResultCount>0 then %>
    <td colspan="2"><textarea cols="80" rows="6" name="contractEtcContetns"><%= ocontract.FOneItem.FcontractEtcContetns %></textarea></td>
    <% else %>
    <td colspan="2"><textarea cols="80" rows="6" name="contractEtcContetns"></textarea></td>
    <% end if %>

</tr>
<tr bgcolor="#FFFFFF" height="40">
    <td bgcolor="#DDDDFF">�������</td>
    <td >
        <% if ocontract.FResultCount>0 then %>
            <b><font color="<%= ocontract.FOneItem.GetContractStateColor %>"><%= ocontract.FOneItem.GetContractStateName %></font></b>
        <% else %>
            <font color="RED"><b>�� �� �� ��</b></font>
        <% end if %>
    </td>
    <td align="right">
    <% if ocontract.FResultCount>0 then %>
        <% if ocontract.FOneItem.FContractState=0 then ''������ %>
            <% if (rowcount<1) and (opartner.FOneItem.Fregdate<"2007-09-01") then %>
                <script language='javascript'>
                    alert('�����ϴ� ��ǰ�� �����ϴ�.<%= rowcount %>');
                </script>
            <% else %>
            <input type="checkbox" name="sendOpenMail" checked >���Ϲ߼�
            <input type="button" value="��ü ����" onclick="goNextState(<%= ocontract.FOneItem.FContractState %>,1,'��ü ��� ���� �Ͻðڽ��ϱ�?');" class="button">
            &nbsp;
            <% end if %>
            <input type="button" value="����" onclick="goNextState(<%= ocontract.FOneItem.FContractState %>,-1,'���� �Ͻðڽ��ϱ�?');" class="button">
        <% elseif ocontract.FOneItem.FContractState=1 then ''���� %>
            <input type="button" value="������ ����" onclick="goNextState(<%= ocontract.FOneItem.FContractState %>,0,'���������� ���� �Ͻðڽ��ϱ�?');" class="button">
            &nbsp;
            <input type="button" value="���Ϸ� ����" onclick="goNextState(<%= ocontract.FOneItem.FContractState %>,7,'��� �Ϸ� ���·� ���� �Ͻðڽ��ϱ�?');" class="button">
        <% elseif ocontract.FOneItem.FContractState=3 then ''��üȮ�� %>
            <input type="button" value="������ ����" onclick="goNextState(<%= ocontract.FOneItem.FContractState %>,0,'���������� ���� �Ͻðڽ��ϱ�?');" class="button">
            &nbsp;
            <input type="button" value="���Ϸ� ����" onclick="goNextState(<%= ocontract.FOneItem.FContractState %>,7,'��� �Ϸ� ���·� ���� �Ͻðڽ��ϱ�?');" class="button">
        <% elseif ocontract.FOneItem.FContractState=7 then ''���Ϸ� %>

        <% elseif ocontract.FOneItem.FContractState=-1 then ''���� %>
            <script >alert('������ �����Դϴ�.');</script>
        <% else %>

        <% end if %>
    <% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="3" align="center" height="30">
    <% if ocontract.FResultCount>0 then %>
        <% if ocontract.FOneItem.FContractState<>-1 then  %>
        <input type="button" value="��༭ ���� ����" onClick="SaveContract(frmReg);" class="button">
        <% end if %>
    <% else %>
        <input type="button" value="�ű԰�� ���" onClick="SaveContract(frmReg);" class="button">
    <% end if %>
    </td>
</tr>
</form>
</table>

<form name="frmReSearch" method="get" action="">
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="ContractType" value="<%= ContractType %>">
<input type="hidden" name="ContractID" value="">
</form>

<%
set opartner = Nothing
set ocontract = Nothing
set ocontractList = Nothing
set ocontractDetail = Nothing
set ocontractProtoType = Nothing
set ocontractProtoTypeDetail = Nothing
%>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->