<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �귣�� ��� ����
' Hieditor : 2009.04.07 ������ ����
'			 2010.05.26 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/CheckLoginReDirect.asp" -->
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/partners/contractcls.asp"-->

<%
dim ContractID, makerid , i , sqlStr , ocontract, ocontractList ,opartner , onoffgubun
	ContractID  = requestCheckVar(request("ContractID"),100)
	makerid     = session("ssBctID")

set ocontractList = new CPartnerContract
ocontractList.FRectMakerid = makerid
if makerid<>"" then
    ocontractList.GetMakerValidContractList
end if

if (ContractID="") or (ContractID="0") then
    if (ocontractList.FResultCount>0) then
        ContractID = ocontractList.FItemList(0).FContractID
    end if
end if


set ocontract = new CPartnerContract
ocontract.FRectContractID = ContractID
ocontract.FRectMakerid = makerid
if ContractID<>"" then
    ocontract.getOneContract
end if

set opartner = new CPartnerUser
opartner.FRectDesignerID = makerid

if (makerid<>"") then
    opartner.GetOnePartnerNUser
end if

if ocontract.FResultCount>0 then
	if ocontract.FOneItem.FContractType <> "" then
		sqlStr = "select contractContents, contractName ,onoffgubun" +vbcrlf
		sqlStr = sqlStr & " from db_partner.dbo.tbl_partner_contractType" +vbcrlf
		sqlStr = sqlStr & " where contractType=" & ocontract.FOneItem.FContractType
		
		'response.write sqlStr &"<br>"
		rsget.Open sqlStr,dbget,1
		if Not rsget.Eof then
		    onoffgubun = rsget("onoffgubun")
		end if
		rsget.Close
	end if
end if
%>

<script language='javascript'>

//window.resizeTo(600,600);

function changeContract(comp){
    document.frmResearch.ContractID.value = comp.value;
    document.frmResearch.submit();
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td align="right">
	        <!-- select Box -->
	        <select class="select" name="ContractID" onChange="changeContract(this);">
		        <% for i=0 to ocontractList.FResultCount-1 %>
		        <option value="<%= ocontractList.FItemList(i).FContractID %>" <% if CStr(ocontractList.FItemList(i).FContractID)=ContractID  then response.write "selected" %> >[<%= ocontractList.FItemList(i).FContractNo %>] <%= ocontractList.FItemList(i).FContractName %>
		        <% next %>
	        </select>
	    </td>
	</tr>
	<% if ocontract.FResultCount>0 then %>
	<tr bgcolor="#FFFFFF">
	    <td>
	        <table width="100%" border="0" cellspacing="1" cellpadding="1" class="a" >
	        <tr>
	            <td>
	            �ȳ��ϼ���<br>
	            (��)�ٹ����ٰ� ���� �ο����� ������ �Ǿ� �ݰ����ϴ�.<br>
	            <br>
	            �Ʒ��� ���� ����� ����ǿ��� <br>
	            ��༭ ��������� �Ĳ��� �о��ֽ� �� <br>
	            ������ ���߾� ��༭�� �������� �߼��� �ֽø� �����ϰڽ��ϴ�.<br>
	            </td>
	        </tr>
	        <tr>
	            <td><br>��༭ �� :  <%= ocontract.FOneItem.FContractName %>  </td>
	        <tr>
	        <tr>
	            <td>��༭ ��ȣ : <%= ocontract.FOneItem.FcontractNo %>  </td>
	        <tr>
	        <% if ocontract.FOneItem.FContractState>=7 then %>
	        <tr>
	            <td>���� : ���Ϸ� (�Ϸ��� : <%= ocontract.FOneItem.FFinishDate %>) </td>
	        <tr>
	        <% else %>
	        <tr>
	            <td>
	            <% if onoffgubun = "ON" then %>	
		            �� ��༭ ������ : �¶��� ���� �� �귣��  <br>
		             - �������ο��� �����ϴ� �귣��� ��󿡼� ���ܵ˴ϴ�. <br>
		             (�������� �����ÿ��� �������� ����ڰ� ���������� �����帳�ϴ�.)
		        <% else %>
		            �� ��༭ ������ : �������� ���� �� �귣��  <br>		        
		        <% end if %>     
	            <br>
	            <br>
	            
	            �� ��༭ �ٿ��� <br>
	            - �Ʒ� [��༭ �ٿ�ε�] Ŭ���Ͽ� �ٿ���� �� ����Ȯ�� �� ��������� �������ּ���!!<br>
	            - ��༭ ���� �ϴ� ����� [���flow �ٿ�ε�] �ٿ������ �� �� ������ ���ֽñ� �ٶ��ϴ�.<br>
	
	            <a href="/designer/company/downLoadContract.asp?ContractID=<%= ContractID %>" target="iTargetFrm"><b><font color="blue">[��༭ �ٿ�ε�]</font></b></a>
	            <br>
	            <a href="/designer/company/contractflow.ppt" target="_blank"><b><font color="blue">[���flow �ٿ�ε�]</font></b></a>
	            
	            <br><br>
	            �� �ʼ� Ȯ�λ��� (�ݵ�� �ι� ���� Ȯ�����ּ���)  <br>
	            ������, ������ �� �ΰ����� �´� �� �� Ȯ�����ּž� �մϴ�. 
	            <br><br>
	            �� ��ü�������� �ʼ� ������� (��!! ���� �����ϼž� �� �κ�)<br>
	            - ǥ��(ù��)�� ������� ���� : ���¾�ü�� ��ǥ�̻� �Ǵ� ����� ������ �����Ͻô� ����� ����<br>
	            <% if ocontract.FOneItem.FContractType=5 then %>
	            - ���å���� ����<br>
	            <% end if %>
	            - ������ ���� '��'�� ��ǥ�̻� �ֹε�Ϲ�ȣ �� �ּ� ���� : ����ڵ������ ��ǥ�� �ֹι�ȣ �� �ּҿ��� �մϴ�.<br>
	            - ���λ������ ��� ���������� ��ǥ�̻� �ֹι�ȣ �� �ּҸ� �����ϼŵ� �Ǹ�, '��'����� ���� '��' ����� ������ ������ �˴ϴ�. 
	
	            <br><br>
	            �� �������� :  <br>
	            �� ��༭ �ٿ�ε� <br>
	            �� ���¾�ü���� ��༭ Ȯ�� �� ���� / 2�� ����߼�  <br>
	            �� �ٹ����ٿ��� ��༭ ���� ����Ȯ�� <br>
	            �� �ٹ����ٿ��� ���¾�ü�� ��༭ 1�� �߼� / ���Ϸ� 
	            <br>
	            <br>
	            
	            �� ��༭ �����ô� �� <br>
	            <% if onoffgubun = "ON" then %>
		            �ּ� : ����� ���α� ������ 1-45���� �������� 5�� �ٹ����� <br>
		            ����� : <%= ocontract.FoneItem.Fusername %> <br>
		            tel : <%= ocontract.FoneItem.Finterphoneno %> (���� <%= ocontract.FoneItem.Fextension %>) / ���� : <%= ocontract.FoneItem.Fdirect070 %><br>
		            fax : 02-2179-9244 <br>
	            <% else %>
		            �ּ� : ����� ���α� ������ 1-74 ������ġ Ȧ�������� 6�� �ٹ����� �������� �繫��<br>
		            ����� : <%= ocontract.FoneItem.Fusername %> <br>
		            tel : <%= ocontract.FoneItem.Finterphoneno %> (���� <%= ocontract.FoneItem.Fextension %>) / ���� : <%= ocontract.FoneItem.Fdirect070 %><br>
		            fax : 02-2179-9058 <br>
	            <% end if %>
	            
	            <br>
	            
	            
	            <% if onoffgubun = "ON" then %>
	            <!-- ���ȭ �ʿ�..
	            �� �����к긯 / ������ : ������ �븮 (���� 153 ) <br>
                �� ��� �� �м� : ������ �븮 (���� 154 )<br>
                �� �����ι��� /���ǽ����� :������ ���� (���� 152)<br>
                �� ī�޶�, book, baby : �ָ����Ҹ� ���� (����159)<br>
                �� Ű��Ʈ ��� : ������ (���� 157)<br>
                �� �ֹ��� ���� ���̽�Ʈ : ������ (���� 156)<br>
                �� �м���ȭ ��� : ������ (���� 155)<br>
                -->
	            <!-- �����繫:������ ���� (���� 152) / �����ֹ�:������ �븮 (���� 153 ) <br>
	            �м����:������ �븮 (���� 154 ) / Ű��Ʈ: ������ (���� 157) <br> -->
				<% end if %>	            
	            
	            <br><br>
	            �� ����߼۽� �Բ� �����ž� �� ���� <br>
	            - ���ε� ��༭ 2�� <br>
	            - �������� �纻 <br>
	            - ����� ����� �纻 <br>
	            - �ΰ����� ���� (��༭�� ������ ����) 
	            <br>
	            <br>
	            �� �� Ÿ <br>
	            �ٹ����� ���� �����ϴ� �귣�� ���̵� 2�� �̻��� ��� <br>
	            ��༭�� �� �귣�� ���̵𸶴� �ۼ��� ���ּž� �ϸ�, <br>
	            ���ü���(����ڵ����,�ΰ�����,��������)�� 1�θ� �ּŵ� �˴ϴ�. 
	            <br>
	            <br>
	            �� ��༭ ���� ���� �ñ��� ���� �� ���MD���� ���� �Ͻñ� �ٶ��ϴ�.
	            
	            <!--
	            �� �߿亯�����  <br>
	            - ����ڵ���� ���� ��� -> �귣�庰 ���  <br>
	            (���� ��� �� ����ڿ��� 3���� �귣�带 ��� ��� �� �귣�帶�� ����Ͽ� �� 3���� ��༭�� �������ּž� �մϴ�.) <br>
	            - �������� ������ �߰� �� ��ǰ��� �Խù��� ���� ��ȭ
	            <br><br>
	            -->
	            
	            <% if onoffgubun = "ON" then %>
	            <!--
	            �� ��༭ ���� ���� �ñ��� ���� �Ʒ��� �����ּ��� (�� ������, ������ ���� ��� MD���� ����)<br>
	            
	            �����ٹ����� �����, ������������Ʈ <br>
	            �븮 �̼���<br>
	            TEL : 02-554-2033(#143)<br>
	            FAX : 02-2179-9245<br>
	            E.MAIL : snowsilver@10x10.co.kr<br>
	            WEB : www.10x10.co.kr<br>
	            -->
				<% end if %>
	            </td>
	        </tr>
	        <% end if %>
	        </table>
	    </td>
	</tr>
	<% else %>
	<tr bgcolor="#FFFFFF">
	    <td height="50" align="center">
	    [<%= makerid %> : ���õ� ��༭�� �����ϴ�. ���� ��༭�� ������ �ּ���.]
	    </td>
	</tr>
	<% end if %>
</table>

<form name="frmResearch">
<input type="hidden" name="ContractID" value="<%= ContractID %>">
</form>
<iframe name="iTargetFrm" id="iTargetFrm" src="" width="1" height="1" frameborder=0 scrolling=no marginheight=0 marginwidth=0 align=center></iframe>
<% 
set ocontract = Nothing
set ocontractList = Nothing
set opartner = Nothing
%>
<!-- #include virtual="/designer/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->