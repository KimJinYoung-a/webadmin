<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->

<%

dim groupid, ogroup, mode, erpCustCD, assignedRow
groupid = requestCheckvar(request("groupid"),32)
mode= requestCheckvar(request("mode"),32)
erpCustCD = requestCheckvar(request("erpCustCD"),32)

dim sqlStr, arrRows, i

if (mode="sv") then
    sqlStr = "update db_partner.dbo.tbl_partner_group "
    IF (erpCustCD="") then
        sqlStr = sqlStr & " set erpCust_CD=NULL"&VbCRLF
    ELSE
        sqlStr = sqlStr & " set erpCust_CD='"&erpCustCD&"'"&VbCRLF
    ENd IF
    sqlStr = sqlStr & " where groupid='"&groupid&"'"
    dbget.Execute  sqlStr,assignedRow
    
    if assignedRow>0 then
        response.write "<script>alert('�����Ǿ����ϴ�.');location.href='?groupid="&groupid&"';</script>"
        dbget.Close()
        response.end
    end if
elseif (mode="erpusey") then
	sqlStr = "update db_partner.dbo.tbl_partner_group "
	sqlStr = sqlStr & " set erpusing=1"
	sqlStr = sqlStr & " where groupid='"&groupid&"'"
	sqlStr = sqlStr & " and erpusing=0"
    dbget.Execute  sqlStr,assignedRow

	if assignedRow>0 then
        response.write "<script>alert('�����Ǿ����ϴ�.');location.href='?groupid="&groupid&"';</script>"
        dbget.Close()
        response.end
    end if
end if

set ogroup = new CPartnerGroup
ogroup.FRectGroupid = groupid
ogroup.GetOneGroupInfo

if (ogroup.FREsultcount<1) then
    rw "�������� �ʴ� �׷��ڵ� - "&groupid
    dbget.close() : response.end
end if

dim company_no : company_no = replace(ogroup.FOneItem.Fcompany_no,"-","")


sqlStr = " select top 10 cust_cd,use_yn,del_yn,cust_nm,MOD_DT,cust_use_cd "
sqlStr = sqlStr & " from [TMSDB].db_SCM_LINK.dbo.vw_BA_CUST_sERP"
sqlStr = sqlStr & " where BIZ_NO='"&company_no&"'"
IF (company_no="") or (company_no="0000000000") then
    sqlStr = sqlStr & " and 1=0"    
end if

rsget.Open sqlStr,dbget,1
if not rsget.Eof then
    arrRows = rsget.getRows
end if
rsget.Close

dim mayErpCode , mayErpUseCode

if IsNULL(ogroup.FOneItem.FerpCust_CD) then
    mayErpCode = ogroup.FOneItem.FGroupId
else
    mayErpCode = ogroup.FOneItem.FerpCust_CD
end if
		
%>
<script language='javascript'>
function saveThis(ierpcode){
    if (ierpcode!=""){
        if (confirm('���� �Ͻðڽ��ϱ�?')){
            document.frmSave.erpCustCD.value=ierpcode;
			document.frmSave.mode.value="sv";
            document.frmSave.submit();
        }
    }
}

function delThis(){
    if (confirm('�����ڵ带 ���� �Ͻðڽ��ϱ�?')){
        document.frmSave.erpCustCD.value='';
		document.frmSave.mode.value="sv";
        document.frmSave.submit();
    }
}

function useThis(){
    if (confirm('ERP ���� ��������� ���� �Ͻðڽ��ϱ�?')){
        document.frmSave.mode.value="erpusey";
        document.frmSave.submit();
    }
}


function popErpBizInfo(hidCcd){
    var popwin = window.open('/admin/linkedERP/cust/regCust.asp?rO=on&hidCcd='+hidCcd,'popErpBizInfo','width=900,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}
</script>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">

<tr height="25">
	<td width="90" bgcolor="<%= adminColor("tabletop") %>">��ü�ڵ�</td>
	<td bgcolor="#FFFFFF" >
		<%= ogroup.FOneItem.FGroupId %>
	</td>
</tr>
<tr height="25">
	<td width="90" bgcolor="<%= adminColor("tabletop") %>">ERP�����ڵ�</td>
	<td bgcolor="#FFFFFF" >
		<%= mayErpCode %>
		<% if (ogroup.FOneItem.FerpUsing<>1) then %>
	        ��������
			<input type="button" value="���������μ���" onClick="useThis()">
	    <% end if %>
	    
	    <% if (mayErpCode<>ogroup.FOneItem.FGroupId) then %>
	    <input type="button" value="����" onClick="delThis()">
	    <% end if %>
	</td>
</tr>
<tr height="25">
	<td width="90" bgcolor="<%= adminColor("tabletop") %>">ȸ���</td>
	<td bgcolor="#FFFFFF" >
		<%= ogroup.FOneItem.FCompany_name %>
	</td>
</tr>
<tr height="25">
	<td width="90" bgcolor="<%= adminColor("tabletop") %>">����ڹ�ȣ</td>
	<td bgcolor="#FFFFFF" >
		<%= ogroup.FOneItem.Fcompany_no %>
	</td>
</tr>
<tr height="25">
	<td width="90" bgcolor="<%= adminColor("tabletop") %>">ERP�ڵ忬��</td>
	<td bgcolor="#FFFFFF" >
	    <table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if isArray(arrrows) then %>
        <tr bgcolor="<%= adminColor("sky") %>" align="center" >
            <td>ERP�ڵ�</td>
		    <td>Ȱ������</td>
		    <td>��������</td>
		    <td>ȸ���</td>
		    <td>����</td>
        </tr>
<% For i = 0 To UBound(arrRows,2) %>
        <tr bgcolor="#FFFFFF" align="center" >
		    <td><a href="javascript:popErpBizInfo('<%= arrRows(0,i) %>')"><%= arrRows(5,i) %><br>(<%= arrRows(0,i) %>)</a></td>
		    <td><%= arrRows(1,i) %></td>
		    <td><%= arrRows(2,i) %></td>
		    <td><%= arrRows(3,i) %></td>
		    <td>
		    <% if (arrRows(1,i)="Y") and (arrRows(2,i)="N") then %>
		        <% if (mayErpCode<>arrRows(0,i)) then %>
		        <input type="button" value="����" onClick="saveThis('<%=arrRows(0,i)%>')">
		        <% else %>
		        ��
		        <% end if %>
		    <% end if %>
		    </td>
	    </tr>
<% next %>
<% else %>
        <tr bgcolor="<%= adminColor("sky") %>" align="center" >
            <td align="center">��ϵ� �ش� ����� ��ȣ�� �����ϴ�.</td>
        </tr>
<% end if %>
        </table>
</td>
</tr>
</table>


<form name="frmSave" method="post" action="">
<input type="hidden" name="groupid" value="<%= ogroup.FOneItem.FGroupId %>">
<input type="hidden" name="mode" value="sv">
<input type="hidden" name="erpCustCD" value="">
</form>
<%
set ogroup = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->