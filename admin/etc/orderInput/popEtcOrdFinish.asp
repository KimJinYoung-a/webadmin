<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdminNoCache.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<%
Dim ooseq : ooseq = requestCheckvar(request("ooseq"),10)
Dim actTp : actTp = requestCheckvar(request("actTp"),10)
Dim finDiv : finDiv = requestCheckvar(request("finDiv"),10)
Dim csIdx : csIdx = requestCheckvar(request("csIdx"),10)
Dim chOutMallOrderSerial : chOutMallOrderSerial = requestCheckvar(request("chOutMallOrderSerial"),30)
Dim orgOutMallOrderSerial : orgOutMallOrderSerial = requestCheckvar(request("orgOutMallOrderSerial"),30)
Dim sellSite : sellSite = requestCheckvar(request("sellSite"),32)
Dim mode : mode = requestCheckvar(request("mode"),10)

Dim sqlStr,AssignedRow

if (finDiv<>"") and (mode="actEtc") then
    if (csIdx="") and ((finDiv="C") or (finDiv="R")) then
        response.write "<script>alert('CS ó����ȣ �ʼ� ����');history.go(-1);</script>"
        dbget.close() : response.end
    end if
    
    if (finDiv="P") then    ''�ߺ��ּ���ó��
        sqlStr = "Update db_temp.dbo.tbl_xSite_tmpOrder " & vbCRLF
        sqlStr = sqlStr & " set ref_outmallorderserial='"&orgOutMallOrderSerial&"'" & vbCRLF
        sqlStr = sqlStr & " ,OutMallOrderSerial='"&chOutMallOrderSerial&"'" & vbCRLF
        sqlStr = sqlStr & " ,etcFinUser='"&session("ssBctID")&"'" & vbCRLF
        sqlStr = sqlStr & " where sellSite='"&sellSite&"'"  & vbCRLF
        sqlStr = sqlStr & " and OutMallOrderSeq="&ooseq&""  & vbCRLF
        sqlStr = sqlStr & " and orderserial is NULL" 
        dbget.Execute sqlStr,AssignedRow
        
    elseif (finDiv="D") then    ''��ҿϷ� (��Ȥ CS ���� ���� ��ҷ� �ö���� ��� ����)
    
    elseif (finDiv="C") then    ''�������ǰ(�ɼ�)����

    elseif (finDiv="R") then    ''(��)��ȯ/ȸ���Ϸ�
        
    end if
    
    if (AssignedRow=1) then
        response.write "<script>alert('�����Ǿ����ϴ�.');opener.location.reload();window.close();</script>"
        dbget.close() : response.end
    else
        response.write "<script>alert('ó���� ����.');</script>"
    end if
end if

Dim otmpOneOrder
set otmpOneOrder = new CxSiteTempOrder
otmpOneOrder.FRectOutMallOrderSeq   = ooseq
otmpOneOrder.getOneTmpOrder


Dim OutMallOrderSerial
IF Not IsNULL(otmpOneOrder.FOneItem.FRef_OutMallOrderSerial) then
    OutMallOrderSerial = otmpOneOrder.FOneItem.FRef_OutMallOrderSerial
else
    OutMallOrderSerial = otmpOneOrder.FOneItem.FOutMallOrderSerial
end if

Dim otmpOrder, i
set otmpOrder = new CxSiteTempOrder
otmpOrder.FPageSize = 100
otmpOrder.FCurrPage = 1
otmpOrder.FRectSellSite   = otmpOneOrder.FOneItem.FSellSite
otmpOrder.FRectoutmallorderserial=OutMallOrderSerial
otmpOrder.getOnlineTmpOrderList(true)

Dim OutOrderserialArr, buf, mxBuf
IF (otmpOrder.FResultCount>0) then
    for i=0 to otmpOrder.FResultCount-1 
        buf = replace(otmpOrder.FItemList(i).FOutMallOrderSerial,OutMallOrderSerial,"")
        buf = replace(buf,"_","")
        
        if (buf<>"") then
            mxBuf=buf
        end if
    next
    
    if (mxBuf<>"") then
        mxBuf = CStr(CLNG(mxBuf)+1)
    else
        mxBuf = "1"
    end if
ENd If
%>
<script language='javascript'>
function finThis(){
    var frm=document.frmAct;
    if (frm.finDiv.value.length<1){
        alert('ó�������� �����ϼ���.');
        frm.finDiv.focus();
        return;   
    }
    
    if ((frm.finDiv.value=="R")&&(frm.csIdx.value.length<1)){
        alert('CS ó�� ��ȣ�� �Է��ϼ���.');
        frm.csIdx.focus();
        return; 
    }
    
    if ((frm.finDiv.value=="P")&&(frm.chOutMallOrderSerial.value.length<1)){
        alert('�ű� ���� �ֹ���ȣ�� �Է��ϼ���.');
        frm.chOutMallOrderSerial.focus();
        return; 
    }

    if (confirm('�Ϸ�ó�� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}

function chgDiv(comp){
    var pval = comp.value;
    
    if ((pval=="R")||(pval=="C")){
        document.getElementById("selDiv_R").style.display="block";
        document.getElementById("selDiv_P").style.display="none";
    }else if (pval=="P"){
        document.getElementById("selDiv_R").style.display="none";
        document.getElementById("selDiv_P").style.display="block";
    }else{
        document.getElementById("selDiv_R").style.display="none";
        document.getElementById("selDiv_P").style.display="none";
    }
    
       
}
</script>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frmAct" method="post">
<input type="hidden" name="mode" value="actEtc">
<input type="hidden" name="sellSite" value="<%= otmpOneOrder.FOneItem.FsellSite %>">
<input type="hidden" name="ooseq" value="<%= otmpOneOrder.FOneItem.FOutMallOrderSeq %>">
<input type="hidden" name="orgOutMallOrderSerial" value="<%= otmpOneOrder.FOneItem.FOutMallOrderSerial %>">
<tr align="center" height="25">
    <td bgcolor="#E8E8FF">������</td>
    <td bgcolor="#FFFFFF"><%= otmpOneOrder.FOneItem.FSellSite %></td>
</tr>
<tr align="center" height="25">
    <td bgcolor="#E8E8FF">�����ֹ���ȣ</td>
    <td bgcolor="#FFFFFF"><%= otmpOneOrder.FOneItem.FOutMallOrderSerial %></td>
</tr>
<tr align="center" height="25">
    <td bgcolor="#E8E8FF">�����ֹ�(��)��ȣ</td>
    <td bgcolor="#FFFFFF"><%= otmpOneOrder.FOneItem.FOrgDetailKey %></td>
</tr>
<tr align="center" height="25">
    <td bgcolor="#E8E8FF">�����ֹ���ǰ</td>
    <td bgcolor="#FFFFFF"><%= otmpOneOrder.FOneItem.ForderItemName %></td>
</tr>
<tr align="center" height="25">
    <td bgcolor="#E8E8FF">�����ֹ��ɼ�</td>
    <td bgcolor="#FFFFFF"><%= otmpOneOrder.FOneItem.ForderItemOptionName %></td>
</tr>
<tr align="center" height="25">
    <td bgcolor="#E8E8FF">ó������</td>
    <td bgcolor="#FFFFFF">
    <select name="finDiv" onChange="chgDiv(this)">
    <option value="">����
    <option value="R">(��)��ȯ/ȸ���Ϸ�
    <option value="C">�������ǰ(�ɼ�)����
    <option value="D">��ҿϷ�
    <option value="P">�ߺ��ּ���ó��(�ű��ֹ���ȣ����)
    </select>
    
    <div id="selDiv_R" name="selDiv_R" style="display:none">
        CSó�� ��ȣ : <input type="text" name="csIdx" value="" size="5" maxlength="9">
    </div>
    
    <div id="selDiv_P" name="selDiv_R" style="display:none">
        �ű� �ֹ���ȣ : <input type="text" name="chOutMallOrderSerial" value="<%=OutMallOrderSerial%>_<%=mxBuf%>">
    </div>
    </td>
</tr>
<tr align="center" height="25" bgcolor="#FFFFFF">
    <td colspan="2" align="center">
    <input type="button" value="�Ϸ�ó��" onClick="finThis();">
    </td>
</tr>
</form>
</table>
<%
set otmpOrder = Nothing
set otmpOneOrder=Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->