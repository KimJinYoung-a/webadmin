<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� �ֹ��Է� ��ġ
' Hieditor : 2019.01.24 eastone
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
public function getXSiteTmpOrderBatchTargetList(iSellSite,iPageSize)
    Dim sqlStr
    sqlStr = "db_temp.dbo.usp_TEN_xSiteTmpOrderBatchInputTarget '"&iSellSite&"',"&iPageSize
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if not rsget.Eof then
        getXSiteTmpOrderBatchTargetList = rsget.GetRows()
    end if
    rsget.close

end function

Dim sellsite : sellsite = requestCheckvar(request("sellsite"),32)
Dim ignore3Err : ignore3Err = requestCheckvar(request("ignore3Err"),32)
Dim research : research = requestCheckvar(request("research"),32)
Dim iMaxpageSize : iMaxpageSize=300
dim ArrRows : ArrRows = getXSiteTmpOrderBatchTargetList(sellsite,iMaxpageSize)

if (ignore3Err="") and (research="") then ignore3Err="on"
Dim isIgnore3Err : isIgnore3Err = ignore3Err="on"

dim i, ttlCnt : ttlCnt = 0
if IsArray(ArrRows) then
    ttlCnt = UBound(ArrRows,2)+1
end if

Dim rowErr : rowErr=0
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">
<script language='javascript'>
var batchstarted = false;
var nextid = 0;
function xlOnlineOrderBatchInput(comp){
    if (batchstarted) return;

    comp.disabled = true;
    comp.style="background-color: #cccccc;color: #888888;"
    batchstarted = true;

    addNotiLog("start");

    fnNextOrderInputProc();
}

function addNotiLog(ilog){
    document.getElementById("disp1").innerHTML += ilog+"<br>";
}

function addResultLog(orderSeq,ilog){
    document.getElementById("oseq_"+orderSeq).innerHTML = ilog;
}

function fnNextOrderInputProc(){
    var frm = document.frmBatchArr;
    if (!frm.ix){
        addNotiLog('������ �����ϴ�.')
        return;
    }

    var ix = -1;
    if (!frm.ix.length){
        ix = frm.ix.value*1;
    }else{
        if (frm.ix.length>nextid){
            ix = frm.ix[nextid].value*1;
        }
    }

    if (nextid>ix){
        addNotiLog('finished !');
        setTimeout(function(){ alert('finished'); }, 100);  
        return;
    }

    if (nextid><%=iMaxpageSize%>){
       ddNotiLog('oops !');
       return;     
    }

    setTimeout(function(){
        oneOrderInput(ix);
    }, 500);  

    
}

function oneOrderInput(iidx){
    nextid = iidx+1;
    var arrfrm = document.frmBatchArr;

    if (!arrfrm.ix.length){
        if (arrfrm.rowErrNo.value*1>0){
            addResultLog(arrfrm.minOutMallOrderSeq.value,"skip");
            fnNextOrderInputProc();    
        }else{
            document.frmBatch.oseq.value = arrfrm.minOutMallOrderSeq.value;
            document.frmBatch.cksel.value = arrfrm.OutMallOrderSerial.value;
            addNotiLog(document.frmBatch.cksel.value);
            document.frmBatch.submit();
        }
    }else{
        if (arrfrm.rowErrNo[iidx].value*1>0){
            addResultLog(arrfrm.minOutMallOrderSeq[iidx].value,"skip");
            fnNextOrderInputProc();
        }else{
            document.frmBatch.oseq.value = arrfrm.minOutMallOrderSeq[iidx].value;
            document.frmBatch.cksel.value = arrfrm.OutMallOrderSerial[iidx].value;
            addNotiLog(document.frmBatch.cksel.value);
            document.frmBatch.submit();
        }
    }
    

    
}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
	    * ���θ� ���� :
	    <% call drawSelectBoxXSiteOrderInputPartner("sellsite", sellsite) %>
		&nbsp;&nbsp;&nbsp;
        <input type="checkbox" name="ignore3Err" <%=CHKIIF(isIgnore3Err,"checked","")%>>ǰ��,����,1+1����
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>

</form>
</table>
<!-- �˻� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
    <td>
    <div id="disp1" style="overflow: auto; width:100%; height: 50px;"></div>
    </td>
    <td width="300">
    <iframe name="xLink3" id="xLink3" frameborder="0" width="100%" height="50"></iframe>
    </td>
</tr>
</table>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
    <td align="right">
        <input type="button" class="button" value="�ֹ� �ϰ����" onClick="xlOnlineOrderBatchInput(this);">
    </td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmBatchArr" >
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
		�˻���� : <b><%= ttlCnt %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="60"></td>
    <td width="90">�Ǹż��θ�</td>
    <td width="90">�����Ǹ���</td>
    <td width="120">�����ֹ���ȣ</td>
    <td width="60">�Ǽ�</td>

    <td width="60">�ּҴ���</td>
    <td width="60">���۹���</td>
    <td width="60">���۹���(11st)</td>
    <td width="60">�����ȣ</td>
    <td width="60">�ǸŰ�<1</td>
    <td width="60">�ɼ�FF</td>
    <td width="60">�ڵ����</td>
    <td width="60">1+1</td>
    <td width="60">ǰ��</td>
    <td width="60">����</td>
    <td >���</td>
 </tr>
 <% if isArray(ArrRows) then %>
 <% For i =0 To UBound(ArrRows,2) %>
 <%
 if (isIgnore3Err) then
    rowErr = ArrRows(5,i)+ArrRows(6,i)+ArrRows(7,i)+ArrRows(8,i)+ArrRows(9,i)+ArrRows(10,i)+ArrRows(11,i)
 else
    rowErr = ArrRows(5,i)+ArrRows(6,i)+ArrRows(7,i)+ArrRows(8,i)+ArrRows(9,i)+ArrRows(10,i)+ArrRows(11,i)+ArrRows(12,i)+ArrRows(13,i)+ArrRows(14,i)
 end if

 
 %>
 <input type="hidden" name="ix" value="<%=i%>">
 <input type="hidden" name="minOutMallOrderSeq" value="<%=ArrRows(0,i)%>">
 <input type="hidden" name="OutMallOrderSerial" value="<%=ArrRows(3,i)%>">
 <input type="hidden" name="rowErrNo" value="<%=rowErr%>">
 <tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
    <td><%=ArrRows(0,i)%></td>
    <td><%=ArrRows(1,i)%></td>
    <td><%=ArrRows(2,i)%></td>
    <td><%=ArrRows(3,i)%></td>
    <td><%=ArrRows(4,i)%></td>

    <td><%=ArrRows(5,i)%></td>
    <td><%=ArrRows(6,i)%></td>
    <td><%=ArrRows(7,i)%></td>
    <td><%=ArrRows(8,i)%></td>
    <td><%=ArrRows(9,i)%></td>
    <td><%=ArrRows(10,i)%></td>
    <td><%=ArrRows(11,i)%></td>
    <td <%=CHKIIF(isIgnore3Err,"bgcolor='#EEEEEE'","") %>><%=CHKIIF(isIgnore3Err,"<font color='#AAAAAA'>","") %><%=ArrRows(12,i)%><%=CHKIIF(isIgnore3Err,"</font>","") %></td>
    <td <%=CHKIIF(isIgnore3Err,"bgcolor='#EEEEEE'","") %>><%=CHKIIF(isIgnore3Err,"<font color='#AAAAAA'>","") %><%=ArrRows(13,i)%><%=CHKIIF(isIgnore3Err,"</font>","") %></td>
    <td <%=CHKIIF(isIgnore3Err,"bgcolor='#EEEEEE'","") %>><%=CHKIIF(isIgnore3Err,"<font color='#AAAAAA'>","") %><%=ArrRows(14,i)%><%=CHKIIF(isIgnore3Err,"</font>","") %></td>
    <td>
    <div id="oseq_<%=ArrRows(0,i)%>"></div>
    </td>
</tr>
<% Next %>
<% elseif (sellsite="") then%>
<tr align="center" bgcolor="FFFFFF" >
    <td colspan="16" >�˻������ �����ϴ�. Mall �� �����ϼ���.</td>
</tr>
<% else %>
<tr align="center" bgcolor="FFFFFF" >
    <td colspan="16">�˻������ �����ϴ�.</td>
</tr>
<% end if %>
</table>
</form>

<form name="frmBatch" method="post" action="OrderInput_Process.asp" target="xLink3">
<input type="hidden" name="mode" value="add">
<input type="hidden" name="oseq" value="">
<input type="hidden" name="cksel" value="">
<input type="hidden" name="xtype" value="batch">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
