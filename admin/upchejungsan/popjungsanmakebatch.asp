<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : OFF ���� �ϰ��ۼ�
' Hieditor : 2020.01.02 eastone
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
public function getJungsanBatchTargetList(ijyyyymm,itargetGbn,ijgubun,iPageSize)
    Dim sqlStr
    '' @yyyymm varchar(7) ,@targetGbn varchar(2) ,@jgubun varchar(2) ,@DLVGbn int ,@vatyn varchar(1)
    sqlStr = "db_jungsan.dbo.usp_TEN_JungsanBatch_getTargetList '"&ijyyyymm&"','"&itargetGbn&"','"&ijgubun&"',"&iPageSize
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if not rsget.Eof then
        getJungsanBatchTargetList = rsget.GetRows()
    end if
    rsget.close

end function

public function isJungsanBatchTargetMaded(ijyyyymm,itargetGbn)
    Dim sqlStr
    '' @yyyymm varchar(7) ,@targetGbn varchar(2) ,@jgubun varchar(2) ,@DLVGbn int ,@vatyn varchar(1)
    sqlStr = "db_jungsan.[dbo].[usp_TEN_JungsanBatch_getTargetMadeCNT] '"&ijyyyymm&"','"&itargetGbn&"'"
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if not rsget.Eof then
        isJungsanBatchTargetMaded = rsget("cnt")>0
    end if
    rsget.close
end function



Dim yyyy1 : yyyy1 = requestCheckvar(request("yyyy1"),10)
Dim mm1 : mm1 = requestCheckvar(request("mm1"),10)
Dim targetGbn : targetGbn = requestCheckvar(request("targetGbn"),10)
Dim jgubun : jgubun = requestCheckvar(request("jgubun"),10)
Dim DLVGbn : DLVGbn = requestCheckvar(request("DLVGbn"),10)
Dim vatyn : vatyn = requestCheckvar(request("vatyn"),10)
Dim research : research = requestCheckvar(request("research"),32)
Dim nloop : nloop = requestCheckvar(request("nloop"),10)
Dim iMaxpageSize : iMaxpageSize=100

if (nloop="") then nloop=1

dim ArrRows : ArrRows = getJungsanBatchTargetList(yyyy1+"-"+mm1,targetGbn,jgubun,iMaxpageSize)


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
function JBatchMake(comp){
    if (batchstarted) return;

    comp.disabled = true;
    comp.style="background-color: #cccccc;color: #888888;"
    batchstarted = true;

    addNotiLog("start");

    fnNextJungsanInputProc();
}

function addNotiLog(ilog){
    document.getElementById("disp1").innerHTML += ilog+"<br>";
}

function addResultLog(orderSeq,ilog){
    document.getElementById("oseq_"+orderSeq).innerHTML = ilog;
}

function fnNextJungsanInputProc(){
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
        if (document.getElementById("nloop").value*1>1){
            setTimeout(function(){ location.href="?targetGbn=<%=targetGbn%>&yyyy1=<%=yyyy1%>&mm1=<%=mm1%>&jgubun=<%=jgubun%>&nloop="+(document.getElementById("nloop").value*1-1)+"&forcerestrt=on"; }, 100); 
        }else{
            setTimeout(function(){ alert('finished'); }, 100); 
        }
        
        return;
    }

    if (nextid><%=iMaxpageSize%>){
       ddNotiLog('oops !');
       return;     
    }

    setTimeout(function(){
        oneJInput(ix);
    }, 200);  

    
}

function oneJInput(iidx){
    nextid = iidx+1;
    var arrfrm = document.frmBatchArr;

    if (!arrfrm.ix.length){
        if (arrfrm.rowErrNo.value*1>0){
            addResultLog(arrfrm.ix.value,"skip");
            fnNextJungsanInputProc();    
        }else{
            document.frmBatch.oseq.value = arrfrm.ix.value;
            document.frmBatch.jyyyymm.value = arrfrm.jyyyymm.value;
            document.frmBatch.targetGbn.value = arrfrm.targetGbn.value;
            document.frmBatch.jgubun.value = arrfrm.jgubun.value;
            document.frmBatch.makerid.value = arrfrm.makerid.value;
            document.frmBatch.DLVGbn.value = arrfrm.DLVGbn.value;
            document.frmBatch.vatyn.value = arrfrm.vatyn.value;

            addNotiLog(document.frmBatch.oseq.value);

            document.frmBatch.submit();
        }
    }else{
        if (arrfrm.rowErrNo[iidx].value*1>0){
            addResultLog(arrfrm.ix[iidx].value,"skip");
            fnNextJungsanInputProc();
        }else{
            document.frmBatch.oseq.value = arrfrm.ix[iidx].value;
            document.frmBatch.jyyyymm.value = arrfrm.jyyyymm[iidx].value;
            document.frmBatch.targetGbn.value = arrfrm.targetGbn[iidx].value;
            document.frmBatch.jgubun.value = arrfrm.jgubun[iidx].value;
            document.frmBatch.makerid.value = arrfrm.makerid[iidx].value;
            document.frmBatch.DLVGbn.value = arrfrm.DLVGbn[iidx].value;
            document.frmBatch.vatyn.value = arrfrm.vatyn[iidx].value;

            addNotiLog(document.frmBatch.oseq.value);
            document.frmBatch.submit();
        }
    }
    

    
}

function makeJtarget(){
    if (confirm("���� ����� �ۼ��Ͻðڽ��ϱ�?")){
        document.frmBatchTarget.submit();
    }
    
}

$(document).ready(function() { 
    <% if (request("forcerestrt")<>"") then %>
    JBatchMake(document.getElementById("btnbatch"));
    <% end if %>
});

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
        * ����� :
        <% DrawYMBox yyyy1,mm1 %>
        &nbsp;&nbsp;

	    * ON/OFF ���� :
        <select name="targetGbn" >
		<option value="">��ü
		<option value="ON" <%= CHKIIF(targetGbn="ON","selected","") %> >ON
		<option value="OF" <%= CHKIIF(targetGbn="OF","selected","") %> >OF
		</select>
        &nbsp;&nbsp;

	    * ���� ���� :
	    <% drawSelectBoxJGubun "jgubun",jgubun %>
        &nbsp;&nbsp;
<% if (FALSE) then %>
        * ��ۺ� ���� :
        <select name="DLVGbn" >
		<option value="">����
		<option value="0" <%= CHKIIF(DLVGbn="0","selected","") %> >�Ϲ�
		<option value="1" <%= CHKIIF(DLVGbn="1","selected","") %> >��ۺ�
		</select>
        &nbsp;&nbsp;

        * ���� ���� :
        <select name="vatyn" >
		<option value="">����
		<option value="Y" <%= CHKIIF(vatyn="Y","selected","") %> >����
		<option value="N" <%= CHKIIF(vatyn="N","selected","") %> >�鼼
		</select>
<% end if %>
        &nbsp;&nbsp;

		&nbsp;&nbsp;&nbsp;
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
        <input type="text" name="nloop" id="nloop" value="<%=nloop%>" size="3" maxlength="3">���ݺ� 
        &nbsp;
        <input type="button" class="button" id="btnbatch" value="�����ϰ��ۼ�" onClick="JBatchMake(this);">
    </td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmBatchArr" >
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= ttlCnt %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="80">�����</td>
    <td width="90">ON/OFF</td>
    <td width="90">���걸��</td>
    <td width="120">�귣�� ID</td>
    <td width="60">��ۺ񿩺�</td>
    <td width="60">��������</td>

    <td width="90">����Ǽ�</td>
    <td width="100">����ݾ�</td>
    <td width="100">���������</td>
    <td width="90">�ۼ��Ǽ�</td>
    <td width="100">�ۼ��ݾ�</td>
    <td width="100">�ۼ�������</td>

    <td width="110">�����</td>
    <td width="110">������</td>
    
    <td >���</td>
 </tr>
 <% if isArray(ArrRows) then %>
 <% For i =0 To UBound(ArrRows,2) %>
 <%
'  if (isIgnore3Err) then
'     rowErr = ArrRows(5,i)+ArrRows(6,i)+ArrRows(7,i)+ArrRows(8,i)+ArrRows(9,i)+ArrRows(10,i)+ArrRows(11,i)
'  else
'     rowErr = ArrRows(5,i)+ArrRows(6,i)+ArrRows(7,i)+ArrRows(8,i)+ArrRows(9,i)+ArrRows(10,i)+ArrRows(11,i)+ArrRows(12,i)+ArrRows(13,i)+ArrRows(14,i)
'  end if

 
 %>
 <input type="hidden" name="ix" value="<%=i%>">
 <input type="hidden" name="jyyyymm" value="<%=ArrRows(0,i)%>">
 <input type="hidden" name="targetGbn" value="<%=ArrRows(1,i)%>">
 <input type="hidden" name="jgubun" value="<%=ArrRows(2,i)%>">
 <input type="hidden" name="makerid" value="<%=ArrRows(3,i)%>">
 <input type="hidden" name="DLVGbn" value="<%=ArrRows(4,i)%>">
 <input type="hidden" name="vatyn" value="<%=ArrRows(5,i)%>">
 <input type="hidden" name="rowErrNo" value="<%=rowErr%>">
 <tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
    <td><%=ArrRows(0,i)%></td>
    <td><%=ArrRows(1,i)%></td>
    <td><%=ArrRows(2,i)%></td>
    <td><%=ArrRows(3,i)%></td>
    <td><%=ArrRows(4,i)%></td>
    <td><%=ArrRows(5,i)%></td>

    <td><%=ArrRows(6,i)%></td>
    <td align="right"><%=ArrRows(7,i)%></td>
    <td align="right"><%=ArrRows(8,i)%></td>
    <td><%=ArrRows(9,i)%></td>
    <td align="right"><%=ArrRows(10,i)%></td>
    <td align="right"><%=ArrRows(11,i)%></td>

    <td><%=ArrRows(12,i)%></td>
    <td><%=ArrRows(13,i)%></td>
    
    <td>
        <div id="oseq_<%=i%>"></div>
    </td>
</tr>
<% Next %>
<% elseif (yyyy1<>"" and mm1<>""  ) then%>
<tr align="center" bgcolor="FFFFFF" >
    <td colspan="15" >�˻������ �����ϴ�. 
    <% if (NOT isJungsanBatchTargetMaded(yyyy1+"-"+mm1,targetGbn)) then %>
    <input type="button" value="�������ۼ�" onClick="makeJtarget();"> 
    <% end if %>
    </td>
</tr>
<% else %>
<tr align="center" bgcolor="FFFFFF" >
    <td colspan="15">�˻������ �����ϴ�.</td>
</tr>
<% end if %>
</table>
</form>

<form name="frmBatch" method="post" action="dobatch.asp" target="xLink3">
<input type="hidden" name="mode" value="addonebatch<%=targetGbn%>">
<input type="hidden" name="oseq" value="">
<input type="hidden" name="jyyyymm" value="">
<input type="hidden" name="targetGbn" value="">
<input type="hidden" name="jgubun" value="">
<input type="hidden" name="makerid" value="">
<input type="hidden" name="DLVGbn" value="">
<input type="hidden" name="vatyn" value="">

</form>

<form name="frmBatchTarget" method="post" action="dobatch.asp" target="xLink3">
<input type="hidden" name="mode" value="makebatchtarget">
<input type="hidden" name="jyyyymm" value="<%=yyyy1%>-<%=mm1%>">
<input type="hidden" name="targetGbn" value="<%=targetGbn%>">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
