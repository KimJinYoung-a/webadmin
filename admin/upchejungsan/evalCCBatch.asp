<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ������ ���ݰ�꼭 ���� ��ġ
' Hieditor : 2019.04.04 ������ ����
'            2021.02.03 �ѿ�� ����(��36524 �÷��ÿ������� ���� api���� �������� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/jungsan/jungsanTaxCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim yyyy1 : yyyy1 = requestCheckvar(request("yyyy1"),4)
Dim mm1   : mm1   = requestCheckvar(request("mm1"),2)
Dim targetGbn : targetGbn = requestCheckVar(request("targetGbn"),10)
Dim jungsan_date : jungsan_date = requestCheckvar(request("jungsan_date"),10)
Dim itopn : itopn   = requestCheckvar(request("itopn"),10)

if (yyyy1="") then
    yyyy1 = LEFT(dateadd("m",-1,now()),4)
    mm1 = MID(dateadd("m",-1,now()),6,2)
end if

Dim yyyymm : yyyymm = yyyy1&"-"&mm1


Dim research : research = requestCheckvar(request("research"),32)
Dim iMaxpageSize : iMaxpageSize=1000

if (itopn="") then itopn=iMaxpageSize
''dim ArrRows : ArrRows = getEvalCCBatchTargetList(yyyymm,targetGbn,jungsan_date,iMaxpageSize)

dim ojungsanTax
set ojungsanTax = new CUpcheJungsanTax
ojungsanTax.FCurrPage = 1
ojungsanTax.FPageSize = itopn
ojungsanTax.FRectYYYYMM = yyyy1+"-"+mm1
ojungsanTax.FRecttargetGbn = targetGbn
ojungsanTax.FRectJungsanDate = jungsan_date

ojungsanTax.getEvalJungsanTaxTargetListAdm


dim i, ttlCnt : ttlCnt = 0
ttlCnt = ojungsanTax.FResultCount

Dim rowErr : rowErr=0
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">
<script language='javascript'>
var batchstarted = false;
var nextid = 0;
var evalWinNm ="_evalWin" + (new Date()).getTime();
function EvalBatchStart(comp){
    if (batchstarted) return;

    comp.style="background-color: border: 1px solid #999999;#cccccc;color: #888888;"
    comp.value="������ ....."
    comp.disabled = true;
    batchstarted = true;

    addNotiLog("start");

    var popwin = window.open("" ,evalWinNm,"width=1200,height=768,status=no, scrollbars=auto, menubar=no, resizable=yes");
	popwin.focus();

    fnNextEvalProc();
}

function addNotiLog(ilog){
    document.getElementById("disp1").innerHTML += ilog+"<br>";
}

function addResultLog(orderSeq,ilog){
    document.getElementById("oseq_"+orderSeq).innerHTML = ilog;
}

function fnNextEvalProc(){
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
        oneEvalProc(ix);
    }, 100);  

    
}

function oneEvalProc(iidx){
    nextid = iidx+1;
    var arrfrm = document.frmBatchArr;

    if (!arrfrm.ix.length){
        if (arrfrm.rowErrNo.value*1>0){
            addResultLog(arrfrm.jidx.value,"skip");
            fnNextEvalProc();    
        }else{
            document.frmBatch.makerid.value = arrfrm.makerid.value;
            document.frmBatch.yyyy1.value = arrfrm.yyyy1.value;
            document.frmBatch.mm1.value = arrfrm.mm1.value;
            document.frmBatch.onoffGubun.value = arrfrm.onoffGubun.value;
            document.frmBatch.jidx.value = arrfrm.jidx.value;
            
            addNotiLog(document.frmBatch.onoffGubun.value + ":" + document.frmBatch.jidx.value);
            document.frmBatch.target=evalWinNm;
            document.frmBatch.submit();
        }
    }else{
        if (arrfrm.rowErrNo[iidx].value*1>0){
            addResultLog(arrfrm.jidx[iidx].value,"skip");
            fnNextEvalProc();
        }else{
            document.frmBatch.makerid.value = arrfrm.makerid[iidx].value;
            document.frmBatch.yyyy1.value = arrfrm.yyyy1[iidx].value;
            document.frmBatch.mm1.value = arrfrm.mm1[iidx].value;
            document.frmBatch.onoffGubun.value = arrfrm.onoffGubun[iidx].value;
            document.frmBatch.jidx.value = arrfrm.jidx[iidx].value;

            addNotiLog(document.frmBatch.onoffGubun.value + ":" + document.frmBatch.jidx.value);
            document.frmBatch.target=evalWinNm;
            document.frmBatch.submit();
        }
    }
    

    
}
</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
	    * ���� ��� ��� : <% DrawYMBox yyyy1,mm1 %>
        &nbsp;
        ON/AC/OF ���� :
        <select name="targetGbn" >
		<option value="">��ü
		<option value="ON" <%= CHKIIF(targetGbn="ON","selected","") %> >ON
		<option value="OF" <%= CHKIIF(targetGbn="OF","selected","") %> >OF
		<option value="AC" <%= CHKIIF(targetGbn="AC","selected","") %> >AC
		</select>

        &nbsp;
        ������ :
        <select name="jungsan_date">
        <option value="" <% if jungsan_date="" then response.write "selected" %> >����
        <option value="15��" <% if jungsan_date="15��" then response.write "selected" %> >15��
        <option value="����" <% if jungsan_date="����" then response.write "selected" %> >����
        <option value="����" <% if jungsan_date="����" then response.write "selected" %> >����
        </select>

        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        �˻��� : 
        <select name="itopn">
        <option value="1" <% if itopn="1" then response.write "selected" %> >1
        <option value="10" <% if itopn="10" then response.write "selected" %> >10
        <option value="100" <% if itopn="100" then response.write "selected" %> >100
        <option value="1000" <% if itopn="1000" then response.write "selected" %> >1000
        </select>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</table>
</form>
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
        <input type="button" class="button" value="�ϰ�����" onClick="EvalBatchStart(this);">
    </td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<form name="frmBatchArr" style="margin:0px;" >
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
		�˻���� : <b><%= ttlCnt %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="60"></td>
    <td width="90">�����</td>
    <td width="90">����ó</td>
    <td width="80">������</td>
    <td width="60">��꼭����</td>

    <td width="60">�׷��ڵ�</td>
    <td width="100">�귣��ID</td>
    <td >����</td>
    <td width="50">������</td>
    <td width="70">���ް�</td>
    <td width="70">�ΰ���</td>
    <td width="70">�հ�</td>
    <td width="60">����</td>
    <td >���</td>
 </tr>
 <% if ojungsanTax.FResultCount>0 then %>
 <% For i =0 To ojungsanTax.FResultCount-1 %>
 <%
 rowErr = 0
 %>
 <input type="hidden" name="ix" value="<%=i%>">
 <input type="hidden" name="yyyy1" value="<%=LEFT(ojungsanTax.FItemList(i).FYYYYMM,4)%>">
 <input type="hidden" name="mm1" value="<%=Right(ojungsanTax.FItemList(i).FYYYYMM,2)%>">
 <input type="hidden" name="onoffGubun" value="<%= ojungsanTax.FItemList(i).FtargetGbn %>">
 <input type="hidden" name="jidx" value="<%= ojungsanTax.FItemList(i).Fid %>">
 <input type="hidden" name="makerid" value="<%= ojungsanTax.FItemList(i).Fmakerid %>">
 <input type="hidden" name="rowErrNo" value="<%=rowErr%>">

 <tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="f1f1f1"; onmouseout=this.style.background="#FFFFFF";>
    <td>
        <%= ojungsanTax.FItemList(i).Fid %>
    </td>
    <td><%=ojungsanTax.FItemList(i).FYYYYMM%></td>
    <td><%=ojungsanTax.FItemList(i).getTargetNm%></td>
    <td><%=ojungsanTax.FItemList(i).getTaxJungsanGubun%></td>
    <td><%=ojungsanTax.FItemList(i).getTaxTypeStrUpcheView%></td>
    <td><%=ojungsanTax.FItemList(i).Fgroupid%></td>
    <td><%=ojungsanTax.FItemList(i).Fmakerid%></td>
    <td align="left"><%=ojungsanTax.FItemList(i).Ftitle%></td>
    <td><%= ojungsanTax.FItemList(i).Ftaxregdate %></td>
    <td align="right"><%= FormatNumber(ojungsanTax.FItemList(i).getJungsanTaxSuply,0) %></td>
    <td align="right"><%= FormatNumber(ojungsanTax.FItemList(i).getJungsanTaxVat,0) %></td>
    <td align="right"><%= FormatNumber(ojungsanTax.FItemList(i).getJungsanTaxSum,0) %></td>
    <td><%= ojungsanTax.FItemList(i).GetTaxEvalStateName %></td>
    <td>
        <div id="oseq_<%=ojungsanTax.FItemList(i).Fid%>"></div>
    </td>
</tr>
<% Next %>

<% else %>
<tr align="center" bgcolor="FFFFFF" >
    <td colspan="16">�˻������ �����ϴ�.</td>
</tr>
<% end if %>
</table>
</form>

<%
'popTaxRegAdmin.asp?makerid=" + makerid + "&yyyy1=" + yyyy1 + "&mm1=" + mm1 + "&onoffGubun=" + onoffGubun + "&jidx="+jidx+"&isauto=on&nextjidx="+nextid
%>

<% '<form name="frmBatch" method="get" action="popTaxRegAdmin.asp" style="margin:0px;" > %>
<form name="frmBatch" method="get" action="/admin/upchejungsan/popTaxRegAdminapi.asp" style="margin:0px;" >
<input type="hidden" name="makerid" value="">
<input type="hidden" name="yyyy1" value="">
<input type="hidden" name="mm1" value="">
<input type="hidden" name="onoffGubun" value="">
<input type="hidden" name="jidx" value="">

<input type="hidden" name="isauto" value="on">
<input type="hidden" name="autotype" value="V2">
</form>
<%
public function getEvalCCBatchTargetList(yyyymm,targetGbn,JungsanDate,iPageSize)
    Dim sqlStr
    sqlStr = "db_jungsan.dbo.[usp_TEN_JungsanBatchEvalTarget] '"&yyyymm&"','"&targetGbn&"','"&JungsanDate&"',"&iPageSize&""
    rsget.CursorLocation = adUseClient
    rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
    if not rsget.Eof then
        getEvalCCBatchTargetList = rsget.GetRows()
    end if
    rsget.close

end function

set ojungsanTax = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
