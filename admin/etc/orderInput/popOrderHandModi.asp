<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<%
Dim ooseq : ooseq = requestCheckvar(request("ooseq"),10)
Dim finDiv : finDiv = requestCheckvar(request("finDiv"),10)
Dim csIdx : csIdx = requestCheckvar(request("csIdx"),10)
Dim chOutMallOrderSerial : chOutMallOrderSerial = requestCheckvar(request("chOutMallOrderSerial"),30)
Dim PchOutMallOrderSerial : PchOutMallOrderSerial = requestCheckvar(request("PchOutMallOrderSerial"),30)
Dim orgOutMallOrderSerial : orgOutMallOrderSerial = requestCheckvar(request("orgOutMallOrderSerial"),30)
Dim sellSite : sellSite = requestCheckvar(request("sellSite"),32)
Dim mode : mode = requestCheckvar(request("mode"),10)

Dim i, j
Dim sqlStr,AssignedRow, csExists, CsOrderserial
Dim porderserial

if (finDiv<>"") and (mode="actEtc") then
    if (csIdx="") and ((finDiv="C") or (finDiv="R")) then
        response.write "<script>alert('CS ó����ȣ �ʼ� ����');history.go(-1);</script>"
        dbget.close() : response.end
    end if

    if (finDiv="P") then    ''�ߺ��ּ���ó�� (���� �Ұ�)
        sqlStr = "Update db_temp.dbo.tbl_xSite_tmpOrder " & vbCRLF
        sqlStr = sqlStr & " set ref_outmallorderserial='"&orgOutMallOrderSerial&"'" & vbCRLF
        sqlStr = sqlStr & " ,OutMallOrderSerial='"&chOutMallOrderSerial&"'" & vbCRLF
        sqlStr = sqlStr & " ,etcFinUser='"&session("ssBctID")&"'" & vbCRLF
        sqlStr = sqlStr & " where sellSite='"&sellSite&"'"  & vbCRLF
        sqlStr = sqlStr & " and OutMallOrderSeq="&ooseq&""  & vbCRLF
        sqlStr = sqlStr & " and orderserial is NULL"
        sqlStr = sqlStr & " and matchState='I'"
        dbget.Execute sqlStr,AssignedRow

    elseif (finDiv="Q") then    ''��ó�� �Ϸ᳻��

        sqlStr = "select orderserial from db_temp.dbo.tbl_xSite_tmpOrder "
        sqlStr = sqlStr & " where outmallorderserial='"&PchOutMallOrderSerial&"'"
        sqlStr = sqlStr & " and ref_outmallorderserial='"&orgOutMallOrderSerial&"'"
        sqlStr = sqlStr & " and sellSite='"&sellSite&"'"  & vbCRLF
rw  sqlStr
        porderserial = ""
        rsget.Open sqlStr,dbget,1
        if (not rsget.EOF) then
            porderserial = rsget("orderserial")
        end if
        rsget.Close

        if (porderserial="") then
            response.write "<script>alert('["&porderserial&"] ���ֹ���ȣ �������� ����.');history.go(-1);</script>"
            dbget.close() : response.end
        end if

        sqlStr = "Update db_temp.dbo.tbl_xSite_tmpOrder " & vbCRLF
        sqlStr = sqlStr & " set ref_outmallorderserial='"&PchOutMallOrderSerial&"'" & vbCRLF
        sqlStr = sqlStr & " ,etcFinUser='"&session("ssBctID")&"'" & vbCRLF
        sqlStr = sqlStr & " ,orderserial='"&porderserial&"'"& vbCRLF
        sqlStr = sqlStr & " ,matchState='Q'"
        sqlStr = sqlStr & " where sellSite='"&sellSite&"'"  & vbCRLF
        sqlStr = sqlStr & " and OutMallOrderSeq="&ooseq&""  & vbCRLF
        sqlStr = sqlStr & " and orderserial is NULL"
        sqlStr = sqlStr & " and matchState='I'"
rw sqlStr
        dbget.Execute sqlStr,AssignedRow

        response.end
    elseif (finDiv="D") then    ''��ҿϷ� (��Ȥ CS ���� ���� ��ҷ� �ö���� ��� ����)

        sqlStr = "select id,divcd,currstate,deleteyn,orderserial from db_cs.dbo.tbl_new_as_list "
        sqlStr = sqlStr & " where id="&csIdx&""
        sqlStr = sqlStr & " and divcd='A008'"
        sqlStr = sqlStr & " and deleteyn='N'"
        sqlStr = sqlStr & " and currstate='B007'"

        CsOrderserial = ""
        rsget.Open sqlStr,dbget,1
        if (not rsget.EOF) then
            CsOrderserial = rsget("orderserial")
        end if
        rsget.Close

        if (CsOrderserial="") then
            response.write "<script>alert('["&csIdx&"] �ش� CS ��ȣ�� ���ų� ��ó�� ���� Ȥ�� ��ǰ���� CS���� �ƴմϴ�.');history.go(-1);</script>"
            dbget.close() : response.end
        end if

        sqlStr = "Update db_temp.dbo.tbl_xSite_tmpOrder " & vbCRLF
        sqlStr = sqlStr & " set etcFinUser='"&session("ssBctID")&"'" & vbCRLF
        sqlStr = sqlStr & " ,orderserial='"&CsOrderserial&"'"
        sqlStr = sqlStr & " ,matchState='D'"
        sqlStr = sqlStr & " ,ref_CsID='"&csIdx&"'"
        sqlStr = sqlStr & " where sellSite='"&sellSite&"'"  & vbCRLF
        sqlStr = sqlStr & " and OutMallOrderSeq="&ooseq&""  & vbCRLF
        sqlStr = sqlStr & " and orderserial is NULL"
        sqlStr = sqlStr & " and matchState='I'"

        ''rw sqlStr
        ''response.end
        dbget.Execute sqlStr,AssignedRow
    elseif (finDiv="C") then    ''�������ǰ(�ɼ�)����
        sqlStr = "select id,divcd,currstate,deleteyn,orderserial from db_cs.dbo.tbl_new_as_list "
        sqlStr = sqlStr & " where id="&csIdx&""
        sqlStr = sqlStr & " and divcd='A900'"
        sqlStr = sqlStr & " and title in ('��ǰ�ɼǺ���')"
        sqlStr = sqlStr & " and deleteyn='N'"
        sqlStr = sqlStr & " and currstate='B007'"

        CsOrderserial = ""
        rsget.Open sqlStr,dbget,1
        if (not rsget.EOF) then
            CsOrderserial = rsget("orderserial")
        end if
        rsget.Close

        if (CsOrderserial="") then
            response.write "<script>alert('["&csIdx&"] �ش� CS ��ȣ�� ���ų� ��ó�� ���� Ȥ�� ��ǰ���� CS���� �ƴմϴ�.');history.go(-1);</script>"
            dbget.close() : response.end
        end if

        sqlStr = "Update db_temp.dbo.tbl_xSite_tmpOrder " & vbCRLF
        sqlStr = sqlStr & " set etcFinUser='"&session("ssBctID")&"'" & vbCRLF
        sqlStr = sqlStr & " ,orderserial='"&CsOrderserial&"'"
        sqlStr = sqlStr & " ,matchState='C'"
        sqlStr = sqlStr & " ,ref_CsID='"&csIdx&"'"
        sqlStr = sqlStr & " where sellSite='"&sellSite&"'"  & vbCRLF
        sqlStr = sqlStr & " and OutMallOrderSeq="&ooseq&""  & vbCRLF
        sqlStr = sqlStr & " and orderserial is NULL"
        sqlStr = sqlStr & " and matchState='I'"

        ''rw sqlStr
        ''response.end
        dbget.Execute sqlStr,AssignedRow
    elseif (finDiv="R") then    ''(��)��ȯ/ȸ���Ϸ�
        sqlStr = "select id,divcd,currstate,deleteyn,orderserial from db_cs.dbo.tbl_new_as_list "
        sqlStr = sqlStr & " where id="&csIdx&""
        sqlStr = sqlStr & " and divcd in ('A000','A001','A002')" ''A002 �߰�
        '''sqlStr = sqlStr & " and title in ('��ȯ���','������߼�')"
        sqlStr = sqlStr & " and deleteyn='N'"
        sqlStr = sqlStr & " and currstate in ('B007','B006') " ''��üó���Ϸᵵ ����

        CsOrderserial = ""
        rsget.Open sqlStr,dbget,1
        if (not rsget.EOF) then
            CsOrderserial = rsget("orderserial")
        end if
        rsget.Close

        if (CsOrderserial="") then
            response.write "<script>alert('["&csIdx&"] �ش� CS ��ȣ�� ���ų� ��ó�� ���� Ȥ�� (��)��ȯ���,������߼� CS���� �ƴմϴ�.');history.go(-1);</script>"
            dbget.close() : response.end
        end if

        sqlStr = "Update db_temp.dbo.tbl_xSite_tmpOrder " & vbCRLF
        sqlStr = sqlStr & " set etcFinUser='"&session("ssBctID")&"'" & vbCRLF
        sqlStr = sqlStr & " ,orderserial='"&CsOrderserial&"'"
        sqlStr = sqlStr & " ,matchState='R'"
        sqlStr = sqlStr & " ,ref_CsID='"&csIdx&"'"
        sqlStr = sqlStr & " where sellSite='"&sellSite&"'"  & vbCRLF
        sqlStr = sqlStr & " and OutMallOrderSeq="&ooseq&""  & vbCRLF
        sqlStr = sqlStr & " and orderserial is NULL"
        sqlStr = sqlStr & " and matchState='I'"

        ''rw sqlStr
        ''response.end
        dbget.Execute sqlStr,AssignedRow
    elseif (finDiv="B") then    ''��ǰ�Ϸ�
        sqlStr = "select id,divcd,currstate,deleteyn,orderserial from db_cs.dbo.tbl_new_as_list "
        sqlStr = sqlStr & " where id="&csIdx&""
        sqlStr = sqlStr & " and divcd in ('A004','A010')"
        ''sqlStr = sqlStr & " and title in ('��ȯ���')"
        sqlStr = sqlStr & " and deleteyn='N'"
        sqlStr = sqlStr & " and currstate  in ('B007','B006') " ''��üó���Ϸᵵ ���� 

        CsOrderserial = ""
        rsget.Open sqlStr,dbget,1
        if (not rsget.EOF) then
            CsOrderserial = rsget("orderserial")
        end if
        rsget.Close

        if (CsOrderserial="") then
            response.write "<script>alert('["&csIdx&"] �ش� CS ��ȣ�� ���ų� ��ó�� ���� Ȥ�� ��ǰ/ȸ�� CS���� �ƴմϴ�.');history.go(-1);</script>"
            dbget.close() : response.end
        end if

        sqlStr = "Update db_temp.dbo.tbl_xSite_tmpOrder " & vbCRLF
        sqlStr = sqlStr & " set etcFinUser='"&session("ssBctID")&"'" & vbCRLF
        sqlStr = sqlStr & " ,orderserial='"&CsOrderserial&"'"
        sqlStr = sqlStr & " ,matchState='B'"
        sqlStr = sqlStr & " ,ref_CsID='"&csIdx&"'"
        sqlStr = sqlStr & " where sellSite='"&sellSite&"'"  & vbCRLF
        sqlStr = sqlStr & " and OutMallOrderSeq="&ooseq&""  & vbCRLF
        sqlStr = sqlStr & " and orderserial is NULL"
        sqlStr = sqlStr & " and matchState='I'"

        ''rw sqlStr
        ''response.end
        dbget.Execute sqlStr,AssignedRow
    end if

    if (AssignedRow=1) then
        response.write "<script>alert('�����Ǿ����ϴ�.');opener.location.reload();window.close();</script>"
        dbget.close() : response.end
    else
        response.write "<script>alert('ó���� ����.');</script>"
    end if
end if

Dim otmpOneOrder, otmpOrder
set otmpOneOrder = new CxSiteTempOrder
otmpOneOrder.FRectOutMallOrderSeq   = ooseq
otmpOneOrder.getOneTmpOrder

IF otmpOneOrder.FResultCount<1 then
    rw "�˻� ����� �����ϴ�."
    dbget.Close() : response.end
end if

Dim OutMallOrderSerial
IF Not IsNULL(otmpOneOrder.FOneItem.FRef_OutMallOrderSerial) then
    OutMallOrderSerial = otmpOneOrder.FOneItem.FRef_OutMallOrderSerial
else
    OutMallOrderSerial = otmpOneOrder.FOneItem.FOutMallOrderSerial
end if

set otmpOrder = new CxSiteTempOrder
otmpOrder.FPageSize = 100
otmpOrder.FCurrPage = 1
otmpOrder.FRectSellSite   = otmpOneOrder.FOneItem.FSellSite
''otmpOrder.FRectMatchState = matchState
''otmpOrder.FRectorderserial=orderserial
otmpOrder.FRectoutmallorderserial=OutMallOrderSerial
otmpOrder.getOnlineTmpOrderList(true)

Dim TenOrderserialArr
IF (otmpOrder.FResultCount>0) then
    for i=0 to otmpOrder.FResultCount-1
        If InStr(TenOrderserialArr,otmpOrder.FItemList(i).FOrderserial)>0 then

        else
            IF (otmpOrder.FItemList(i).FOrderserial<>"") then
                TenOrderserialArr = TenOrderserialArr & otmpOrder.FItemList(i).FOrderserial & ","
            END IF
        end if
    next
ENd If

Dim buf, mxBuf
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

dim ocsArrList, csCnt : csCnt = 0
IF (TenOrderserialArr<>"") then
    IF Right(TenOrderserialArr,1)="," then TenOrderserialArr=Left(TenOrderserialArr,Len(TenOrderserialArr)-1)
    TenOrderserialArr = replace(TenOrderserialArr,",","','")

    rw "TenOrderserialArr="&TenOrderserialArr
    sqlStr = "select top 20 A.id,A.orderserial,A.divcd,A.currstate,A.title,A.deleteyn,A.regdate,A.finishdate,A.writeuser,A.finishuser"
    sqlStr = sqlStr & " ,C.comm_name,C2.comm_name as stateNm"
    sqlStr = sqlStr & " ,A.contents_jupsu,A.contents_finish"
    sqlStr = sqlStr & " ,A.songjangno"
    sqlStr = sqlStr & " ,S.divname,S.findURL"
    sqlStr = sqlStr & " from db_cs.dbo.tbl_new_as_list A"
    sqlStr = sqlStr & "     left join db_cs.dbo.tbl_cs_comm_code C"
    sqlStr = sqlStr & "     on A.divcd=C.comm_cd"
    sqlStr = sqlStr & "     left join db_cs.dbo.tbl_cs_comm_code C2"
    sqlStr = sqlStr & "     on A.currstate=C2.comm_cd"
    sqlStr = sqlStr & "     left join db_order.dbo.tbl_songjang_div S"
	sqlStr = sqlStr & "     on A.songjangdiv=S.divcd"
    sqlStr = sqlStr & " where A.orderserial in ('"&TenOrderserialArr&"')"
    sqlStr = sqlStr & " order by A.id desc"

    rsget.Open sqlStr,dbget,1
    if Not rsget.Eof then
        ocsArrList = rsget.getRows()
    end if
    rsget.Close
end if

IF IsArray(ocsArrList) then
    csCnt = UBound(ocsArrList,2)+1
end if

dim OCsDetail, isChangeOutOrderSerialValid

rw OutMallOrderSerial
rw otmpOneOrder.FOneItem.ForderItemID
rw otmpOneOrder.FOneItem.ForderItemOption

Dim pInsOrder : pInsOrder=False
Dim pInsOutMallOrderserial
for i=0 to otmpOrder.FResultCount-1
    if Not IsNULL(otmpOrder.FItemList(i).Fref_outmallorderserial) then
        if (otmpOrder.FItemList(i).Fref_outmallorderserial=otmpOneOrder.FOneItem.FOutMallOrderSerial) _
            and (otmpOrder.FItemList(i).ForderItemID=otmpOneOrder.FOneItem.ForderItemID) _
                and (otmpOrder.FItemList(i).ForderItemOption=otmpOneOrder.FOneItem.ForderItemOption) then
            pInsOrder = true
            pInsOutMallOrderserial = otmpOrder.FItemList(i).Foutmallorderserial
        end if
    end if
next
%>
<script language='javascript'>
function popEtcOrdFinish(actTp,ooseq){
    var popwin = window.open('popEtcOrdFinish.asp?actTp='+actTp+'&ooseq='+ooseq,'popEtcOrdFinish','scrollbars=yes,resizable=yes,width=600,height=300');
    popwin.focus();
}

function finThis(){
    var frm=document.frmAct;
    if (frm.finDiv.value.length<1){
        alert('ó�������� �����ϼ���.');
        frm.finDiv.focus();
        return;
    }

    if (((frm.finDiv.value=="R")||(frm.finDiv.value=="C")||(frm.finDiv.value=="D"))&&(frm.csIdx.value.length<1)){
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

    if ((pval=="R")||(pval=="C")||(pval=="D")||(pval=="B")){
        document.getElementById("selDiv_R").style.display="block";
        document.getElementById("selDiv_P").style.display="none";
        document.getElementById("selDiv_Q").style.display="none";
    }else if (pval=="P"){
        document.getElementById("selDiv_R").style.display="none";
        document.getElementById("selDiv_P").style.display="block";
        document.getElementById("selDiv_Q").style.display="none";
    }else if (pval=="Q"){
        document.getElementById("selDiv_R").style.display="none";
        document.getElementById("selDiv_P").style.display="none";
        document.getElementById("selDiv_Q").style.display="block";
    }else{
        document.getElementById("selDiv_R").style.display="none";
        document.getElementById("selDiv_P").style.display="none";
        document.getElementById("selDiv_Q").style.display="none";
    }


}
</script>
<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frmAct" method="post">
    <input type="hidden" name="mode" value="actEtc">
    <input type="hidden" name="sellSite" value="<%= otmpOneOrder.FOneItem.FsellSite %>">
    <input type="hidden" name="ooseq" value="<%= otmpOneOrder.FOneItem.FOutMallOrderSeq %>">
    <input type="hidden" name="orgOutMallOrderSerial" value="<%= otmpOneOrder.FOneItem.FOutMallOrderSerial %>">
<tr align="center" height="25" bgcolor="#E8E8FF">
    <td width="20" ></td>
    <td width="50" >�ֹ�����</td>
	<td >TEN �ֹ���ȣ</td>
	<td >���޻� �ֹ���ȣ</td>
	<td>��ǰ�ڵ�</td>
	<td>�ɼ��ڵ�</td>
	<td>��ǰ��</td>
	<td>�ɼǸ�</td>
	<td >���</td>
</tr>
<% for i=0 to otmpOrder.FResultCount-1 %>
<tr align="center" height="25" bgcolor="#FFFFFF">
    <td rowspan="2"><%= CHKIIF(CStr(otmpOrder.FItemList(i).FOutMallOrderSeq)=ooseq,"<b><font color=red>&gt;</font></b>","") %></td>
    <td rowspan="2"><%= otmpOrder.FItemList(i).getOrderCsGbnName %>
	<% if (otmpOrder.FItemList(i).FDuppExists) then %>
	<br>��ǰ�ߺ�
	<% end if %>
	<% if (otmpOrder.FItemList(i).FaddDlvExists) then %>
	<br>�ټ�����
	<% end if %>
    <td  rowspan="2"><%=otmpOrder.FItemList(i).FOrderSerial %></td>
    <td  rowspan="2"><%=otmpOrder.FItemList(i).FOutMallOrderSerial %></td>
    <td><%=otmpOrder.FItemList(i).ForderItemID %></td>
    <td><%=otmpOrder.FItemList(i).ForderItemOption %></td>
    <td><%=otmpOrder.FItemList(i).ForderItemName %></td>
    <td><%=otmpOrder.FItemList(i).ForderItemOptionName %></td>
    <td width=250  rowspan="2">
    <% if CStr(otmpOrder.FItemList(i).FOutMallOrderSeq)=ooseq then %>

        <% if (application("Svr_Info")="Dev") or isNULL(otmpOrder.FItemList(i).FOrderSerial) then %>
        <%
            isChangeOutOrderSerialValid = FALSE
            isChangeOutOrderSerialValid = (otmpOrder.FItemList(i).getOrderCsGbnName="")
            isChangeOutOrderSerialValid = isChangeOutOrderSerialValid and isNULL(otmpOrder.FItemList(i).FOrderSerial)
            isChangeOutOrderSerialValid = isChangeOutOrderSerialValid and otmpOrder.FItemList(i).FaddDlvExists
        %>
            <select name="finDiv" onChange="chgDiv(this)">
            <option value="">����
            <option value="R">(��)��ȯ ��� �Ϸ�
            <option value="C">�������ǰ(�ɼ�)����
            <option value="D">��ҿϷ�
            <option value="B">��ǰ�Ϸ�
            <% if (isChangeOutOrderSerialValid) or (otmpOrder.FItemList(i).FDuppExists) or (otmpOneOrder.FOneItem.FsellSite = "gseshop") then %>
            <option value="P">�ߺ��ּ���(��ǰ)ó��(�ű��ֹ���ȣ����)
            <% end if %>
            <% if (pInsOrder) then %>
            <option value="Q">��ó������
            <% end if %>
            </select>
            <div id="selDiv_R" name="selDiv_R" style="display:none">
                CSó�� ��ȣ : <input type="text" name="csIdx" value="" size="5" maxlength="9"> <input type="button" value="�Ϸ�ó��" onClick="finThis();">
            </div>

            <div id="selDiv_P" name="selDiv_P" style="display:none">
                �ű� �ֹ���ȣ : <input type="text" name="chOutMallOrderSerial" value="<%=OutMallOrderSerial%>_<%=mxBuf%>"> <input type="button" value="�Ϸ�ó��" onClick="finThis();">
            </div>

            <div id="selDiv_Q" name="selDiv_Q" style="display:none">
                ��ó�� �ֹ���ȣ : <input type="text" name="PchOutMallOrderSerial" value="<%=pInsOutMallOrderserial%>"> <input type="button" value="�Ϸ�ó��" onClick="finThis();">
            </div>
        <% end if %>
    <% end if %>
    </td>
</tr>
<tr align="center" height="25" bgcolor="#FFFFFF">
    <td ><%= otmpOrder.FItemList(i).FReceiveName %></td>
    <td colspan="3">
    [<%= otmpOrder.FItemList(i).FReceiveZipCode %>]
    &nbsp;
    <%= otmpOrder.FItemList(i).FReceiveAddr1 %>
    &nbsp;
    <%= otmpOrder.FItemList(i).FReceiveAddr2 %>
    </td>
</tr>
<% next %>
</form>
</table>

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<% if csCnt>0 then %>
<tr align="center" height="25" bgcolor="#E8E8FF">
    <td>CSID</td>
	<td>�ֹ���ȣ</td>
	<td>��������</td>
	<td >Title</td>
	<td>����/����/ó����</td>
	<td>���ü���</td>
	<td>���</td>
</tr>
<% for i=0 to csCnt-1 %>
<tr height="25" bgcolor="<%= CHKIIF(ocsArrList(5,i)="N","#FFFFFF","#CCCCCC") %>">
    <td><%= ocsArrList(0,i) %></td>
    <td><%= ocsArrList(1,i) %></td>
    <td><%= ocsArrList(10,i) %></td>
    <td><%= ocsArrList(4,i) %></td>
    <td><%= ocsArrList(11,i) %><%= CHKIIF(ocsArrList(5,i)<>"N","<font color=red>(����)</font>","") %></td>
    <td rowspan="3"><%= ocsArrList(15,i) %><%= ocsArrList(14,i) %></td>
    <td></td>
</tr>

<tr height="25" bgcolor="<%= CHKIIF(ocsArrList(5,i)="N","#FFFFFF","#CCCCCC") %>">
    <td align="right"> - ���� </td>
    <td colspan="3"><%= ocsArrList(12,i) %></td>
    <td><%= ocsArrList(6,i) %></td>
    <td></td>
</tr>
<tr height="25" bgcolor="<%= CHKIIF(ocsArrList(5,i)="N","#FFFFFF","#CCCCCC") %>">
    <td align="right"> - ó�� </td>
    <td colspan="3"><%= ocsArrList(13,i) %></td>
    <td><%= ocsArrList(7,i) %></td>
    <td></td>
</tr>
<%
set OCsDetail = new CCSASList
OCsDetail.FRectCsAsID = ocsArrList(0,i)
OCsDetail.GetCsDetailList
    for j=0 to OCsDetail.FResultCount-1
%>
<tr height="25" bgcolor="<%= CHKIIF(ocsArrList(5,i)="N","#FFFFFF","#CCCCCC") %>">
    <td align="right" colspan="2"> - �� </td>
    <td ><%= OCsDetail.FItemList(j).FItemId %></td>
    <td><%= OCsDetail.FItemList(j).FItemName %>
    <% if OCsDetail.FItemList(j).Fitemoptionname<>"" then %>
    [<%=OCsDetail.FItemList(j).Fitemoptionname%>]
    <% end if %>
    </td>
    <td><%= OCsDetail.FItemList(j).Fregitemno %></td>
    <td></td>
    <td></td>
</tr>
<%
    next
set OCsDetail = nothing
%>
<tr>
    <td colspan="7"></td>
</tr>
<% next %>
<% else %>
<tr align="center" height="25" bgcolor="#E8E8FF">
    <td align="center"> ���� CS �� ����</td>
</tr>
<% end if %>
</table>
<%
set otmpOneOrder = Nothing
set otmpOrder = Nothing

%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
