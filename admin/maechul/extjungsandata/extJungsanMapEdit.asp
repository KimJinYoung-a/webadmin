<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsancls.asp"-->
<!-- #include virtual="/cscenter/delivery/deliveryTrackCls.asp" -->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim FormatDotNo : FormatDotNo=0
dim research, page
dim sellsite, jungsantype, searchfield, searchtext, tmpcssearch

Dim i

research = requestCheckvar(request("research"),10)
page = requestCheckvar(request("page"),10)
if (page="") then page = 1

sellsite		= requestCheckvar(request("sellsite"),32)
jungsantype		= requestCheckvar(request("jungsantype"),32)
searchfield 	= requestCheckvar(request("searchfield"),32)
searchtext 		= Replace(Replace(requestCheckvar(request("searchtext"),32), "'", ""), Chr(34), "")
tmpcssearch     = requestCheckvar(request("tmpcssearch"),1)


Dim oCExtJungsan
set oCExtJungsan = new CExtJungsan
	oCExtJungsan.FPageSize = 25
	oCExtJungsan.FCurrPage = page

	oCExtJungsan.FRectSellSite = sellsite
	oCExtJungsan.FRectJungsanType = jungsantype

	oCExtJungsan.FRectSearchField = searchfield
	oCExtJungsan.FRectSearchText = searchtext

    oCExtJungsan.GetExtJungsanMapCheckList

Dim oCExtJungsanOrderTmp
Set oCExtJungsanOrderTmp = new CExtJungsan
    oCExtJungsanOrderTmp.FPageSize = 25
	oCExtJungsanOrderTmp.FCurrPage = page

	oCExtJungsanOrderTmp.FRectSellSite = sellsite

	oCExtJungsanOrderTmp.FRectSearchField = searchfield
	oCExtJungsanOrderTmp.FRectSearchText = searchtext

    oCExtJungsanOrderTmp.GetExtJungsanMapCheckListTmpOrder

Dim mapTenOrderserial, mapRtnTenOrderserial
Dim mapTenOrderserial2
%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.action = "";
    document.frm.submit();
}

function jsSubmit(){
    document.frm.page.value = "1";
    document.frm.action = "";

    document.frm.submit();
}

function popcenter_Action_List(orderserial) {
    var window_width = 1280;
    var window_height = 960;
	var popwin = window.open("<%=replace(manageUrl,"http://","https://")%>/cscenter/action/cs_action.asp?orderserial=" + orderserial ,"cs_action_pop","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
}

function chgCompValChk(comp,ix){
    var frm = comp.form;

    if (comp.value.length>3){
        if (frm.chkix.length>1){
            frm.chkix[ix].checked=true;
            AnCheckClick(frm.chkix[ix]);
        }else{
            frm.chkix.checked=true;
            AnCheckClick(frm.chkix);
        }
    }

}

function switchCheckBox(comp){
    var frm = comp.form;

    if(frm.chkix.length>1){
        for(i=0;i<frm.chkix.length;i++){
            if (!frm.chkix[i].disabled){
                frm.chkix[i].checked = comp.checked;
                AnCheckClick(frm.chkix[i]);
            }
        }
    }else{
        if (!frm.chkix.disabled){
            frm.chkix.checked = comp.checked;
            AnCheckClick(frm.chkix);
        }
    }
}

function chkNChangeVal(comp){
    var frm = comp.form;
    var pass = false;

    if (!frm.chkix){
        alert("���� ������ �����ϴ�.");
        return;
    }

    if(frm.chkix.length>1){
        for (var i=0;i<frm.chkix.length;i++){
            pass = (pass||frm.chkix[i].checked);
        }
    }else{
        pass = frm.chkix.checked;
    }

    if (!pass) {
        alert("���� ������ �����ϴ�.");
        return;
    }

    if(frm.chkix.length>1){
        for (var i=0;i<frm.chkix.length;i++){
            if (frm.chkix[i].checked){
                if (frm.OrgOrderserialArr[i].value.length!=11){
                    alert("TEN �ֹ���ȣ�� �Է� �Ͻñ� �ٶ��ϴ�.(11�ڸ�)");
                    frm.OrgOrderserialArr[i].focus();
                    return;
                }else if (frm.itemidArr[i].value.length<1){
                    alert("��ǰ�ڵ带 �Է��Ͻñ� �ٶ��ϴ�.");
                    frm.itemidArr[i].focus();
                    return;
                }else if (frm.itemoptionArr[i].value.length!=4){
                    alert("�ɼ��ڵ带 �Է��Ͻñ� �ٶ��ϴ�.");
                    frm.itemoptionArr[i].focus();
                    return;
                }
            }
        }
    }else{
        if (frm.chkix.checked){
            if (frm.OrgOrderserialArr.value.length<1){
                alert("TEN �ֹ���ȣ�� �Է� �Ͻñ� �ٶ��ϴ�.");
                frm.OrgOrderserialArr.focus();
                return;
            }else if (frm.itemidArr.value.length<1){
                alert("�����ȣ�� �Է��Ͻñ� �ٶ��ϴ�.");
                frm.itemidArr.focus();
                return;
            }else if (frm.itemoptionArr.value.length!=4){
                alert("�ɼ��ڵ带 �Է��Ͻñ� �ٶ��ϴ�.");
                frm.itemoptionArr.focus();
                return;
            }
        }
    }


    if (confirm("���� ������ ���� �Ͻðڽ��ϱ�?")){
        frm.mode.value="chgmaporder";
        frm.submit();
    }
}

function copyOserial(kk){
    var frm = document.frm1;

    if(frm.cpchhk.length>1){
        for(i=0;i<frm.cpchhk.length;i++){
            if (frm.cpchhk[i].checked){
                frm.OrgOrderserialArr[i].value = frm.OrgOrderserialArr[kk].value;
                frm.itemidArr[i].value = frm.itemidArr[kk].value;
                frm.itemoptionArr[i].value = frm.itemoptionArr[kk].value;
                frm.chkix[i].checked=true;
                AnCheckClick(frm.chkix[i]);
                frm.cpchhk[i].checked = false;
            }
        }
    }else{
        if (frm.cpchhk.checked){
            frm.OrgOrderserialArr[frm.cpchhk.value].value = frm.OrgOrderserialArr[kk].value;
            frm.itemidArr[frm.cpchhk.value].value = frm.itemidArr[kk].value;
            frm.itemoptionArr[frm.cpchhk.value].value = frm.itemoptionArr[kk].value;
            frm.chkix[frm.cpchhk.value].checked=true;
            AnCheckClick(frm.chkix[frm.cpchhk.value]);
            frm.cpchhk.checked = false;
        }
    }
}

function jsOrgOrderSerialInput(isellsite,iextOrderserial,iextOrderserSeq,iextorgorderserial){
    var extorgorderserial = prompt("extorgorderserial", iextorgorderserial);

    if (confirm('���ֹ���ȣ�� : '+extorgorderserial+' ���� �Ͻðڽ��ϱ�?')){
        var frm = document.extEdtFrm;
        frm.sellsite.value=isellsite;
        frm.extOrderserial.value=iextOrderserial;
        frm.extOrderserSeq.value=iextOrderserSeq;
        frm.extorgorderserial.value=extorgorderserial;

        frm.submit();
    }
}

function jsSliceItemno(isellsite,iextOrderserial,iextOrderserSeq,iextItemNo){
    var iSliceItemno;

    if (iextItemNo==2) {
        iSliceItemno=1;
    }else if (iextItemNo==-2) {
        iSliceItemno=-1;
    }else{
        iSliceItemno = prompt("SliceNum", "0");
        if (iSliceItemno == null) return;
    }
    iSliceItemno = iSliceItemno*1;

    if (!Number.isInteger(iSliceItemno)){
        alert('���ڸ� �Է��ϼ���.');
        return;
    }

    if (iSliceItemno*1==0){
        alert('0 �� �ƴ� ���ڸ� �Է��ϼ���.');
        return;
    }

    if (Math.abs(iSliceItemno)>=Math.abs(iextItemNo)){
        alert('������ ���� ���� ���� �Է��ϼ���.');
        return;
    }

    if ((iSliceItemno>0&&iextItemNo*1<0)||(iSliceItemno<0&&iextItemNo*1>0)){
        alert('����� �����, ������ ������ ������ �ֽ��ϴ�.');
        return;
    }

    if (confirm('������ '+iextItemNo+'���� '+ iSliceItemno+ '�� / '+ (iextItemNo*1-iSliceItemno*1)+ '�� �� �����ðڽ��ϱ�?')){
        var frm = document.slicefrm;
        frm.sellsite.value=isellsite;
        frm.extOrderserial.value=iextOrderserial;
        frm.extOrderserSeq.value=iextOrderserSeq;
        frm.newSliceNo.value=iSliceItemno;

        frm.submit();
    }

}

function popJcomment(iorderserial,iitemid,iitemoption){
    var addcmt = "";
    addcmt = prompt("���� comment", "");
    if (addcmt == null) return;

    if (addcmt.length<1){
        alert("�ڸ�Ʈ�� �ۼ����ּ���.");
        return;
    }

    var frm = document.frmcmt;
    frm.mode.value="addcmt";
    frm.orderserial.value=iorderserial;
    frm.itemid.value=iitemid;
    frm.itemoption.value=iitemoption;
    frm.addcomment.value=addcmt;

    frm.submit();
}

function delJcomment(rowidx){
    if (confirm("���� ���� �ڸ�Ʈ�� ���� �Ͻðڽ��ϱ�?")){
        var frm = document.frmcmt;
        frm.mode.value="delcmt";
        frm.rowidx.value=rowidx;

        frm.submit();
    }
}

function chgTmpOrderRealsellprice(ioutmallorderseq,orgrealsellprice){
    var chgrealsellprice = "";
    chgrealsellprice = prompt("�����ұݾ�", "");
    if (chgrealsellprice == null) return;

    if (chgrealsellprice.length<1){
        alert("���ǸŰ��� �Է��ϼ���.");
        return;
    }

    if (!IsDigit(chgrealsellprice)){
        alert('���ڸ� �Է��ϼ���.');
        return;
    }

    var frm = document.frmXsiteTmpval;
    frm.mode.value="chgrealsellprice";
    frm.outmallorderseq.value=ioutmallorderseq;
    frm.chgval.value=chgrealsellprice;

    if (confirm("�ӽ� �ֹ����� ���ǸŰ� ���� "+orgrealsellprice+" => "+chgrealsellprice+" �� �����Ͻðڽ��ϱ�?")){
        frm.submit();
    }
}

function chgTmpOrderMatchitemOption(ioutmallorderseq,orgmatchitemoption){
    var chgmatchitemoption = "";
    chgmatchitemoption = prompt("�����ҿɼ��ڵ�", "");
    if (chgmatchitemoption == null) return;

    if (chgmatchitemoption.length!=4){
        alert("�ɼ��ڵ� 4�ڸ��� �Է��ϼ���.");
        return;
    }


    var frm = document.frmXsiteTmpval;
    frm.mode.value="chgmatchitemoption";
    frm.outmallorderseq.value=ioutmallorderseq;
    frm.chgval.value=chgmatchitemoption;

    if (confirm("�ӽ� �ֹ����� ��Ī �ɼ��ڵ� ���� "+orgmatchitemoption+" => "+chgmatchitemoption+" �� �����Ͻðڽ��ϱ�?")){
        frm.submit();
    }
}

function popDeliveryTrackingSummaryOne(iorderserial,isongjangno,isongjangdiv){
    var iurl = "/cscenter/delivery/DeliveryTrackingSummaryOne.asp?songjangno="+isongjangno+"&orderserial="+iorderserial+"&songjangdiv="+isongjangdiv;
    var popwin = window.open(iurl,'DeliveryTrackingSummaryOne','width=1200 height=800 scrollbars=yes resizable=yes');
    popwin.focus();

}

</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		&nbsp;
		���޸�:
        <%= getJungsanXsiteComboHTML("sellsite",sellsite,"") %>

		&nbsp;
		������:
		<select class="select" name="jungsantype">
			<option></option>
			<option value="C" <% if (jungsantype = "C") then %>selected<% end if %> >��ǰ��(�Һ��ڸ���)</option>
			<option value="D" <% if (jungsantype = "D") then %>selected<% end if %> >��ۺ�</option>
			<option value="E" <% if (jungsantype = "E") then %>selected<% end if %> >��Ÿ����</option>
		</select>

	</td>
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="jsSubmit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		&nbsp;
		* �˻����� :
		<select class="select" name="searchfield">
			<option value=""></option>
			<option value="extOrderserial" <% if (searchfield = "extOrderserial") then %>selected<% end if %> >�����ֹ���ȣ</option>
			<option value="OrgOrderserial" <% if (searchfield = "OrgOrderserial") then %>selected<% end if %> >����(TEN)�ֹ���ȣ</option>
            <option value="extitemid" <% if (searchfield = "extitemid") then %>selected<% end if %> >���޻�ǰ�ڵ�</option>
		</select>
		<input type="text" class="text" name="searchtext" size="30" value="<%= searchtext %>">

	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<p  >
     <%= getExtsongjangInputNOTIStr %>
<p  >
<!-- ���� ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm1" method="post" action="extJungsan_process.asp">
<input type="hidden" name="mode" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="21" align="right">
		<input type="button" value="���ó��� ����" onClick="chkNChangeVal(this);">
	</td>
</tr>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="21" align="center">
        <strong> ���� �ֹ� �Է� ����Ʈ(Excel or API) / TABLE : db_temp.dbo.tbl_XSite_tmporder </strong>
	</td>
</tr>
<!-- ���� �ӽ��ֹ�����Ʈ -->
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20">IDX</td>
	<td width="100">���޸�</td>
	<td width="80">��������</td>
	<td width="150">����<br>�ֹ���ȣ</td>
	<td width="60">����<br>�ֹ�����</td>
	<td width="150">����<br>���ֹ���ȣ</td>
	<td width="40">����</td>

	<td width="60">�ǸŰ�</td>
	<td width="60">��������</td>
    <td width="60">-</td>
	<td width="60">-</td>
	<td width="70"><strong>������</strong></td>
	<td width="60">����ۺ�</td>
	<td width="70">-</td>
	<td width="70">-</td>
    <td width="70">���޻�ǰ�ڵ�</td>
	<td width="80">TEN �ֹ���ȣ</td>
	<td width="100">��ǰ�ڵ�</td>
	<td width="60">�ɼ��ڵ�</td>
    <td>���ۻ���</td>
	<td>���</td>
</tr>
<% for i=0 to oCExtJungsanOrderTmp.FresultCount -1 %>
<%
    if (mapTenOrderserial="") then
        if NOT isNULL(oCExtJungsanOrderTmp.FItemList(i).FOrderSerial) and (oCExtJungsanOrderTmp.FItemList(i).FOrderSerial<>"") then
            mapTenOrderserial = oCExtJungsanOrderTmp.FItemList(i).FOrderSerial
        end if
    end if
%>
<tr align="center" bgcolor="<%=CHKIIF(oCExtJungsanOrderTmp.FItemList(i).FItemOrderCount=0,"DDDDDD","FFFFFF")%>" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
    <td><%= oCExtJungsanOrderTmp.FItemList(i).FOutMallOrderSeq %></td>
	<td><%= oCExtJungsanOrderTmp.FItemList(i).FSellSite %></td>
	<td><%= oCExtJungsanOrderTmp.FItemList(i).FSellDate %></td>
	<td><%= oCExtJungsanOrderTmp.FItemList(i).FOutMallOrderSerial %></td>
	<td><%= oCExtJungsanOrderTmp.FItemList(i).FOrgDetailKey %></td>
	<td><%= oCExtJungsanOrderTmp.FItemList(i).Fref_outmallorderserial %></td>
	<td>
        <% if oCExtJungsanOrderTmp.FItemList(i).FItemOrderCount=0 then %>
        <strong><%= oCExtJungsanOrderTmp.FItemList(i).FItemOrderCount %></strong>
        <% else %>
        <%= oCExtJungsanOrderTmp.FItemList(i).FItemOrderCount %>
        <% end if %>
    </td>
	<td align="right"><%= FormatNumber(oCExtJungsanOrderTmp.FItemList(i).Fsellprice, 0) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsanOrderTmp.FItemList(i).Fsellprice-oCExtJungsanOrderTmp.FItemList(i).Frealsellprice, 0) %></td>
    <td></td>
    <td></td>
	<td align="right"><a href="#" onClick="chgTmpOrderRealsellprice('<%= oCExtJungsanOrderTmp.FItemList(i).FOutMallOrderSeq %>',<%=oCExtJungsanOrderTmp.FItemList(i).Frealsellprice%>);return false;"><%= FormatNumber(oCExtJungsanOrderTmp.FItemList(i).Frealsellprice, 0) %></a></td>
	<td align="right"><%= FormatNumber(oCExtJungsanOrderTmp.FItemList(i).ForderDlvPay, 0) %></td>

	<td></td>
	<td></td>
    <td><%= oCExtJungsanOrderTmp.FItemList(i).FoutMallGoodsNo %></td>
	<td><%= oCExtJungsanOrderTmp.FItemList(i).FOrderSerial %></td>
	<td><%= oCExtJungsanOrderTmp.FItemList(i).FMatchitemid %></td>
	<td><a href="#" onClick="chgTmpOrderMatchitemOption('<%= oCExtJungsanOrderTmp.FItemList(i).FOutMallOrderSeq %>','<%=oCExtJungsanOrderTmp.FItemList(i).FMatchitemoption%>');return false;"><%= oCExtJungsanOrderTmp.FItemList(i).FMatchitemoption %></a></td>
	<td><%= oCExtJungsanOrderTmp.FItemList(i).Fsendstate %></td>
    <td></td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="21" align="center">
        <strong> ���� ���� ����Ʈ(Excel or API) / TABLE : db_jungsan.dbo.tbl_xSite_Jungsandata </strong>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20"><input type="checkbox" name="chkALL" onClick="switchCheckBox(this);"></td>
	<td width="100">���޸�</td>
	<td width="80">��������</td>
	<td width="150">����<br>�ֹ���ȣ</td>
	<td width="60">����<br>�ֹ�����</td>
	<td width="150">����<br>���ֹ���ȣ</td>
	<td width="40">����</td>

	<td width="60">�ǸŰ�</td>
	<td width="60">���޺δ�<br>����</td>
	<td width="60">���ٺδ�<br>����</td>
	<td width="60">������</td>
	<td width="70"><b>����ݾ�</b></td>
	<td width="60">������</td>
	<td width="70">����ݾ�</td>
	<td width="70">��������</td>
    <td width="70">���޻�ǰ�ڵ�</td>
	<td width="80">TEN �ֹ���ȣ</td>
	<td width="100">��ǰ�ڵ�</td>
	<td width="60">�ɼ��ڵ�</td>
    <td width="80">��ǰ�ֹ���ȣ</td>
	<td>���</td>
</tr>

<% for i=0 to oCExtJungsan.FresultCount -1 %>
<%
if NOT isNULL(oCExtJungsan.FItemList(i).FOrgOrderserial) and (oCExtJungsan.FItemList(i).FOrgOrderserial<>"") then
    if (mapTenOrderserial="") then
        mapTenOrderserial = oCExtJungsan.FItemList(i).FOrgOrderserial
    elseif (mapTenOrderserial<>oCExtJungsan.FItemList(i).FOrgOrderserial) then
        mapTenOrderserial2 = oCExtJungsan.FItemList(i).FOrgOrderserial
    end if
end if

' if Application("Svr_Info")="Dev" then
'     if (mapTenOrderserial="") then mapTenOrderserial="19062490802"
' end if
%>
<input type="hidden" name="sellsiteArr" value="<%= oCExtJungsan.FItemList(i).FsellSite %>">
<input type="hidden" name="extOrderserialArr" value="<%= oCExtJungsan.FItemList(i).FextOrderserial %>">
<input type="hidden" name="extOrderserSeqArr" value="<%= oCExtJungsan.FItemList(i).FextOrderserSeq %>">
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
    <td><input type="checkbox" name="chkix" value="<%=i%>" onClick="AnCheckClick(this);" ></td>
	<td><%= oCExtJungsan.FItemList(i).GetSellSiteName %></td>
	<td><%= oCExtJungsan.FItemList(i).FextMeachulDate %></td>
	<td>
        <a href="#" onClick="popByExtorderserial('<%= oCExtJungsan.FItemList(i).FextOrderserial %>');return false;"><%= oCExtJungsan.FItemList(i).FextOrderserial %></a>
        <%
            If oCExtJungsan.FItemList(i).FsellSite = "interpark" and tmpcssearch = "Y" Then 
                response.write "</br><font color='blue'>" & getCsOrgOrderserila(oCExtJungsan.FItemList(i).FextOrderserial) & "</font>"
            End If
        %>
    </td>
	<td><%= oCExtJungsan.FItemList(i).FextOrderserSeq %></td>
    <% if (oCExtJungsan.FItemList(i).FextItemNo<0) or (oCExtJungsan.FItemList(i).FsellSite="gseshop") then %>
    <td>
        <a href="#" onClick="jsOrgOrderSerialInput('<%=oCExtJungsan.FItemList(i).Fsellsite%>','<%=oCExtJungsan.FItemList(i).FextOrderserial%>','<%=oCExtJungsan.FItemList(i).FextOrderserSeq%>','<%=NULL2Blank(oCExtJungsan.FItemList(i).FextOrgOrderserial)%>')">
        <% if NULL2Blank(oCExtJungsan.FItemList(i).FextOrgOrderserial)="" then %>
        &nbsp;&nbsp;&nbsp;
        <% else %>
        <%= oCExtJungsan.FItemList(i).FextOrgOrderserial %>
        <% end if %>
        </a>
    </td>
    <% else %>
	<td><%= oCExtJungsan.FItemList(i).FextOrgOrderserial %></td>
    <% end if %>
	<td>
    <% if (oCExtJungsan.FItemList(i).FextItemNo>1) or (oCExtJungsan.FItemList(i).FextItemNo<-1) then %>
        <a href="#" onClick="jsSliceItemno('<%=oCExtJungsan.FItemList(i).Fsellsite%>','<%=oCExtJungsan.FItemList(i).FextOrderserial%>','<%=oCExtJungsan.FItemList(i).FextOrderserSeq%>',<%= oCExtJungsan.FItemList(i).FextItemNo %>);return false;"><%= oCExtJungsan.FItemList(i).FextItemNo %></a>
    <% else %>
        <%= oCExtJungsan.FItemList(i).FextItemNo %>
    <% end if %>
    </td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextItemCost, FormatDotNo) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextOwnCouponPrice, FormatDotNo) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextTenCouponPrice, FormatDotNo) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextReducedPrice, FormatDotNo) %></td>
	<td align="right"><b><%= FormatNumber(oCExtJungsan.FItemList(i).FextTenMeachulPrice, FormatDotNo) %></b>
	<% if (oCExtJungsan.FItemList(i).GetDiffMeachulPrice<>0) then %>
		<br>(<font color="red"><%=formatNumber(oCExtJungsan.FItemList(i).GetDiffMeachulPrice,FormatDotNo)%></font>)
	<% end if %>
	</td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextCommPrice, FormatDotNo) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextTenJungsanPrice, FormatDotNo) %>
	<% if (oCExtJungsan.FItemList(i).GetDiffJungsanPrice<>0) then %>
		<br>(<font color="red"><%=formatNumber(oCExtJungsan.FItemList(i).GetDiffJungsanPrice,FormatDotNo)%></font>)
	<% end if %>
	</td>
	<td>
		<%=oCExtJungsan.FItemList(i).GetSusumargin%>
	</td>
    <td><%=oCExtJungsan.FItemList(i).FExtitemid%></td>
	<td><input type="text" name="OrgOrderserialArr" value="<%= oCExtJungsan.FItemList(i).FOrgOrderserial %>" size="11" maxlength="11" onChange="chgCompValChk(this,<%= i %>);"></td>
	<td><input type="text" name="itemidArr" value="<%= oCExtJungsan.FItemList(i).Fitemid %>" size="6" maxlength="9" onChange="chgCompValChk(this,<%= i %>);"></td>
	<td><input type="text" name="itemoptionArr" value="<%= oCExtJungsan.FItemList(i).Fitemoption %>" size="4" maxlength="4" onChange="chgCompValChk(this,<%= i %>);"></td>
	<td>
		<% if NOT isNULL(oCExtJungsan.FItemList(i).FMinusOrderserial) then %>
			<%= oCExtJungsan.FItemList(i).FMinusOrderserial %>
		<% end if %>
        <% if NOT isNULL(oCExtJungsan.FItemList(i).Fref_Slice_extOrderserSeq) then %>
			<br>(<%= oCExtJungsan.FItemList(i).Fref_Slice_extOrderserSeq %>)
		<% end if %>

        <% if oCExtJungsan.FItemList(i).FExtjungsanType="D" then %>
        ��ۺ�
        <% end if %>
	</td>
    <td>
        <% if NOT isNULL(oCExtJungsan.FItemList(i).FOrgOrderserial) then %>
        <a href="#" onClick="copyOserial(<%=i%>);return false;">v</a>
        <div style="display:none"><input type="checkbox" name="cpchhk" value="<%=i%>"></div>
        <% else %>
        <input type="checkbox" name="cpchhk" value="<%=i%>">
        <% end if %>

        <% if (oCExtJungsan.FItemList(i).GetDiffReducedPrice <> 0) then %>
		<br><%= oCExtJungsan.FItemList(i).GetDiffReducedPrice %>
		<% end if %>
    </td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="21" align="center">

	</td>
</tr>
</form>
</table>
<p>
<%
if (mapTenOrderserial="") and (searchfield="OrgOrderserial") and (searchtext<>"") then
    mapTenOrderserial = searchtext
end if

dim oJungsanCheckOrder
SET oJungsanCheckOrder = New CExtJungsan
oJungsanCheckOrder.FRectOrderserial = mapTenOrderserial
if (mapTenOrderserial<>"") then
    oJungsanCheckOrder.getOutJungsanCheckOrderInfo()
end if

%>
<p  >
<% if (oJungsanCheckOrder.FResultCount>0) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">

		�ֹ����� �ֹ���ȣ : <a href="#" onClick="popDeliveryTrackingSummaryOne(<%= mapTenOrderserial %>,'',<%= 0 %>);return false;"><%= mapTenOrderserial %></a> / <%= GetUsernameWithAsterisk(oJungsanCheckOrder.FItemList(0).Fbuyname,true) %> / <%= GetUsernameWithAsterisk(oJungsanCheckOrder.FItemList(0).Freqname,true) %> / <%=oJungsanCheckOrder.FItemList(0).FreqZipAddr %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="60">��ǰ�ڵ�</td>
    <td width="60">�ɼ��ڵ�</td>
    <td width="80">�귣��ID</td>
    <td width="30">D</td>
    <td width="140">��ǰ��[�ɼ�]</td>
    <td width="40">����</td>
    <td width="70">�����Ѿ�</td>
    <td width="70"><strong>�����Ѿ�</strong></td>
    <td width="70">���Ծ�</td>
    <td width="50">����<br>����</td>
    <td width="90">�����</td>
    <td width="90">�����</td>
    <td width="90">������</td>
    <td width="110">�ù��</td>
    <td width="110">�����ȣ</td>
    <td width="100">���</td>
</tr>
<% for i=0 to oJungsanCheckOrder.FResultCount-1 %>
<tr align="center" bgcolor="<%=CHKIIF(oJungsanCheckOrder.FItemList(i).FDCancelyn="Y","#DDDDDD","#FFFFFF")%>">
    <td><%=oJungsanCheckOrder.FItemList(i).FItemid %></td>
    <td><%=oJungsanCheckOrder.FItemList(i).FItemOption %></td>
    <td><%=oJungsanCheckOrder.FItemList(i).FMakerid %></td>
    <td>
        <% if oJungsanCheckOrder.FItemList(i).FCancelyn<>"N" then response.write "<strong>"&oJungsanCheckOrder.FItemList(i).FCancelyn&"</strong>" %>
        /
        <% if oJungsanCheckOrder.FItemList(i).FDCancelyn<>"N" then response.write "<strong>"&oJungsanCheckOrder.FItemList(i).FDCancelyn&"</strong>" %>
    </td>
    <td align="left">
        <%=oJungsanCheckOrder.FItemList(i).FItemname %>
        <%
        if (oJungsanCheckOrder.FItemList(i).FItemOptionname<>"") then
            response.write " <font color=blue>["&oJungsanCheckOrder.FItemList(i).FItemOptionname&"]</font>"
        end if
        %>
    </td>
    <td><%=oJungsanCheckOrder.FItemList(i).FItemNo %></td>
    <td align="right"><%=FormatNumber(oJungsanCheckOrder.FItemList(i).FItemCost,0) %></td>
    <td align="right"><%=FormatNumber(oJungsanCheckOrder.FItemList(i).FReducedprice,0) %></td>
    <td align="right"><%=FormatNumber(oJungsanCheckOrder.FItemList(i).FBuycash,0) %></td>
    <td><%=oJungsanCheckOrder.FItemList(i).FoMwDiv %></td>
    <td><%=oJungsanCheckOrder.FItemList(i).FBeasongdate %></td>
    <td><%=oJungsanCheckOrder.FItemList(i).Fdlvfinishdt %></td>
    <td><%=oJungsanCheckOrder.FItemList(i).Fjungsanfixdate %></td>
    <td><%=getSongjangDiv2Val(oJungsanCheckOrder.FItemList(i).FSongjangDiv,1) %></td>
    <td><a target="_dlv2" href="<%=getTrackNaverURIByTrName(oJungsanCheckOrder.FItemList(i).Fsongjangdiv,oJungsanCheckOrder.FItemList(i).Fsongjangno)%>"><%=oJungsanCheckOrder.FItemList(i).Fsongjangno %></a></td>
    <td>
    <% if (oJungsanCheckOrder.FItemList(i).Fitemid<>0 and oJungsanCheckOrder.FItemList(i).Fitemid<>100) then %>
    <a href="#" onClick="popJcomment('<%=oJungsanCheckOrder.FItemList(i).FOrderserial%>','<%=oJungsanCheckOrder.FItemList(i).Fitemid%>','<%=oJungsanCheckOrder.FItemList(i).Fitemoption%>');return false;">
    <%=CHKIIF(isNULL(oJungsanCheckOrder.FItemList(i).Fcomment),"<img src='/images/icon_new.gif' alt='�ڸ�Ʈ�ۼ�'>",oJungsanCheckOrder.FItemList(i).Fcomment)%>
    </a>
    <% end if %>
    </td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16" align="center">

	</td>
</tr>
</table>
<% end if %>

<%

if mapTenOrderserial2<>"" then

    oJungsanCheckOrder.FRectOrderserial = mapTenOrderserial2
    if (mapTenOrderserial2<>"") then
        oJungsanCheckOrder.getOutJungsanCheckOrderInfo()
    end if

%>
<p  >
<% if (oJungsanCheckOrder.FResultCount>0) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
		�ֹ����� �ֹ���ȣ : <%= mapTenOrderserial2 %> / <%= GetUsernameWithAsterisk(oJungsanCheckOrder.FItemList(0).Fbuyname,true) %> / <%= GetUsernameWithAsterisk(oJungsanCheckOrder.FItemList(0).Freqname,true) %> / <%=oJungsanCheckOrder.FItemList(0).FreqZipAddr %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="60">��ǰ�ڵ�</td>
    <td width="60">�ɼ��ڵ�</td>
    <td width="80">�귣��ID</td>
    <td width="30">D</td>
    <td width="140">��ǰ��[�ɼ�]</td>
    <td width="40">����</td>
    <td width="70">�����Ѿ�</td>
    <td width="70"><strong>�����Ѿ�</strong></td>
    <td width="70">���Ծ�</td>
    <td width="50">����<br>����</td>
    <td width="90">�����</td>
    <td width="90">�����</td>
    <td width="90">������</td>
    <td width="110">�ù��</td>
    <td width="110">�����ȣ</td>
    <td width="100">���</td>
</tr>
<% for i=0 to oJungsanCheckOrder.FResultCount-1 %>
<tr align="center" bgcolor="<%=CHKIIF(oJungsanCheckOrder.FItemList(i).FDCancelyn="Y","#DDDDDD","#FFFFFF")%>">
    <td><%=oJungsanCheckOrder.FItemList(i).FItemid %></td>
    <td><%=oJungsanCheckOrder.FItemList(i).FItemOption %></td>
    <td><%=oJungsanCheckOrder.FItemList(i).FMakerid %></td>
    <td>
        <% if oJungsanCheckOrder.FItemList(i).FCancelyn<>"N" then response.write "<strong>"&oJungsanCheckOrder.FItemList(i).FCancelyn&"</strong>" %>
        /
        <% if oJungsanCheckOrder.FItemList(i).FDCancelyn<>"N" then response.write "<strong>"&oJungsanCheckOrder.FItemList(i).FDCancelyn&"</strong>" %>
    </td>
    <td align="left">
        <%=oJungsanCheckOrder.FItemList(i).FItemname %>
        <%
        if (oJungsanCheckOrder.FItemList(i).FItemOptionname<>"") then
            response.write " <font color=blue>["&oJungsanCheckOrder.FItemList(i).FItemOptionname&"]</font>"
        end if
        %>
    </td>
    <td><%=oJungsanCheckOrder.FItemList(i).FItemNo %></td>
    <td align="right"><%=FormatNumber(oJungsanCheckOrder.FItemList(i).FItemCost,0) %></td>
    <td align="right"><%=FormatNumber(oJungsanCheckOrder.FItemList(i).FReducedprice,0) %></td>
    <td align="right"><%=FormatNumber(oJungsanCheckOrder.FItemList(i).FBuycash,0) %></td>
    <td><%=oJungsanCheckOrder.FItemList(i).FoMwDiv %></td>
    <td><%=oJungsanCheckOrder.FItemList(i).FBeasongdate %></td>
    <td><%=oJungsanCheckOrder.FItemList(i).Fdlvfinishdt %></td>
    <td><%=oJungsanCheckOrder.FItemList(i).Fjungsanfixdate %></td>
    <td><%=getSongjangDiv2Val(oJungsanCheckOrder.FItemList(i).FSongjangDiv,1) %></td>
    <td><a target="_dlv2" href="<%=getTrackNaverURIByTrName(oJungsanCheckOrder.FItemList(i).Fsongjangdiv,oJungsanCheckOrder.FItemList(i).Fsongjangno)%>"><%=oJungsanCheckOrder.FItemList(i).Fsongjangno %></a></td>
    <td>
    <% if (oJungsanCheckOrder.FItemList(i).Fitemid<>0 and oJungsanCheckOrder.FItemList(i).Fitemid<>100) then %>
    <a href="#" onClick="popJcomment('<%=oJungsanCheckOrder.FItemList(i).FOrderserial%>','<%=oJungsanCheckOrder.FItemList(i).Fitemid%>','<%=oJungsanCheckOrder.FItemList(i).Fitemoption%>');return false;"><img src="/images/icon_new.gif" alt="�ڸ�Ʈ�ۼ�"></a>
    <% end if %>
    </td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16" align="center">

	</td>
</tr>
</table>
<% end if %>
<% end if %>

<% SET oJungsanCheckOrder = Nothing %>

<p>

<%
'' CS����
dim oJungsanCheckCS
SET oJungsanCheckCS = New CExtJungsan
oJungsanCheckCS.FRectOrderserial = mapTenOrderserial
if (mapTenOrderserial<>"") then
    oJungsanCheckCS.getOutJungsanCheckCSInfo()
end if

%>
<% if (oJungsanCheckCS.FResultCount>0) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="13">
		CS���� �ֹ���ȣ : <%= mapTenOrderserial %>

        &nbsp;<input type="button" class="button" value="����CS <%=oJungsanCheckCS.FResultCount%>��" class="csbutton" style="width:90px;" onclick="popcenter_Action_List('<%= mapTenOrderserial %>','','');">
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="60">csID</td>
    <td width="60">����</td>
    <td width="80">�귣��ID</td>
    <td width="30">D</td>
    <td width="140">TITLE</td>
    <td width="40">����</td>
    <td width="70">������</td>
    <td width="70">�Ϸ���</td>
    <td width="70">Ȯ����</td>
    <td width="70">���(����)��</td>

    <td width="90">����CsID</td>
    <td width="90">�����ֹ���ȣ</td>
    <td width="100">���</td>
</tr>
<% for i=0 to oJungsanCheckCS.FResultCount-1 %>
<%
if NOT isNULL(oJungsanCheckCS.FItemList(i).getRefOrderSerial) and (oJungsanCheckCS.FItemList(i).getRefOrderSerial<>"") then
    mapRtnTenOrderserial = oJungsanCheckCS.FItemList(i).getRefOrderSerial
end if

' if Application("Svr_Info")="Dev" then
'     if (mapRtnTenOrderserial="") then mapRtnTenOrderserial="19040190697"
' end if
%>
<tr align="center" bgcolor="<%=CHKIIF(oJungsanCheckCS.FItemList(i).Fdeleteyn="Y","#DDDDDD","#FFFFFF")%>">
    <td><%=oJungsanCheckCS.FItemList(i).FCsID %></td>
    <td><%=oJungsanCheckCS.FItemList(i).FdivName %></td>
    <td>
        <%=oJungsanCheckCS.FItemList(i).Fmakerid %>
        <% if ((oJungsanCheckCS.FItemList(i).Fmakerid<>"") and (oJungsanCheckCS.FItemList(i).Frequireupche<>"Y")) or ((oJungsanCheckCS.FItemList(i).Fmakerid="") and (oJungsanCheckCS.FItemList(i).Frequireupche="Y")) then %>
        <br>(<%=oJungsanCheckCS.FItemList(i).Frequireupche%>)
        <% end if %>
    </td>
    <td>
        <% if oJungsanCheckCS.FItemList(i).Fdeleteyn<>"N" then response.write "<strong>"&oJungsanCheckCS.FItemList(i).Fdeleteyn&"</strong>" %>
    </td>
    <td align="left"><%=oJungsanCheckCS.FItemList(i).Ftitle %></td>
    <td><%=oJungsanCheckCS.FItemList(i).getCsStateName %> (<%=oJungsanCheckCS.FItemList(i).Fcurrstate%>)</td>
    <td><%=oJungsanCheckCS.FItemList(i).Fregdate %></td>
    <td><%=oJungsanCheckCS.FItemList(i).Ffinishdate %></td>
    <td><%=oJungsanCheckCS.FItemList(i).Fconfirmdate %></td>
    <td><%=oJungsanCheckCS.FItemList(i).Fdeletedate %></td>
    <td><%=oJungsanCheckCS.FItemList(i).Frefasid %></td>
    <td><%=oJungsanCheckCS.FItemList(i).getRefOrderSerial %></td>
    <td></td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="13" align="center">

	</td>
</tr>
</table>
<% end if %>
<% SET oJungsanCheckCS = Nothing %>
<p>

<%
'' ��ǰ/��ȯ�ֹ���.
dim oJungsanCheckOrderRtn
SET oJungsanCheckOrderRtn = New CExtJungsan
oJungsanCheckOrderRtn.FRectOrderserial = mapRtnTenOrderserial
if (mapRtnTenOrderserial<>"") then
    oJungsanCheckOrderRtn.getOutJungsanCheckOrderInfo()
end if

%>
<p  >
<% if (oJungsanCheckOrderRtn.FResultCount>0) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
        <% if oJungsanCheckOrderRtn.FItemList(0).Fjumundiv="9" then %>
        <strong>��ǰ</strong>/��ȯ
        <% elseif oJungsanCheckOrderRtn.FItemList(0).Fjumundiv="6" then %>
        ��ǰ/<strong>��ȯ</strong>
        <% else %>
        ��ǰ/��ȯ
        <% end if %>
		 �ֹ����� �ֹ���ȣ : <%= mapRtnTenOrderserial %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="60">��ǰ�ڵ�</td>
    <td width="60">�ɼ��ڵ�</td>
    <td width="80">�귣��ID</td>
    <td width="30">D</td>
    <td width="140">��ǰ��[�ɼ�]</td>
    <td width="40">����</td>
    <td width="70">�����Ѿ�</td>
    <td width="70"><strong>�����Ѿ�</strong></td>
    <td width="70">���Ծ�</td>
    <td width="50">����<br>����</td>
    <td width="90">�����</td>
    <td width="90">�����</td>
    <td width="90">������</td>
    <td width="110">�ù��</td>
    <td width="110">�����ȣ</td>
    <td width="100">���</td>
</tr>
<% for i=0 to oJungsanCheckOrderRtn.FResultCount-1 %>
<tr align="center" bgcolor="<%=CHKIIF(oJungsanCheckOrderRtn.FItemList(i).FDCancelyn="Y","#DDDDDD","#FFFFFF")%>">
    <td><%=oJungsanCheckOrderRtn.FItemList(i).FItemid %></td>
    <td><%=oJungsanCheckOrderRtn.FItemList(i).FItemOption %></td>
    <td><%=oJungsanCheckOrderRtn.FItemList(i).FMakerid %></td>
    <td>
        <% if oJungsanCheckOrderRtn.FItemList(i).FCancelyn<>"N" then response.write "<strong>"&oJungsanCheckOrderRtn.FItemList(i).FCancelyn&"</strong>" %>
        /
        <% if oJungsanCheckOrderRtn.FItemList(i).FDCancelyn<>"N" then response.write "<strong>"&oJungsanCheckOrderRtn.FItemList(i).FDCancelyn&"</strong>" %>
    </td>
    <td align="left">
        <%=oJungsanCheckOrderRtn.FItemList(i).FItemname %>
        <%
        if (oJungsanCheckOrderRtn.FItemList(i).FItemOptionname<>"") then
            response.write " <font color=blue>["&oJungsanCheckOrderRtn.FItemList(i).FItemOptionname&"]</font>"
        end if
        %>
    </td>
    <td><%=oJungsanCheckOrderRtn.FItemList(i).FItemNo %></td>
    <td align="right"><%=FormatNumber(oJungsanCheckOrderRtn.FItemList(i).FItemCost,0) %></td>
    <td align="right"><%=FormatNumber(oJungsanCheckOrderRtn.FItemList(i).FReducedprice,0) %></td>
    <td align="right"><%=FormatNumber(oJungsanCheckOrderRtn.FItemList(i).FBuycash,0) %></td>
    <td><%=oJungsanCheckOrderRtn.FItemList(i).FoMwDiv %></td>
    <td><%=oJungsanCheckOrderRtn.FItemList(i).FBeasongdate %></td>
    <td><%=oJungsanCheckOrderRtn.FItemList(i).Fdlvfinishdt %></td>
    <td><%=oJungsanCheckOrderRtn.FItemList(i).Fjungsanfixdate %></td>
    <td><%=getSongjangDiv2Val(oJungsanCheckOrderRtn.FItemList(i).FSongjangDiv,1) %></td>
    <td><%=oJungsanCheckOrderRtn.FItemList(i).Fsongjangno %></td>
    <td></td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16" align="center">

	</td>
</tr>
</table>
<% end if %>
<% SET oJungsanCheckOrderRtn = Nothing %>

<%
'' ���� ����α� by �ֹ���ȣ
dim oSongjangChgLog
SET oSongjangChgLog = new CDeliveryTrack
oSongjangChgLog.FRectOrderserial = mapTenOrderserial
if (mapTenOrderserial<>"") then
    oSongjangChgLog.getSongjangChangeLogList()
end if
%>
<p  >
<% if (oSongjangChgLog.FResultCount>0) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16">
        ���庯��α� �ֹ���ȣ : <%= mapTenOrderserial %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="60">LogIdx</td>
    <td width="60">��ǰ�ڵ�</td>
    <td width="60">�ɼ��ڵ�</td>
    <td width="110">�����ù��</td>
    <td width="110">���������ȣ</td>
    <td width="110">�����ù��</td>
    <td width="110">��������ȣ</td>

    <td width="80">������</td>
    <td width="70">�����</td>
    <td width="70">���汸��</td>

    <td width="70">�����ù��</td>
    <td width="50">��������ȣ</td>
    <td width="90">�����</td>
    <td width="90">�����</td>
    <td width="90">������</td>

    <td width="100">���</td>
</tr>
<% for i=0 to oSongjangChgLog.FResultCount-1 %>
<tr align="center" bgcolor="#FFFFFF">
    <td><%=oSongjangChgLog.FItemList(i).Fsongjangchgidx %></td>
    <td><%=oSongjangChgLog.FItemList(i).FItemid %></td>
    <td><%=oSongjangChgLog.FItemList(i).FItemOption %></td>
    <td><%=getSongjangDiv2Val(oSongjangChgLog.FItemList(i).Fpsongjangdiv,1) %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fpsongjangno %></td>
    <td><%=getSongjangDiv2Val(oSongjangChgLog.FItemList(i).Fchgsongjangdiv,1) %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fchgsongjangno %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fchguserid %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fregdt %></td>
    <td><%=oSongjangChgLog.FItemList(i).FactionType %></td>
    <td><%=getSongjangDiv2Val(oSongjangChgLog.FItemList(i).Fsongjangdiv,1) %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fsongjangno %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fbeasongdate %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fdlvfinishdt %></td>
    <td><%=oSongjangChgLog.FItemList(i).Fjungsanfixdate %></td>
    <td></td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="16" align="center">

	</td>
</tr>
</table>
<% end if %>

<% SET oSongjangChgLog = Nothing %>

<%
'' �������� �ڸ�Ʈ �α�
dim oExtJungsanCom
SET oExtJungsanCom = new CExtJungsan
oExtJungsanCom.FRectOrderserial = mapTenOrderserial
if (mapTenOrderserial<>"") then
    oExtJungsanCom.getExtjungsanCommentList()
end if
%>
<p  >
<% if (oExtJungsanCom.FResultCount>0) then %>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="#FFFFFF">
	<td colspan="9">
        ��������Comment �ֹ���ȣ : <%= mapTenOrderserial %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="60">LogIdx</td>
    <td width="110">�ֹ���ȣ</td>
    <td width="80">��ǰ�ڵ�</td>
    <td width="80">�ɼ��ڵ�</td>

    <td width="90">�����</td>
    <td >����</td>

    <td width="120">�����</td>
    <td width="120">������</td>

    <td width="80">���</td>
</tr>
<% for i=0 to oExtJungsanCom.FResultCount-1 %>
<tr align="center" bgcolor="<%=CHKIIF(isNULL(oExtJungsanCom.FItemList(i).Fdeldate),"#FFFFFF","#CCCCCC")%>">
    <td><%=oExtJungsanCom.FItemList(i).Frowidx %></td>
    <td><%=oExtJungsanCom.FItemList(i).Forderserial %></td>
    <td><%=oExtJungsanCom.FItemList(i).FItemid %></td>
    <td><%=oExtJungsanCom.FItemList(i).FItemOption %></td>
    <td><%=oExtJungsanCom.FItemList(i).Freguserid %></td>
    <td><%=oExtJungsanCom.FItemList(i).Fcomment %></td>
    <td><%=oExtJungsanCom.FItemList(i).Fregdate %></td>
    <td><%=oExtJungsanCom.FItemList(i).Fdeldate %></td>
    <td>
        <% if isNULL(oExtJungsanCom.FItemList(i).Fdeldate) then %>
        <a href="#" onClick="delJcomment('<%=oExtJungsanCom.FItemList(i).Frowidx %>');return false;">[X]</a>
        <% end if %>
    </td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9" align="center">

	</td>
</tr>
</table>
<% end if %>

<% SET oExtJungsanCom = Nothing %>

<p>
<form name="extEdtFrm" method="post" action="extJungsan_process.asp">
<input type="hidden" name="mode" value="extorgorderserialedit">
<input type="hidden" name="sellsite" value="">
<input type="hidden" name="extOrderserial" value="">
<input type="hidden" name="extOrderserSeq" value="">
<input type="hidden" name="extorgorderserial" value="">
</form>
<form name="slicefrm" method="post" action="extJungsan_process.asp">
<input type="hidden" name="mode" value="slicejitemno">
<input type="hidden" name="sellsite" value="">
<input type="hidden" name="extOrderserial" value="">
<input type="hidden" name="extOrderserSeq" value="">
<input type="hidden" name="newSliceNo" value="">
</form>
<form name="frmcmt" method="post" action="extJungsan_process.asp">
<input type="hidden" name="mode" value="addcmt">
<input type="hidden" name="orderserial" value="">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="itemoption" value="">
<input type="hidden" name="addcomment" value="">
<input type="hidden" name="rowidx" value="">
</form>
<form name="frmXsiteTmpval" method="post" action="extJungsan_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="outmallorderseq" value="">
<input type="hidden" name="chgval" value="">
</form>

<%
SET oCExtJungsanOrderTmp = Nothing
set oCExtJungsan = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->