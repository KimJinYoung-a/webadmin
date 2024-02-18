<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_refundcls.asp" -->
<%
dim upfiledate
dim sitegubun

upfiledate  = request("upfiledate")
sitegubun      	= RequestCheckVar(request("sitegubun"),32)

dim OrefundList
set OrefundList = new CCSRefund
OrefundList.FCurrPage           = 1
OrefundList.FPageSize           = 5000

if upfiledate="" then
    OrefundList.FRectCurrstate      = "B001"
else

end if

OrefundList.FRectReturnmethod   = "R007"
OrefundList.FRectUploadState    = "uploaded"
OrefundList.FRectUpfiledate     = upfiledate

if (sitegubun = "academy") then
	OrefundList.GetRefundRequireAcademyList
else
	'10x10
	OrefundList.GetRefundRequireList
end if

dim i, refundsum

dim ISIBKREFUND : ISIBKREFUND = false
dim RefundSuccCnt : RefundSuccCnt=0
dim Is90ProOverRefund : Is90ProOverRefund = false


for i=0 to OrefundList.FREsultCount-1
    if (OrefundList.FItemList(i).IsIBKRefund) then
        ISIBKREFUND = true

        if (OrefundList.FItemList(i).IsIBKRefund and  OrefundList.FItemList(i).FIBK_PROC_YN="Y") then
            RefundSuccCnt = RefundSuccCnt + 1
        end if
        ''Exit for
    end if
next

if (OrefundList.FREsultCount>0) then
    Is90ProOverRefund = (RefundSuccCnt/OrefundList.FREsultCount*100>90)
end if
%>

<script language='javascript'>
function popXl(){
    //var popwin = window.open('poprefundfile.asp?xl=on&upfiledate=<%= upfiledate %>','popXl','');
    //popwin.focus();
}

function popCSV(){
    var popwin = window.open('poprefundfile_CSV.asp?upfiledate=<%= upfiledate %>','popCSV','');
    popwin.focus();
}

function popTXT(){
    var popwin = window.open('poprefundfile_TXT.asp?upfiledate=<%= upfiledate %>','popTxt','');
    popwin.focus();
}

function FinishRefundIBK(frm){
    if (confirm('��ü �Ϸ�� ������ ���� ȯ�� �Ϸ� ó�� �Ͻðڽ��ϱ�?\n\n�Ϸ�ó���� �ڵ����� ���ڸ޼����� �߼۵˴ϴ�.\n\n([�ٹ�����(�ΰŽ� ��ī����)] ���� 000,000 �� ȯ���� �Ϸ�Ǿ����ϴ�. ��ſ� �Ϸ� �Ǽ���.)')){
		frm.submit();
	}
}

function FinishRefund(frm){
    var chk = 0;

	if(frm.ckidx.length>1){
		for(i=0;i<frm.ckidx.length;i++){
			if(frm.ckidx[i].checked)
				chk++;
		}
	}else{
		if(frm.ckidx.checked)
			chk++;
	}

	if(chk==0){
		alert("�Ϸ��� ������ �������ֽʽÿ�.");
		return false;
	}else{
	    if (confirm('���� �Ͻ� ������ ȯ�� �Ϸ� ó�� �Ͻðڽ��ϱ�?\n\n�Ϸ�ó���� �ڵ����� ���ڸ޼����� �߼۵˴ϴ�.\n\n([�ٹ�����] ���� 000,000 �� ȯ���� �Ϸ�Ǿ����ϴ�. ��ſ� �Ϸ� �Ǽ���.)')){
			frm.submit();
		}
	}
}

function regConfirmMsg(iid,fin){
    var frm = document.frm_list;
    var sitegubun = frm.sitegubun.value;

    var popwin = window.open('/cscenter/action/pop_ConfirmMsg.asp?sitegubun=' + sitegubun + '&id=' + iid + '&fin=' + fin,'regConfirmMsg','width=600,height=400,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function switchCheckBox(){
	var frm=document.frm_list;
    var swchecked = frm.switchCheck.checked;

	if(frm.ckidx.length>1){
		for(i=0;i<frm.ckidx.length;i++){
		    if (!frm.ckidx[i].disabled){
    		    frm.ckidx[i].checked=swchecked;

    		    checkRow(frm.ckidx[i]);
		    }
		}
	}else{
	    if (!frm.ckidx.disabled){
    		frm.ckidx.checked=swchecked;
    	    checkRow(frm.ckidx);
    	}
	}
}

function checkRow(comp){
    if (comp.checked){
        hL(comp);
    }else{
        dL(comp);
    }
}

</script>

<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" >
 <% if (ISIBKREFUND) then %>
 <tr bgcolor="#FFFFFF">
    <td align="left">
	    ����Ʈ : <%= sitegubun %>
	    &nbsp;&nbsp;
	    |
	    &nbsp;&nbsp;
	    ���� �ۼ��� [<%= upfiledate %>]
    </td>
    <td align="right">
        <input class="button" type="button" value="�Ϸ� ó��" onClick="FinishRefundIBK(frmFinish);" onFocus="this.blur();">
    </td>
 </tr>
 <% else %>
 <tr bgcolor="#FFFFFF">
    <td align="left">���� �ۼ��� [<%= upfiledate %>]
        <input class="button" type="button" value="���� ���� ȯ�ҿϷ� ó��" onClick="FinishRefund(frm_list);" onFocus="this.blur();">
    </td>
    <td align="right">
        <input class="button" type="button" value="��ó�� ���� TXT ���� �ޱ�" onClick="popTXT();">
        <input class="button" type="button" value="��ó�� ���� CSV ���� �ޱ�" onClick="popCSV();">
    </td>
 </tr>
 <% end if %>
</table>

<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="#BABABA">
<form name="frm_list" method="post" action="refundlist_process.asp">
<input type="hidden" name="sitegubun" value="<%= sitegubun %>">
<input type="hidden" name="mode" value="finisharray">
 <tr align="center" bgcolor="#E6E6E6">
    <td width="30"><input type="checkbox" name="switchCheck" onClick="switchCheckBox()"></td>
    <td width="60">����</td>
    <td>����</td>
    <td width="80">ȯ�ұݾ�</td>
    <td width="120">������</td>
    <td width="80">CSó������</td>
    <td width="80">IBK����</td>
    <td width="80">Ȯ�ο�û</td>
    <td width="80">�Ϸ�ó��</td>
 </tr>
 <% for i=0 to OrefundList.FResultCount - 1 %>
 <%
    refundsum = refundsum + OrefundList.FItemList(i).Frefundrequire
    if (OrefundList.FItemList(i).Fencmethod = "TBT") then
	    ''��� ����.
		OrefundList.FItemList(i).Frebankaccount = TBTDecrypt(OrefundList.FItemList(i).FencAccount)
    elseif (OrefundList.FItemList(i).Fencmethod = "PH1") then
        OrefundList.FItemList(i).Frebankaccount = OrefundList.FItemList(i).Fdecaccount
    elseif (OrefundList.FItemList(i).Fencmethod = "AE2") then
        OrefundList.FItemList(i).Frebankaccount = OrefundList.FItemList(i).Fdecaccount
	end if
 %>
 <tr bgcolor="#FFFFFF">
    <td align="center" >
    <% if OrefundList.FItemList(i).FCurrstate="B001" then %>
        <input type="checkbox" name="ckidx" value="<%= OrefundList.FItemList(i).FasId %>" onClick="checkRow(this);" <%= ChkIIF(OrefundList.FItemList(i).IsIBKRefund,"disabled","") %> >
    <% end if %>
    </td>
    <td><%= OrefundList.FItemList(i).Frebankname %></td>
    <td><%= OrefundList.FItemList(i).Frebankaccount %></td>
    <td align="right"><%= FormatNumber(OrefundList.FItemList(i).Frefundrequire,0) %></td>
    <td><%= OrefundList.FItemList(i).Frebankownername %></td>
    <td align="center">
        <font color="<%= OrefundList.FItemList(i).GetCurrStateColor %>"><%= OrefundList.FItemList(i).GetCurrStateName %></font>
    </td>
    <td align="center"><%= OrefundList.FItemList(i).getIBKstateName %>
    <% if (OrefundList.FItemList(i).FIBK_ERR_MSG<>"") then %>
    <br>(<%= OrefundList.FItemList(i).FIBK_ERR_MSG %>)
    <% end if %>
    </td>
    <td align="center">
        <% if (FALSE) and (Is90ProOverRefund) and (OrefundList.FItemList(i).IsIBKRefund) and IsNULL(OrefundList.FItemList(i).FIBK_PROC_YN) then %>
        <input class="button" type="button" value="Ȯ�ο�û" onclick="regConfirmMsg('<%= OrefundList.FItemList(i).Fasid %>','');" >
        <% else %>
            <% if OrefundList.FItemList(i).FCurrstate="B001" then %>
            <input class="button" type="button" value="Ȯ�ο�û" onclick="regConfirmMsg('<%= OrefundList.FItemList(i).Fasid %>','');" <%= ChkIIF(OrefundList.FItemList(i).IsIBKRefund and (Not OrefundList.FItemList(i).IsIBKProcERR),"disabled","") %>>
            <% end if %>
        <% end if %>
    </td>
    <td align="center">
        <% if OrefundList.FItemList(i).FCurrstate="B001" then %>
        <input type="button" class="button" value="�Ϸ�ó��" onClick="PopCSActionFinish('<%= OrefundList.FItemList(i).FasId %>','finishreginfo');" onFocus="this.blur();" <%= ChkIIF(OrefundList.FItemList(i).IsIBKRefund,"disabled","") %>>
        <% end if %>
    </td>
 </tr>
 <% next %>
  <tr bgcolor="#FFFFFF">
    <td colspan="2">Total</td>
    <td></td>
    <td align="right"><%= FormatNumber(refundsum,0) %></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
  </tr>
</form>
</table>

<%
set OrefundList = Nothing
%>
<form name="frmFinish" method="post" action="refundlist_process.asp">
<input type="hidden" name="mode" value="finishfile">
<input type="hidden" name="upfiledate" value="<%= upfiledate %>">
<input type="hidden" name="sitegubun" value="<%= sitegubun %>">
</form>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->