<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/tax/eseroTaxCls.asp"-->
<!-- #include virtual="/lib/classes/jungsan/upchejungsanIcheFileCls.asp"-->
<!-- #include virtual="/admin/upchejungsan/upchejungsan_function.asp"-->
<%
Dim ipFileNo : ipFileNo = requestCheckVar(request("ipFileNo"),10)
Dim page     : page = requestCheckVar(request("page"),10)
Dim iDetailState : iDetailState = requestCheckVar(request("iDetailState"),10)
Dim ierpSendState  : ierpSendState = requestCheckVar(request("ierpSendState"),10)
Dim MappingYn : MappingYn = requestCheckVar(request("MappingYn"),10)
if (page="") then page=1

Dim clsICheFile
set clsICheFile = new CupcheJungsanIcheFile
clsICheFile.FPageSize = 50
clsICheFile.FCurrPage  = page
clsICheFile.FRectipFileNo = ipFileNo
clsICheFile.FRectisMappingYn =MappingYn
clsICheFile.FRectErpSendState = ierpSendState
clsICheFile.FRectDetailState = iDetailState
clsICheFile.fnGetIcheFileMappingList

Dim oneIcheFile
set oneIcheFile = new CupcheJungsanIcheFile
oneIcheFile.FRectipFileNo = ipFileNo
oneIcheFile.getOneIcheFileMaster

Dim isWonChonFile : isWonChonFile = oneIcheFile.FOneItem.isWonChonFile
Dim i
%>
<script language='javascript'>
function sendERPDoc(frm){
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}

function ERPInoutMapping(frm){
    if (confirm('�������� �� ��꼭 ���� �� ���� �Ͻñ� �ٶ��ϴ�.\n���ް���>����/��������Ȯ�� �ݾװ� ���並 ���Ͽ� �ڵ� ���� �ϴ� ����Դϴ�.\n����Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}

function CheckAll(comp){
    var frm = comp.form;
    var bool =comp.checked;
	for (var i=0;i<frm.elements.length;i++){
		//check optioon
		var e = frm.elements[i];

		//check itemEA
		if ((e.type=="checkbox")) {
		    if (e.disabled) continue;
			e.checked=bool;
			AnCheckClick(e)
		}
	}
}

function popTargetDetail(itargetGb,iidx,iridx){
    var popURL ='';
    if (itargetGb=="ON"){
        popURL = "/admin/upchejungsan/nowjungsanmasteredit.asp?id="+iidx;
    }else if (itargetGb=="OF"){
        popURL = "/admin/offupchejungsan/off_jungsanstateedit.asp?idx="+iidx;
    }else if (itargetGb=="9"){
        popURL = "/admin/approval/eapp/modeappPayDoc.asp?ipridx="+iidx+"&iridx="+iridx;
    }else if (itargetGb=="11"){
        popURL = "/cscenter/taxsheet/Tax_view.asp?taxIdx="+iidx;
    }
    
    var popWin = window.open(popURL,'popTargetDetail','width=900,height=600,scrollbars=yes,resizable=yes');
    popWin.focus();
}
</script>
<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td valign="top">
            ���Ϲ�ȣ : <input type="text" name="ipFileNo" value="<%= ipFileNo %>" size="4" maxlength="7" readOnly class="text_ro" >
            &nbsp;
            <!--
            ���� ���� : 
            <select name="iDetailState">
            <option value="" >��ü
            <option value="0" <%= CHKIIF(iDetailState="0","selected","") %> >�ۼ���
            <option value="7" <%= CHKIIF(iDetailState="7","selected","") %> >�ԱݿϷ�
            <option value="8" <%= CHKIIF(iDetailState="8","selected","") %> >ERP����
            <option value="M" <%= CHKIIF(iDetailState="M","selected","") %> >��������ü
            </select>
            &nbsp;
            -->
            ���� ���� : 
            <select name="ierpSendState">
            <option value="" >��ü
            <option value="Y" <%= CHKIIF(ierpSendState="Y","selected","") %> >����
            <option value="N" <%= CHKIIF(ierpSendState="N","selected","") %> >������
            </select>
            &nbsp;
            ���λ���
            <select name="MappingYn">
            <option value="" >��ü
            <option value="Y" <%= CHKIIF(MappingYn="Y","selected","") %> >����
            <option value="N" <%= CHKIIF(MappingYn="N","selected","") %> >�̸���
            </select>
        </td>
        <td valign="top" align="right">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->
<p>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr bgcolor="#FFFFFF">
        <td colspan="2">��<%= clsICheFile.FTotCnt %>�� </td>
        <td colspan="5" align="left">
        <!-- ���Ұ�
        <input type="button" value="erp ��� ��Ī" class="button" onClick="ERPInoutMapping(frmBuf);">
        -->
        </td>
        <td colspan="5" align="right">
        <input type="button" value="erp ���� ����" class="button" onClick="sendERPDoc(frmList);">
        </td>
    </tr>
    <form name="frmList" method="post" action="/admin/tax/eTax_process.asp">
    <input type="hidden" name="mode" value="sendDocErp">
    <input type="hidden" name="ipFileNo" value="<%= ipFileNo %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td width="20"><input type="checkbox" name="chkALL" value="" onClick="CheckAll(this)"></td>
		<td width="80">����</td>
		<td width="80">����</td>
		<td width="80">����ڹ�ȣ</td>
		<td width="120">����ڸ�</td>
		<% if (isWonChonFile) then %>
		<td width="100">(��õ)����ݾ�</td>
		<% else %> 
		<td width="100">����ݾ�</td>
		<% end if %>
		<td width="120">TaxKey</td>
		<td width="90">������</td>
      	<td width="150">��꼭�ݾ�</td>
		<td width="100">���޾�</td>
		<td width="100">����</td>
		<td width="100">���ۻ���</td>
	</tr>
	<% for i=0 to clsICheFile.FResultCount-1 %>
	<tr bgcolor="#FFFFFF">
	    <td><input type="checkbox" name="chkTaxKey" value="<%= clsICheFile.FItemList(i).FtaxKey %>" <%= chkIIF(IsNULL(clsICheFile.FItemList(i).FtaxKey) or ( Not IsNULL(clsICheFile.FItemList(i).FerpLinkType)),"disabled","") %> ></td>
	    <td><%= clsICheFile.FItemList(i).FtargetGbn %>
	    <% If IsNULL(clsICheFile.FItemList(i).FtaxKey) THEN %>
	    [<%= clsICheFile.FItemList(i).FtargetIdx %><img src="/images/icon_arrow_link.gif" onClick="popTargetDetail('<%= clsICheFile.FItemList(i).FtargetGbn %>','<%=clsICheFile.FItemList(i).FtargetIdx%>','')" style="cursor:pointer">]
	    <% end if %>
	    </td>
	    <td><%= clsICheFile.FItemList(i).getIpFileDetailStateName %></td>
	    <td><%= clsICheFile.FItemList(i).FSellCorpNo %></td>
	    <td><%= clsICheFile.FItemList(i).FSellCorpName %></td>
	    <td></td>
	    <td><%= clsICheFile.FItemList(i).FtaxKey %></td>
	    <% IF IsNULL(clsICheFile.FItemList(i).FtaxKey) then %>
	    <td></td>
	    <td></td>
	    <td></td>
	    <td></td>
	    <% else %>
	    <td><%= clsICheFile.FItemList(i).FAppDate %></td>
	    <td><%= clsICheFile.FItemList(i).FTotSum %></td>
	    <td><%= clsICheFile.FItemList(i).Fsuplysum %></td>
	    <td><%= clsICheFile.FItemList(i).FtaxSum %></td>
	    <% end if %>
	    <td>
	        <% if Not IsNULL(clsICheFile.FItemList(i).FerpLinkType) then %>
	        [<%= clsICheFile.FItemList(i).FerpLinkType %>]
	        <%= clsICheFile.FItemList(i).FerpLinkKey %>
	        <% end if %>
	    </td>
	</tr>
	<% next %>
    </form>    	
</table>


<form name="frmBuf" method="post" action="/admin/tax/eTax_process.asp">
<input type="hidden" name="mode" value="ErpInOutMapping">
<input type="hidden" name="ichedate" value="<%= oneIcheFile.FOneItem.FIcheDate %>"> 
<input type="hidden" name="BIZSECTION_CD" value="<%= oneIcheFile.FOneItem.getBizSectionCD %>">
</form>
<%
set clsICheFile = Nothing
set oneIcheFile = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->