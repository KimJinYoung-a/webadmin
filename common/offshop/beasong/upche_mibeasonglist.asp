<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : �������� ���
' Hieditor : 2011.02.26 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->

<%
dim searchType, searchValue, MisendReason ,ojumun ,i,iy
	searchType      = request("searchType")
	searchValue     = request("searchValue")
	MisendReason    = request("MisendReason")

set ojumun = new cupchebeasong_list
	ojumun.FRectSearchType  = SearchType
	ojumun.FRectSearchValue = SearchValue

	if (MisendReason="") then
	    ojumun.FRectMisendReason = "AA"
	else
	    ojumun.FRectMisendReason = MisendReason
	end if

	ojumun.FRectDesignerID = session("ssBctID")
	ojumun.fDesignerDateBaljuinputlist()
%>

<script language='javascript'>

function chksubmit(){
    var frm = document.frm;

    if ((frm.searchType.value.length>0)&&(frm.searchValue.value.length<1)){
        alert('�˻� ������ �Է��ϼ���.');
        frm.searchValue.focus();
        return;
    }

    if ((frm.searchType.value=="orderno")||(frm.searchType.value=="itemid")){
        if (!IsDigit(frm.searchValue.value)){
            alert('�˻� ������ ���ڸ� �����մϴ�.');
            frm.searchValue.focus();
            return;
        }
    }

    frm.submit();
}


function ShowOrderInfo(masteridx){
	var ShowOrderInfo = window.open('/common/offshop/beasong/upche_viewordermaster.asp?masteridx='+masteridx,'ShowOrderInfo','width=800,height=768,scrollbars=yes,resizable=yes');
	ShowOrderInfo.focus();
}

function switchCheckBox(comp){
    var frm = comp.form;

	if(frm.chkidx.length>1){
		for(i=0;i<frm.chkidx.length;i++){
			frm.chkidx[i].checked = comp.checked;
			AnCheckClick(frm.chkidx[i]);
		}
	}else{
		frm.chkidx.checked = comp.checked;
		AnCheckClick(frm.chkidx);
	}
}

function BaljuReprint(){
    var frm = document.frmbalju;
	var pass = false;

    if(!frm.chkidx.length){
    	pass = frm.chkidx.checked;
    }else{
        for (var i=0;i<frm.chkidx.length;i++){
    	    pass = (pass||frm.chkidx[i].checked);
    	}
    }

	if (!pass) {
		alert("������� ������ �����ϼ���.");
		return;
	}else{
	    var popwin = window.open("about:blank","PopBaljuList","width=800,scrollbars=yes,resizable");
	    frm.target = "PopBaljuList";
	    frm.isall.value = "";
 		frm.action = "/common/offshop/beasong/upche_reselectbaljulist.asp";
		frm.submit();
	}
}

function BaljuReprintAll(){
    var frm = document.frmbalju;

    if (confirm('����� ���� ��ü ���ּ��� ����� �Ͻðڽ��ϱ�?')){
        var popwin = window.open("about:blank","PopBaljuList","width=800,scrollbars=yes,resizable");
	    frm.target = "PopBaljuList";
	    frm.isall.value = "on";
 		frm.action = "/common/offshop/beasong/upche_reselectbaljulist.asp";
		frm.submit();
    }
}

function popMisendInput(detailidx){
    var popwin = window.open('/common/offshop/beasong/upche_popMisendInput.asp?detailidx=' + detailidx,'popMisendInput','width=600,height=768,scrollbars=yes,resizable=yes');
    popwin.focus();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="" onsubmit="chksubmit(); return false">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="<%= adminColor("topbar") %>" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left" bgcolor="#FFFFFF">
		<select class="select" name="searchType" >
			<option value="">�˻�����</option>
			<option value="orderno" <%= ChkIIF(searchType="orderno","selected","") %> >�ֹ���ȣ</option>
			<option value="itemid" <%= ChkIIF(searchType="itemid","selected","") %> >��ǰ�ڵ�</option>
			<option value="buyname" <%= ChkIIF(searchType="buyname","selected","") %> >������</option>
			<option value="reqname" <%= ChkIIF(searchType="reqname","selected","") %> >������</option>
		</select>
		<input type="text" class="text" name="searchValue" value="<%= searchValue %>" size="20" maxlength="20">
		&nbsp;
		<!--�����Է¿��� :-->
		<!--<select class="select" name="MisendReason">-->
		<!--	<option value="" >��ü</option>-->
		<!--	<option value="NN" <%'= ChkIIF(MisendReason="NN","selected","") %> >�������Է�</option>-->
		<!--	<option value="03" <%'= ChkIIF(MisendReason="03","selected","") %> >�������</option>-->
			<!--<option value="05" <%'= ChkIIF(MisendReason="05","selected","") %> >ǰ�����Ұ�</option>-->
			<!--<option value="02" <%'= ChkIIF(MisendReason="02","selected","") %> >�ֹ�����</option>-->
		<!--</select>-->
		<br>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:chksubmit();">
	</td>
</tr>
</form>
</table>

<br>
<!--
�� ��������� ��� ������ SMS �� �ȳ����� �߼�<br>
ǰ�����Ұ��� ���, �����Ϳ��� ó��
-->
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr align="center">
	<td align="left">
    	<input type="button" class="button" value="���ó��� ���ּ� �����" onclick="javascript:BaljuReprint()">
		&nbsp;
    	<input type="button" class="button" value="�������ü ���ּ� �����" onclick="javascript:BaljuReprintAll()">
    </td>
    <td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmbalju" method="post" >
<input type="hidden" name="mode" value="">
<input type="hidden" name="isall" value="">
<input type="hidden" name="ArrChkVal" value="">
<tr bgcolor="FFFFFF">
	<td height="25" colspan="15">
		�˻���� : <b><% = ojumun.FResultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="chkAll" onClick="switchCheckBox(this);"></td>
	<td>�ϷĹ�ȣ</td>
	<td>�ֹ���ȣ</td>
	<td>������</td>
	<td>��ǰ�ڵ�</td>
	<td>��ǰ��<font color="blue">&nbsp;[�ɼ�]</font></td>
	<td>���ް�</td>
	<td>�ǸŰ�</td>
	<td>����</td>
	<td>��������<!-- �ֹ��뺸�� --></td>
	<td>�ֹ�Ȯ����</td>
	<td>�����</td>
	<!--<td>��������</td>
	<td>�������</td>
	<td>��������<br>�Է�</td>-->
</tr>
<% if ojumun.FResultCount > 0 then %>
<% for i=0 to ojumun.FresultCount-1 %>
<input type="hidden" name="detailidx" value="<%= ojumun.FItemList(i).fdetailidx %>">
<tr align="center" class="a" bgcolor="#FFFFFF">
	<td><input type="checkbox" name="chkidx" value="<%= ojumun.FItemList(i).Fdetailidx %>" onClick="AnCheckClick(this);"></td>
	<td><%= ojumun.FItemList(i).fdetailidx %></td>
	<td height="25">
		<a href="javascript:ShowOrderInfo('<%= ojumun.FItemList(i).fmasteridx %>')">
		<%= ojumun.FItemList(i).Forderno %></a>
	</td>
	<td><%= ojumun.FItemList(i).FReqname %></td>
	<td><%= ojumun.fitemlist(i).fitemgubun %>-<%= FormatCode(ojumun.fitemlist(i).FitemID) %>-<%= ojumun.fitemlist(i).fitemoption %></td>
	<td align="left">
		<%= ojumun.FItemList(i).FItemname %>
		<% if (ojumun.FItemList(i).fitemoptionname<>"") then %>
		<font color="blue">[<%= ojumun.FItemList(i).fitemoptionname %>]</font>
		<% end if %>
	</td>
	<td><%= FormatNumber(ojumun.fitemlist(i).fsuplyprice,0) %></td>
	<td><%= FormatNumber(ojumun.fitemlist(i).fsellprice,0) %></td>
	<td><%= ojumun.FItemList(i).Fitemno %></td>
	<td><acronym title="<%= ojumun.FItemList(i).Fbaljudate %>"><%= left(ojumun.FItemList(i).Fbaljudate,10) %></acronym></td>
	<td><acronym title="<%= ojumun.FItemList(i).Fupcheconfirmdate %>"><%= left(ojumun.FItemList(i).Fupcheconfirmdate,10) %></acronym></td>
	<td>
	    <% if IsNULL(ojumun.FItemList(i).Fbaljudate) then %>
        D+0
        <% elseif datediff("d",(left(ojumun.FItemList(i).Fbaljudate,10)) , (left(now,10)) )>2 then %>
        <font color="red"><b>D+<%= datediff("d",(left(ojumun.FItemList(i).Fbaljudate,10)) , (left(now,10)) ) %></b></font>
        <% else %>
        D+<%= datediff("d",(left(ojumun.FItemList(i).Fbaljudate,10)) , (left(now,10)) ) %>
        <% end if %>
    </td>
	<!--<td><%'= ojumun.FItemList(i).getMisendText %></td>
	<td><%'= ojumun.FItemList(i).FMisendIpgodate %></td>
    <td>
        <%' if (ojumun.FItemList(i).isMisendAlreadyInputed) then %>
        <a href="javascript:popMisendInput('<%= ojumun.FItemList(i).Fdetailidx %>');">�󼼺���</a>
        <%' else %>
        <a href="javascript:popMisendInput('<%= ojumun.FItemList(i).Fdetailidx %>');"><font color="#AAAAAA">�Է�</font></a>
        <%' end if %>
    </td>-->
</tr>
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</form>
</table>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->