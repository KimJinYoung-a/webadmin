<%@ language=vbscript %>
<% option explicit %>
<%
Server.ScriptTimeOut = 60
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim yyyy1, mm1, gubun, page, grpon
dim yyyy_t, mm_t
yyyy1 = requestCheckvar(request("yyyy1"),10)
mm1 = requestCheckvar(request("mm1"),10)
gubun = requestCheckvar(request("gubun"),16)
page = requestCheckvar(request("page"),10)
grpon = requestCheckvar(request("grpon"),10)



if (gubun="") then gubun="ext"
if (page="") then page=1

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

yyyy_t  = yyyy1 'request("yyyy1")
mm_t    = mm1 'request("mm1")

dim ojungsan, ArrRows
set ojungsan = new CUpcheJungsan
ojungsan.FPageSize = 3000
ojungsan.FCurrPage = page
ojungsan.FRectGubun = gubun
ojungsan.FRectYYYYMM = yyyy1 + "-" + mm1

if (gubun="witakchulgo") or (gubun="witakchulgoJS") then
    if (gubun="witakchulgoJS") then ojungsan.FRectNotIncDivcode999="on"

    if (grpon<>"") and (gubun="witakchulgoJS") then
        ArrRows = ojungsan.SearchWitakMaeipChulgoJungsanListGrp
    else
	    ojungsan.SearchWitakMaeipChulgoJungsanList
    end if
end if

dim i, precode, ischeckd, isdisabled
dim checkdate1, checkdate2

%>
<script language='javascript'>
function popConfirm(yyyymm){
    var popwin = window.open('checkDuplicatedJungsan.asp?yyyymm=' + yyyymm,'checkDuplicatedJungsan','width=800,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popConfirm2(yyyymm){
    var popwin = window.open('checkDuplicatedJungsan_etc.asp?yyyymm=' + yyyymm,'checkDuplicatedJungsan','width=800,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function SelectCkMonly(opt){
	var bool = opt.checked;

	for (var i=0;i<document.forms.length;i++){
		var frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.hideMw.value=="M") {
			    frm.cksel.checked = bool;
			    AnCheckClick(frm.cksel);
			}
		}
	}


}

function SaveArr(igubun){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	upfrm.mode.value= igubun;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	var ret;

    upfrm.idx.value = "";
    upfrm.yyyy.value = frmDumi.yyyy1.value;
    upfrm.mm.value  = frmDumi.mm1.value;

	if (!pass) {
		ret = confirm('���� ������ �����ϴ�. \r\n\r\n ' + upfrm.yyyy.value + '-' + upfrm.mm.value + ' ������ �������� ���� �Ͻðڽ��ϱ�?');
		if (!ret){
			return;
		}else{

		}
	}else{
		ret = confirm('���� ������ ' + upfrm.yyyy.value + '-' + upfrm.mm.value + ' ������ �������� ���� �Ͻðڽ��ϱ�?');
	}



	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.idx.value = upfrm.idx.value + frm.idx.value + ",";
				}
			}
		}
		upfrm.mode.value=igubun;
		upfrm.submit();
	}
}


function jsEtcChulgoJungsan(mayjacctcd,makerid,vatinclude){
	var upfrm1 = document.frmEtcJOne;
    upfrm1.mayjacctcd.value=mayjacctcd;
    upfrm1.makerid.value=makerid;
    upfrm1.vatyn.value=vatinclude;

    if (confirm("�ۼ� �Ͻðڽ��ϱ�?")){
        upfrm1.submit();    
    }
}
</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	��������:<% DrawYMBox yyyy1,mm1 %>&nbsp;&nbsp;
        	<input type="radio" name="gubun" value="witakchulgoJS" <% if gubun="witakchulgoJS" then response.write "checked" %> >��Ÿ���(������)
            (<input type="checkbox" name="grpon" <% if grpon="on" then response.write "checked" %> <%=CHKIIF(gubun<>"witakchulgoJS","disabled","")%> >�հ躸��)
            &nbsp;&nbsp;&nbsp;
			<input type="radio" name="gubun" value="witakchulgo" <% if gubun="witakchulgo" then response.write "checked" %> >��Ÿ���(�������)
        </td>
        <td align="right">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->

<% if (grpon<>"") and (gubun="witakchulgoJS") then %>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <form name="frmDumi2">
	<tr>
		<td height="1" colspan="3" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td>
            
        </td>
        <td align="right">
            
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	
      <td width="20" >��������</td>
      <td width="100">�귣��ID</td>
      <td width="60">��������</td>
      <td width="80">������</td>
      <td width="80">���ݾ�</td>
      <td width="60">�����Ծ�</td>
      <td width="20"></td>
      <td>����TITLE</td>
      <td width="80">�������</td>
      <td width="80">���걸��</td>
      <td width="80">�����������</td>
      <td width="50">����</td>

      <td width="80">�������</td>
      <td width="80">�����ǸŰ���</td>
      <td width="80">������԰���</td>
      <td width="50">����</td>
      <td width="50">���</td>
    </tr>
    <% if isArray(arrRows) then %>
    <% For i =0 To UBound(ArrRows,2) %>
    
    <tr align="center" bgcolor="#FFFFFF"  >
      <td><%= ArrRows(0,i) %></td>
      <td><%= ArrRows(1,i) %></td>
      <td><%= ArrRows(2,i) %></td>
      <td><%= ArrRows(3,i) %></td>
      <td align="right"><%= FormatNumber(ArrRows(4,i),0) %></td>
      <td align="right"><%= FormatNumber(ArrRows(5,i),0) %></td>
      <td></td>
      <td><%= ArrRows(6,i) %></td>
      <td><%= ArrRows(7,i) %></td>
      <td><%= ArrRows(8,i) %></td>
      <td><%= ArrRows(9,i) %></td>
      <td><%= ArrRows(10,i) %></td>
      <td><%= ArrRows(11,i) %></td>
      <td><%= ArrRows(12,i) %></td>
      <td><%= ArrRows(13,i) %></td>
      <td><%= CHKIIF(ArrRows(14,i)=1,"x","") %></td>
      <td>
        <% if (ArrRows(14,i)=1) then %>
        <input type="button" value="�ۼ�" onClick="jsEtcChulgoJungsan('<%= ArrRows(0,i) %>','<%= ArrRows(1,i) %>','<%= ArrRows(2,i) %>')">
        <% end if %>
      </td>
    </tr>
    <% next %>
    <% end if %>
</table>
<% elseif (gubun="witakchulgo") or (gubun="witakchulgoJS") then %>
<!-- ��Ź ��� -->

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <form name="frmDumi">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td>
            <input type="checkbox" name="ckM" onClick="SelectCkMonly(this)">���Ը� üũ

            <b>itemoutlet Ȯ��..</b>
        </td>
        <td align="right">
            ����ݿ����:<% DrawYMBox yyyy_t,mm_t %>
            <!--
        	<input type="button" value="�����������(����̹ݿ�)" onclick="SaveArr('maeipchulgo');">
        	-->
            

			<input type="button" value="���ó��� ��Ź�������(����ݿ�)" onclick="SaveArr('witakchulgo');">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    </form>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="20" ><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
      <td width="120">����ڵ�</td>
      <td width="100">���ó</td>
      <td>���ó����</td>
      <td>�������</td>
      <td>�������</td>
      <td width="100">�����԰�</td>
      <td width="40">����</td>
      <td width="50">���Ա���</td>
      <td width="50">��������</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <% if precode<>ojungsan.FItemList(i).FCode then %>
    <%
    	ischeckd = false
    	isdisabled = false
    	if Not IsNull(ojungsan.FItemList(i).FScheduleDate) then
    		ischeckd = ((Cdate(ojungsan.FItemList(i).FScheduleDate)<checkdate1) and (Cdate(ojungsan.FItemList(i).FScheduleDate)>=checkdate2))
    	end if

    	isdisabled = (ojungsan.FItemList(i).FDivCode="999") or (ojungsan.FItemList(i).FDesignerID="itemstockmodify")
    %>
    <tr bgcolor="#CCCCCC">
    	<td colspan="10"></td>
    </tr>
    <tr  align="center" bgcolor="#FFFFFF">
      <td align="left" colspan="2"><b><%= ojungsan.FItemList(i).FCode %></b></td>
      <td><b><%= ojungsan.FItemList(i).FDesignerID %></b></td>
      <td><%= ojungsan.FItemList(i).Fchargename %></td>
      <td><%= ojungsan.FItemList(i).FScheduleDate %></td>
	  <td><%= left(ojungsan.FItemList(i).FRegDate,10) %></td>
      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).FTotalbuycash,0) %></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <tr bgcolor="#FFFFFF">
    	<td colspan="10"><%= ojungsan.FItemList(i).FDivCode %> ��Ÿ���� : <%= ojungsan.FItemList(i).FComment %> </td>
    </tr>
    <% end if %>
    <% precode = ojungsan.FItemList(i).FCode %>
    <form name="frmBuyPrc_<%= i %>" >
    <input type="hidden" name="idx" value="<%= ojungsan.FItemList(i).FID %>">
    <tr align="center" bgcolor="#FFFFFF" <% if ischeckd then response.write "class='H'" %> >
      <td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" <% if ischeckd then response.write "checked" %> <% if isdisabled then response.write "disabled" %>></td>
      <td><%= ojungsan.FItemList(i).FMakerid %></td>
      <td><%= ojungsan.FItemList(i).FItemGubun %>-<%= ojungsan.FItemList(i).FItemId %>-<%= ojungsan.FItemList(i).FItemOption %></td>
      <td align="left" colspan="2"><%= ojungsan.FItemList(i).FItemName %></td>
      <td><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).FSuplycash,0) %>(<%= FormatNumber(ojungsan.FItemList(i).FSuplycash2,0) %>)</td>
      <td align="center"><%= ojungsan.FItemList(i).FItemNo %></td>
      <td><font color="<%= mwdivColor(ojungsan.FItemList(i).FMWDiv) %>"><%= ojungsan.FItemList(i).FMWDiv %></font>
        <% if ojungsan.FItemList(i).FMWDiv="H" then %><font color="red">Ȯ��</font><% end if %>
      <input type="hidden" name="hideMw" value="<%= ojungsan.FItemList(i).FMWDiv %>">
      </td>
      <td>
      	<% if ojungsan.FItemList(i).Fvatinclude<>"Y" then %>
      	<font color=red>�鼼</font>
      	<% end if %>
      </td>
    </tr>
    </form>
    <% next %>
    <tr bgcolor="#CCCCCC">
    	<td colspan="10"></td>
    </tr>
</table>
<% end if %>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<form name="frmArrupdate" method="post" action="dobatch.asp">
<input type="hidden" name="idx" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="yyyy" value="<%= yyyy1 %>">
<input type="hidden" name="mm" value="<%= mm1 %>">
</form>
<form name="frmEtcJOne" method="post" action="dobatch.asp">
<input type="hidden" name="mode" value="etcChulgoJOne">
<input type="hidden" name="yyyy" value="<%= yyyy1 %>">
<input type="hidden" name="mm" value="<%= mm1 %>">
<input type="hidden" name="mayjacctcd" value="">
<input type="hidden" name="makerid" value="">
<input type="hidden" name="vatyn" value="">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->