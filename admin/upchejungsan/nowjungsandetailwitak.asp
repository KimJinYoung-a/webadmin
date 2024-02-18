<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->

<%
dim id,gubun,rectorder
dim witakstep,yyyymm
id = request("id")
gubun = request("gubun")
rectorder = request("rectorder")
witakstep = request("witakstep")

if witakstep="" then witakstep="1"
if rectorder="" then rectorder="itemid"

if gubun="" then gubun="upche"

if (gubun="upche") or (gubun="maeip") then
	witakstep="1"
end if

dim ojungsan, ojungsanmaster
set ojungsan = new CUpcheJungsan
ojungsan.FRectid = id
ojungsan.FRectgubun = gubun

if witakstep="1" then
	ojungsan.JungsanDetailListSum
end if

dim i, suplysum, suplytotalsum
suplysum = 0
suplytotalsum = 0

set ojungsanmaster = new CUpcheJungsan
ojungsanmaster.FRectId = id
ojungsanmaster.JungsanMasterList

yyyymm = ojungsanmaster.FItemList(0).FYYYYmm

dim duplicated
%>
<script language='javascript'>
function reOrder(comp){
	document.frm.rectorder.value=comp.value;
	document.frm.submit();
}

function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function DelDetail(frm){
	var ret = confirm('���� ������ ���� �Ͻðڽ��ϱ�?');
	if (ret){
		frm.mode.value="deldetail";
		frm.submit();
	}
}

function ModiDetail(frm){
	var ret = confirm('���� ������ ���� �Ͻðڽ��ϱ�?');
	if (ret){
		frm.mode.value="modidetail";
		frm.submit();
	}
}

function savememo(frm){
	var ret = confirm('�޸� �����Ͻðڽ��ϱ�?');
	if (ret){
		frm.mode.value = "memoedit";
		frm.submit();
	}
}

function addEtcList(iid,igubun){
	window.open('lib/popetclistadd.asp?id=' + iid + '&gubun=' + igubun,'popetc','width=700, height=300, location=no,menubar=no,resizable=yes,scrollbars=no,status=no,toolbar=no');
}

function Char2Zero(v){
	if (isNaN(v)){
		return 0
	}else{
		return v ;
	}
}

function ReCalcu(frm){
	var ireal

	ireal = Char2Zero(frm.realjaego.value) * 1 ;
	frm.tmpsysjaego.value = Char2Zero(frm.prejaego.value) * 1 + Char2Zero(frm.ipgono.value) * 1	- Char2Zero(frm.chulgono.value) * 1 - Char2Zero(frm.sellno.value) * 1;
	frm.ocha.value = Char2Zero(frm.tmpsysjaego.value) * 1 - ireal;
	frm.jungsanno.value = Char2Zero(frm.chulgono.value) * 1 + Char2Zero(frm.sellno.value) * 1 + Char2Zero(frm.ocha.value) * 1;

}

function ReSearch(frm){
	if (frm.gubun[2].checked){
		frm.action="nowjungsandetailwitak.asp"
	}else{
		frm.action="nowjungsandetail.asp"
	}
	frm.submit();
}
</script>
<table width="760" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="id" value="<%= id %>">
	<input type="hidden" name="rectorder" value="<%= rectorder %>">
	<input type="hidden" name="witakstep" value="<%= witakstep %>">
	<tr>
		<td class="a" >
		<input type="radio" name="gubun" value="upche" <% if gubun="upche" then response.write "checked" %> >��ü���
		<input type="radio" name="gubun" value="maeip" <% if gubun="maeip" then response.write "checked" %> >����
		<input type="radio" name="gubun" value="witak" <% if gubun="witak" then response.write "checked" %> >��Ź

		<td class="a" align="right">
			<a href="javascript:ReSearch(frm);"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width="760" cellspacing="0" class="a">
<tr>
  <td align="right"><a href="/admin/storage/nowjungsanlist.asp?menupos=130">���</a></td>
</tr>
</table>
<% if gubun="witak" then %>
<table width="760" class="a" >
<tr>
	<td></td>
	<td width="120" align="right">
		<% if witakstep="1" then %>
		<a href="?menupos=<%= menupos %>&id=<%= id %>&rectorder=<%= rectorder %>&gubun=<%= gubun %>&witakstep=1"><b>A.��Ź�������</b></a>
		<% else %>
		<a href="?menupos=<%= menupos %>&id=<%= id %>&rectorder=<%= rectorder %>&gubun=<%= gubun %>&witakstep=1">A.��Ź�������</a>
		<% end if %>
	</td>
	<td width="120" align="right">
		<% if witakstep="2" then %>
		<a href="?menupos=<%= menupos %>&id=<%= id %>&rectorder=<%= rectorder %>&gubun=<%= gubun %>&witakstep=2"><b>B.��Ź���곻��Ȯ��</b></a>
		<% else %>
		<a href="?menupos=<%= menupos %>&id=<%= id %>&rectorder=<%= rectorder %>&gubun=<%= gubun %>&witakstep=2">B.��Ź���곻��Ȯ��</a>
		<% end if %>
	</td>
</tr>
</table>
<% end if %>

<% if witakstep="1" then %>
<div class="a">1.������ �հ� (������ �����ϸ� �հ迡 ����˴ϴ�.)</div>
<table width="760" cellspacing="1"  class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
      <td width="40">��ǰID</td>
      <td width="200">��ǰ��</td>
      <td width="80">�ɼǸ�</td>
      <td width="40">����</td>
      <td width="50"><font color="#AAAAAA">�ǸŰ� (����)</font></td>
      <td width="50"><font color="#AAAAAA">���ް� (����)</font></td>
      <td width="80">�ǸŰ�</td>
      <td width="80">���ް�</td>
      <td width="40">����</td>
      <td width="80">���ް��հ�</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    suplysum =0
    suplysum = suplysum + ojungsan.FItemList(i).Fsuplycash * ojungsan.FItemList(i).FItemNo
    suplytotalsum = suplytotalsum + suplysum

    duplicated = ojungsan.CheckDuplicated(i)
    %>

	<% if duplicated then %>
	<tr bgcolor="#EEEEEE">
	<% else %>
    <tr bgcolor="#FFFFFF">
    <% end if %>
      <td ><%= ojungsan.FItemList(i).FItemID %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td align="center"><%= ojungsan.FItemList(i).FItemNo %></td>
      <% if ojungsan.FItemList(i).FOrgsellcash<>ojungsan.FItemList(i).Fsellcash then %>
      <td align="right"><font color="#FF0000"><b><%= FormatNumber(ojungsan.FItemList(i).FOrgsellcash,0) %></b></font></td>
      <% else %>
      <td align="right"><font color="#AAAAAA"><%= FormatNumber(ojungsan.FItemList(i).FOrgsellcash,0) %></font></td>
      <% end if %>

      <% if ojungsan.FItemList(i).FOrgsuplycash<>ojungsan.FItemList(i).Fsuplycash then %>
      <td align="right"><font color="#FF0000"><b><%= FormatNumber(ojungsan.FItemList(i).FOrgsuplycash,0) %></b></font></td>
      <% else %>
      <td align="right"><font color="#AAAAAA"><%= FormatNumber(ojungsan.FItemList(i).FOrgsuplycash,0) %></font></td>
      <% end if %>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsellcash,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsuplycash,0) %></font></td>
      <td align="center"><%= 100-CLng(ojungsan.FItemList(i).Fsuplycash/ojungsan.FItemList(i).Fsellcash*100) %></td>
      <td align="right"><font color="<%= MinusFont(suplysum) %>"><%= FormatNumber(suplysum,0) %></font></td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
      <td colspan="9"></td>
      <td align="right"><%= FormatNumber(suplytotalsum,0) %></td>
    </tr>
</table>
<% end if %>
<%
ojungsan.FRectOrder= rectorder

if witakstep="1" then
ojungsan.JungsanDetailList
%>
<br>
<div class="a">
<% if gubun="upche" then %>
2.��ü��� ����޸�(���� ���׹� �߰� ������ ������ ���� ������ �Է��ϼ���)
<% elseif gubun="maeip" then %>
2.���� ����޸�(���� ���׹� �߰� ������ ������ ���� ������ �Է��ϼ���)
<% else %>
2.��Ź ����޸�(���� ���׹� �߰� ������ ������ ���� ������ �Է��ϼ���)
<% end if %>
<table width="760" cellspacing="1"  class="a" bgcolor=#3d3d3d>
<form name="memofrm" method="post" action="dodesignerjungsan.asp">
<input type="hidden" name="idx" value="<%= ojungsanmaster.FItemList(0).FID %>">
<input type="hidden" name="gubun" value="<%= gubun %>">
<input type="hidden" name="mode" value="memoedit">
<tr bgcolor="#FFFFFF">
	<td>
		<textarea name="tx_memo" cols="90" rows="7"><%= ojungsanmaster.FItemList(0).Fub_comment %></textarea>
		<input type="button" value="�޸�����" onclick="savememo(memofrm)">
	</td>
</tr>
</form>
</table>
</div>
<br>
<table width="760" class="a" border="0">
<tr>
	<td>
	<% if gubun="upche" then %>
	3.��ü��۳���
	<% elseif gubun="maeip" then %>
	3.���� �԰���
	<% else %>
	3.��Ź �԰���
	<% end if %>
	<select name="rectorder" onchange="reOrder(this)">
	<option value="orderserial" <% if rectorder="orderserial" then response.write "selected" %> >�ֹ���ȣ��
	<option value="itemid" <% if rectorder="itemid" then response.write "selected" %> >�����ۼ�
	</select>
	</td>
	<td align="right"><input type="button" value="��Ÿ�����߰�" onclick="addEtcList(<%= ojungsanmaster.FItemList(0).FID %>,'<%= gubun %>')"></td>
</tr>
</div>
<table width="760" cellspacing="1"  class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
      <% if gubun="upche" then %>
      <td width="80">�ֹ���ȣ</td>
      <% elseif gubun="maeip" then %>
      <td width="80">�����ڵ�</td>
      <% else %>
      <td width="80">��Ź�ڵ�</td>
      <% end if %>
      <td width="50">������</td>
      <td width="50">������</td>
      <td width="120">�����۸�</td>
      <td width="80">�ɼǸ�</td>
      <td width="40">����</td>
      <td width="50">�ǸŰ�</td>
      <td width="50">���ް�</td>
      <td width="30">����</td>
      <td width="30">����</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <form name="frmBuyPrc_<%= i %>" method="post" action="dodesignerjungsan.asp">
    <input type="hidden" name="idx" value="<%= ojungsan.FItemList(i).Fid %>">
    <input type="hidden" name="midx" value="<%= id %>">
    <input type="hidden" name="gubun" value="<%= gubun %>">
    <input type="hidden" name="mode" value="">
    <tr bgcolor="#FFFFFF">
      <td ><%= ojungsan.FItemList(i).Fmastercode %></td>
      <td ><%= ojungsan.FItemList(i).FBuyname %></td>
      <td ><%= ojungsan.FItemList(i).FReqname %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td ><input type="text" size="3" name="itemno" value="<%= ojungsan.FItemList(i).FItemNo %>"></td>
      <td ><input type="text" size="5" name="sellcash" value="<%= ojungsan.FItemList(i).Fsellcash %>"></td>
      <td ><input type="text" size="5" name="suplycash" value="<%= ojungsan.FItemList(i).Fsuplycash %>"></td>
      <td ><a href="javascript:DelDetail(frmBuyPrc_<%= i %>)">����</a></td>
      <td ><a href="javascript:ModiDetail(frmBuyPrc_<%= i %>)">����</a></td>
    </tr>
    </form>
    <% next %>
</table>
<br>

<% end if %>

<% if gubun="witak" and witakstep="1" then %>
<%
ojungsan.FRectgubun = "witakchulgo"
ojungsan.JungsanDetailList
%>
<table width="760" border="0">
<tr>
<td align="right"><input type="button" value="��Ÿ�����߰�" onclick="addEtcList(<%= ojungsanmaster.FItemList(0).FID %>,'witakchulgo')"></td>
</tr>
</table>
<table width="760" cellspacing="1"  class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
      <td width="80">����ڵ�</td>
      <td width="50">������</td>
      <td width="50">������</td>
      <td width="120">�����۸�</td>
      <td width="80">�ɼǸ�</td>
      <td width="40">����</td>
      <td width="50">�ǸŰ�</td>
      <td width="50">���ް�</td>
      <td width="30">����</td>
      <td width="30">����</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <form name="frmBuyPrc1_<%= i %>" method="post" action="dodesignerjungsan.asp">
    <input type="hidden" name="midx" value="<%= id %>">
    <input type="hidden" name="idx" value="<%= ojungsan.FItemList(i).Fid %>">
    <input type="hidden" name="gubun" value="<%= gubun %>">
    <input type="hidden" name="mode" value="">
    <tr bgcolor="#FFFFFF">
      <td ><%= ojungsan.FItemList(i).Fmastercode %></td>
      <td ><%= ojungsan.FItemList(i).FBuyname %></td>
      <td ><%= ojungsan.FItemList(i).FReqname %></td>
      <td ><%= ojungsan.FItemList(i).FItemName %></td>
      <td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
      <td ><input type="text" size="3" name="itemno" value="<%= ojungsan.FItemList(i).FItemNo %>"></td>
      <td ><input type="text" size="5" name="sellcash" value="<%= ojungsan.FItemList(i).Fsellcash %>"></td>
      <td ><input type="text" size="5" name="suplycash" value="<%= ojungsan.FItemList(i).Fsuplycash %>"></td>
      <td ><a href="javascript:DelDetail(frmBuyPrc1_<%= i %>)">����</a></td>
      <td ><a href="javascript:ModiDetail(frmBuyPrc1_<%= i %>)">����</a></td>
    </tr>
    </form>
    <% next %>
</table>
<% end if %>

<% if gubun="witak" and witakstep<>"1" then %>
<%
ojungsan.FRectid = id
ojungsan.FrectDesigner = ojungsanmaster.FItemList(0).FDesignerid
ojungsan.FRectStartDay = yyyymm + "-" + "01"
ojungsan.FRectEndDay   = CStr(DateSerial(Left(yyyymm,4), CLng(Right(yyyymm,2))+1,1))
ojungsan.FRectYYYYMM   = yyyymm
ojungsan.FRectPreYYYYMM   = Left(CStr(DateSerial(Left(yyyymm,4), CLng(Right(yyyymm,2))-1,1)),7)

'response.write ojungsan.FRectStartDay
'response.write ojungsan.FRectEndDay
ojungsan.GetWitakJungSanByItem


dim i_pjaego, i_rjaego
dim sysjaego, ocha
dim bufipgo, bufchulgo
dim totjungsanno, totjungsansum
%>
<script language='javascript'>
function saveArr(){
	var frm;
	var upfrm = document.frmarr;

	var ret = confirm('�����Ͻðڽ��ϱ�?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				upfrm.detailidx.value = upfrm.detailidx.value + frm.detailidx.value + "|";
				upfrm.itemid.value = upfrm.itemid.value + frm.itemid.value + "|";
				upfrm.itemoption.value = upfrm.itemoption.value + frm.itemoption.value + "|";
				upfrm.sellcash.value = upfrm.sellcash.value + frm.sellcash.value + "|";
				upfrm.suplycash.value = upfrm.suplycash.value + frm.suplycash.value + "|";
				upfrm.prejaego.value = upfrm.prejaego.value + frm.prejaego.value + "|";
				upfrm.ipgono.value = upfrm.ipgono.value + frm.ipgono.value + "|";
				upfrm.chulgono.value = upfrm.chulgono.value + frm.chulgono.value + "|";
				upfrm.sellno.value = upfrm.sellno.value + frm.sellno.value + "|";
				upfrm.ocha.value = upfrm.ocha.value + frm.ocha.value + "|";
				upfrm.realjaego.value = upfrm.realjaego.value + frm.realjaego.value + "|";
				upfrm.jungsanno.value = upfrm.jungsanno.value + frm.jungsanno.value + "|";
				if (frm.isdelete.checked){
					upfrm.isdelete.value = upfrm.isdelete.value + "Y" + "|";
				}else{
					upfrm.isdelete.value = upfrm.isdelete.value + "N" + "|";
				}

			}
		}

		upfrm.submit();
	}
}

function delArr(){
	var upfrm = document.frmarr;
	var ret = confirm('Ȯ���� �����͸� ���� �Ͻðڽ��ϱ�?');
	if (ret){
		upfrm.gubun.value = "witakjungsan_del";
		upfrm.submit();
	}
}
</script>

<table width="1000" cellspacing="1"  class="a" >
<tr>
	<td width="140"><%= yyyymm %> ��Ź���곻��</td>
	<% if ojungsan.FWitakInsserted then %>
	<td width="200"><font color="red"> �� �����ʹ� Ȯ���� ����Ÿ �Դϴ�.</font></td>
	<td><!-- <b><a href="javascript:delArr()">[����]</a></b> --></td>
	<% end if %>
	<td align="right"><input type="button" value="����Ȯ��" onclick="saveArr()"></td>
</tr>
</table>

<table width="1000" cellspacing="1"  class="a" bgcolor=#3d3d3d>
    <tr bgcolor="#DDDDFF">
    	<td width="50">��ǰ��ȣ</td>
    	<td width="120">��ǰ��</td>
    	<td width="80">�ɼ�</td>
    	<td width="50">�Һ��ڰ�(����)</td>
    	<td width="50">���ް�(����)</td>
    	<td width="50">�Һ��ڰ�(�Ǹ�)</td>
    	<td width="50">���ް�(�Ǹ�)</td>
    	<td width="30"></td>
    	<td width="50">�̿���� (A)</td>
    	<td width="50">�԰���� (B)</td>
    	<td width="50">������ (C)</td>
    	<td width="50">�Ǹż��� (D)</td>
    	<td width="60">�ý������ (S=A+B-C-D)</td>
    	<td width="50">����<br>(E=S-R)</td>
    	<td width="50">�ǻ���� (R)</td>
    	<td width="50">���꿹������(C+D+E)</td>
    	<td width="50">�������</td>
    	<td width="20">����</td>
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    	duplicated = ojungsan.CheckDuplicated(i)
   	%>

    <% if (ojungsan.FItemList(i).FIsUsing<>"Y") and (ojungsan.FItemList(i).FIpGoNo=0) and (ojungsan.FItemList(i).FChulgoNo=0) and (ojungsan.FItemList(i).FsellNo=0) then %>
    <% else %>
        <form name="frmBuyPrc_<%= i %>" method="post" action="">
        <input type="hidden" name="detailidx" value="<%= ojungsan.FItemList(i).Fdetailidx %>">
		<input type="hidden" name="itemid" value="<%= ojungsan.FItemList(i).Fitemid %>">
		<input type="hidden" name="itemoption" value="<%= ojungsan.FItemList(i).Fitemoption %>">
    	<% if duplicated then %>
    		<tr bgcolor="#EEEEEE">
    	<% else %>
		    <% if ojungsan.FItemList(i).FIsUsing<>"Y" then %>
		    <tr bgcolor="#FFFFFF" class="gray">
		    <% else %>
		    <tr bgcolor="#FFFFFF">
		    <% end if %>
		<% end if %>
	    	<td><%= ojungsan.FItemList(i).Fitemid %></td>
	    	<td><%= ojungsan.FItemList(i).Fitemname %></td>
	    	<td><%= ojungsan.FItemList(i).Fitemoptionname %></td>
	    	<td align="right"><%= ojungsan.FItemList(i).FSellcash %></td>
	    	<td align="right"><%= ojungsan.FItemList(i).FSuplycash %></td>
	    	<td align="right"><input type="text" name="sellcash" value="<%= ojungsan.FItemList(i).FSellcash_sell %>" size="6" style="border-width:1; border-color:#AAAAAA; border-style:solid;" ></td>
	    	<td align="right"><input type="text" name="suplycash" value="<%= ojungsan.FItemList(i).FSuplycash_sell %>" size="6" style="border-width:1; border-color:#AAAAAA; border-style:solid;" ></td>
	    	<td align="center"><%= ojungsan.FItemList(i).FPrejaego %></td>

	    	<td align="center"><input type="text" name="prejaego" size=3 value="<%= ojungsan.FItemList(i).FPrejaego %>" style="border-width:1; border-color:#AAAAAA; border-style:solid;" onKeyUp="javascript:ReCalcu(frmBuyPrc_<%= i %>)"></td>
	    	<td align="center"><input type="text" name="ipgono" size=3 value="<%= ojungsan.FItemList(i).FIpgoNo %>" style="border-width:1; border-color:#AAAAAA; border-style:solid;" onKeyUp="javascript:ReCalcu(frmBuyPrc_<%= i %>)"></td>
	    	<td align="center"><input type="text" name="chulgono" size=3 value="<%= ojungsan.FItemList(i).FChulgono %>" style="border-width:1; border-color:#AAAAAA; border-style:solid;" onKeyUp="javascript:ReCalcu(frmBuyPrc_<%= i %>)"></td>
	    	<td align="center"><input type="text" name="sellno" value="<%= ojungsan.FItemList(i).FsellNo %>" size="4" style="border-width:1; border-color:#AAAAAA; border-style:solid;" onKeyUp="javascript:ReCalcu(frmBuyPrc_<%= i %>)"></td>
	    	<td align="center"><input type="text" name="tmpsysjaego" size="3" value="<%= ojungsan.FItemList(i).FsysJaeGo %>" style="border-width:1; border-color:#FFFFFF; border-style:solid; " readonly ></td>
	    	<% if ojungsan.FItemList(i).FOCha<>0 then %>
	    	<td align="center"><input type="text" name="ocha" size="3" value="<%= ojungsan.FItemList(i).FOCha %>" style="border-width:1; border-color:#FFFFFF; border-style:solid; color:#FF0000"></td>
	    	<% else %>
	    	<td align="center"><input type="text" name="ocha" size="3" value="<%= ojungsan.FItemList(i).FOCha %>" style="border-width:1; border-color:#FFFFFF; border-style:solid;"></td>
	    	<% end if %>
	    	<td align="center"><input type="text" name="realjaego" size=3 value="<%= ojungsan.FItemList(i).FRealJaego %>" style="border-width:1; border-color:#AAAAAA; border-style:solid;" onKeyUp="javascript:ReCalcu(frmBuyPrc_<%= i %>)"></td>
	    	<td align="center"><%= ojungsan.FItemList(i).FjungsanNo %></td>
	    	<td align="center"><input type="text" name="jungsanno" size="3" value="<%= ojungsan.FItemList(i).FjungsanNo %>" style="border-width:1; border-color:#AAAAAA; border-style:solid;"></td>
	    	<td align="center"><input type="checkbox" name="isdelete" <% if ojungsan.FItemList(i).FIsDelete="Y" then response.write "checked" %> ></td>
	    	<%
	    		if ojungsan.FItemList(i).FIsDelete<>"Y" then
		    		totjungsanno = totjungsanno + ojungsan.FItemList(i).FjungsanNo
		    		totjungsansum = totjungsansum + ojungsan.FItemList(i).FSuplycash_sell * ojungsan.FItemList(i).FjungsanNo
	    		end if
	    	%>
	    </tr>
	    </form>
	<% end if %>
    <% next %>
    <tr bgcolor="#FFFFFF">
    	<td width="50">�Ѱ�</td>
    	<td align="right" colspan="17">�� �Ǽ� : <%= totjungsanno %> �� �ݾ� : <%= FormatNumber(totjungsansum,0) %></td>
    </tr>
</table>
<form name="frmarr" method="post" action="dodesignerjungsan.asp">
<input type="hidden" name="mode" value="arrsave">
<input type="hidden" name="gubun" value="witakjungsan">
<input type="hidden" name="idx" value="<%= id %>">
<input type="hidden" name="detailidx" value="">
<input type="hidden" name="itemid" value="">
<input type="hidden" name="itemoption" value="">
<input type="hidden" name="sellcash" value="">
<input type="hidden" name="suplycash" value="">
<input type="hidden" name="prejaego" value="">
<input type="hidden" name="ipgono" value="">
<input type="hidden" name="chulgono" value="">
<input type="hidden" name="sellno" value="">
<input type="hidden" name="ocha" value="">
<input type="hidden" name="realjaego" value="">
<input type="hidden" name="jungsanno" value="">
<input type="hidden" name="isdelete" value="">

</form>
<% end if %>
<%
set ojungsan = Nothing
set ojungsanmaster = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->