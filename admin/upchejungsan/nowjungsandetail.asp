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
dim yyyymm
dim IsCommissionTax ''������ ���� ����
id      = requestCheckvar(request("id"),10)
gubun   = requestCheckvar(request("gubun"),16)
rectorder = requestCheckvar(request("rectorder"),16)

if rectorder="" then rectorder="orderserial"

if gubun="" then gubun="upche"


dim sqlStr
'dim isLecture
'sqlStr = "select top 1 m.id,m.designerid,c.userdiv "
'sqlStr = sqlStr + " from [db_jungsan].[dbo].tbl_designer_jungsan_master m,"
'sqlStr = sqlStr + " [db_user].[dbo].tbl_user_c c"
'sqlStr = sqlStr + " where m.id="  & id
'sqlStr = sqlStr + " and m.designerid=c.userid"
'
'rsget.Open sqlStr,dbget,1
'if Not rsget.Eof then
'    isLecture = rsget("userdiv")="14"
'end if
'rsget.close

dim isAcademyJungsan

dim ojungsan, ojungsanmaster
set ojungsanmaster = new CUpcheJungsan
ojungsanmaster.FRectId = id
ojungsanmaster.JungsanMasterList

if (ojungsanmaster.FResultCount<1) then
    dbget.Close(): response.end
end if

IsCommissionTax = ojungsanmaster.FItemList(0).IsCommissionTax
isAcademyJungsan = ojungsanmaster.FItemList(0).FtargetGbn="AC"

if (isAcademyJungsan) and (Not IsCommissionTax) and (gubun="upche") then
    gubun="lecture"
end if

set ojungsan = new CUpcheJungsan
ojungsan.FRectid = id
ojungsan.FRectgubun = gubun

'ojungsan.FREctSitename = "N10x10"


if (gubun<>"") and (gubun<>"witakconfirm") then
    if (isAcademyJungsan) then
        ojungsan.JungsanDetailListLectureSum
    else
	    ojungsan.JungsanDetailListWitakSum
	end if
end if

dim i, suplysum, suplytotalsum, selltotalsum
suplysum = 0
suplytotalsum = 0
selltotalsum  = 0


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
	window.open('popetclistadd.asp?id=' + iid + '&gubun=' + igubun,'popetc','width=700, height=150, location=no,menubar=no,resizable=yes,scrollbars=no,status=no,toolbar=no');
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
	//if (frm.gubun[2].checked){
	//	frm.action="nowjungsandetailwitak.asp"
	//}else{
	//	frm.action="nowjungsandetail.asp"
	//}
	frm.submit();
}

function popBatchDetailEdit(id,gubun,itemid,itemoption,sellcash,suplycash,itemname,itemoptionname){
	var popwin = window.open('','jungsandetailedit','width=600,height=200,scrollbars=yes,resizable=yes');
	popwin.focus();

	bufFrm.target="jungsandetailedit";
	bufFrm.id.value = id;
	bufFrm.gubun.value = gubun;
	bufFrm.itemid.value = itemid;
	bufFrm.itemoption.value = itemoption;
	bufFrm.sellcash.value = sellcash;
	bufFrm.suplycash.value = suplycash;
	bufFrm.itemname.value = itemname;
	bufFrm.itemoptionname.value = itemoptionname;

	bufFrm.submit();
}

</script>

<form name="bufFrm" method=post action="popjungsandetailedit.asp">
<input type="hidden" name="id" value=''>
<input type="hidden" name="gubun" value=''>
<input type="hidden" name="itemid" value=''>
<input type="hidden" name="itemoption" value=''>
<input type="hidden" name="sellcash" value=''>
<input type="hidden" name="suplycash" value=''>
<input type="hidden" name="itemname" value=''>
<input type="hidden" name="itemoptionname" value=''>
</form>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="id" value="<%= id %>">
	<input type="hidden" name="rectorder" value="<%= rectorder %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	�귣��ID : <b><%= ojungsanmaster.FItemList(i).Fdesignerid %></b>
        	&nbsp;
			<input type="radio" name="gubun" value="upche" <% if gubun="upche" then response.write "checked" %> >��ü���
			<input type="radio" name="gubun" value="maeip" <% if gubun="maeip" then response.write "checked" %> >����
			<input type="radio" name="gubun" value="witaksell" <% if gubun="witaksell" then response.write "checked" %> >��Ź�Ǹ�
			<input type="radio" name="gubun" value="witakchulgo" <% if gubun="witakchulgo" then response.write "checked" %> >��Ÿ���
			<input type="radio" name="gubun" value="DL" <% if gubun="DL" then response.write "checked" %> >��ۺ�
			<input type="radio" name="gubun" value="DT" <% if gubun="DT" then response.write "checked" %> >�߰���ۺ�
			<input type="radio" name="gubun" value="DP" <% if gubun="DP" then response.write "checked" %> >��Ÿ(���θ��)
			<input type="radio" name="gubun" value="DE" <% if gubun="DE" then response.write "checked" %> >��Ÿ(����)
		<input type="radio" name="gubun" value="lecture" <% if gubun="lecture" then response.write "checked" %> >����
        </td>
        <td align="right">
        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->

<% if gubun<>"witakconfirm" then %>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
        	<img src="/images/icon_arrow_down.gif" align="absbottom">
			<font color="red"><strong>������ �հ�</strong></font>
			(������ �����ϸ� �հ迡 ����˴ϴ�.)
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="50">��ǰ�ڵ�</td>
		<td>��ǰ��</td>
		<td width="80">�ɼǸ�</td>
		<td width="40">����</td>
		<td width="50"><font color="#AAAAAA">�ǸŰ�(����)</font></td>
		<td width="50"><font color="#AAAAAA">���ް�(����)</font></td>
		<td width="50"><font color="#AAAAAA">�ɼ�(����)</font></td>
		<td width="50"><font color="#AAAAAA">�ɼǰ��ް�(����)</font></td>
		<td width="80">�ǸŰ�</td>
		<td width="60">���簡���<br>����</td>
		<td width="80">�Ǹ���</td>
		<td width="80">������</td>
		<td width="80">PG������</td>
		<td width="80">���ް�</td>
		<td width="40">����</td>
		<td width="80">���ް��հ�</td>
		<!-- td width="40">�ϰ�<br>����</td -->
    </tr>
    <% for i=0 to ojungsan.FResultCount-1 %>
    <%
    suplysum =0
    suplysum = suplysum + ojungsan.FItemList(i).Fsuplycash * ojungsan.FItemList(i).FItemNo
    suplytotalsum = suplytotalsum + suplysum
    selltotalsum = selltotalsum + ojungsan.FItemList(i).Fsellcash * ojungsan.FItemList(i).FItemNo

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
      <% if ojungsan.FItemList(i).FOrgsellcash+ojungsan.FItemList(i).FOrgOptaddprice<>ojungsan.FItemList(i).Fsellcash then %>
          <td align="right">
          <font color="#FF0000"><b><%= FormatNumber(ojungsan.FItemList(i).FOrgsellcash,0) %></b></font>
          </td>
      <% else %>
          <td align="right">
          <% if Not IsNULL(ojungsan.FItemList(i).FOrgsellcash) then %>
          <font color="#AAAAAA"><%= FormatNumber(ojungsan.FItemList(i).FOrgsellcash,0) %></font>
          <% end if %>
          </td>
      <% end if %>

      <% if ojungsan.FItemList(i).FOrgsuplycash+ojungsan.FItemList(i).FOrgOptaddbuyprice<>ojungsan.FItemList(i).Fsuplycash then %>
          <td align="right">
          		<font color="#FF0000"><b><%= FormatNumber(ojungsan.FItemList(i).FOrgsuplycash,0) %></b></font>
          </td>
      <% else %>
          <td align="right">
          <% if Not IsNULL(ojungsan.FItemList(i).FOrgsuplycash) then %>
    	      <font color="#AAAAAA"><%= FormatNumber(ojungsan.FItemList(i).FOrgsuplycash,0) %></font>
    	  <% end if %>
          </td>
      <% end if %>
      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).FOrgOptaddprice,0) %></td>
      <td align="right"><%= FormatNumber(ojungsan.FItemList(i).FOrgOptaddbuyprice,0) %></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsellcash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsellcash,0) %></font></td>
      <td align="center">
        <% if ojungsan.FItemList(i).FOrgsellcash<>0 then %>
        <%= 100-CLng((ojungsan.FItemList(i).Fsellcash-ojungsan.FItemList(i).FOrgOptaddprice)/ojungsan.FItemList(i).FOrgsellcash*100) %>
        <% end if %>
      </td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Freducedprice) %>"><%= FormatNumber(ojungsan.FItemList(i).Freducedprice,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fcommission) %>"><%= FormatNumber(ojungsan.FItemList(i).Fcommission,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).FPgcommission) %>"><%= FormatNumber(ojungsan.FItemList(i).FPgcommission,0) %></font></td>
      <td align="right"><font color="<%= MinusFont(ojungsan.FItemList(i).Fsuplycash) %>"><%= FormatNumber(ojungsan.FItemList(i).Fsuplycash,0) %></font></td>
      <td align="center">
      <% if ojungsan.FItemList(i).Fsellcash<>0 then %>
      	<%= 100-CLng(ojungsan.FItemList(i).Fsuplycash/ojungsan.FItemList(i).Fsellcash*100*100)/100 %>
      <% end if %>
      </td>
      <td align="right"><font color="<%= MinusFont(suplysum) %>"><%= FormatNumber(suplysum,0) %></font></td>
      <!-- td><input type="button" value="����" onclick="popBatchDetailEdit('<%= id %>','<%= gubun %>','<%= ojungsan.FItemList(i).FItemID %>','<%= ojungsan.FItemList(i).FItemOption %>','<%= ojungsan.FItemList(i).Fsellcash %>','<%= ojungsan.FItemList(i).Fsuplycash %>','<%= replace(ojungsan.FItemList(i).Fitemname,"'","||39||") %>','<%= replace(ojungsan.FItemList(i).Fitemoptionname,"'","||39||") %>');"></td -->
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF">
      <td colspan="9"></td>

      <td align="right"><%= FormatNumber(selltotalsum,0) %></td>
      <td colspan="5"></td>
      <td align="right"><b><%= FormatNumber(suplytotalsum,0) %></b></td>
      <!-- td></td -->
    </tr>
</table>



<%
ojungsan.FRectOrder= rectorder
ojungsan.JungsanDetailList
%>
<br>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
        	<img src="/images/icon_arrow_down.gif" align="absbottom">
			<% if gubun="upche" then %>
			<font color="red"><strong>��ü��� ����޸�</strong></font>(���� ���׹� �߰� ������ ������ ���� ������ �Է��ϼ���)
			<% elseif gubun="maeip" then %>
			<font color="red"><strong>���� ����޸�</strong></font>(���� ���׹� �߰� ������ ������ ���� ������ �Է��ϼ���)
			<% elseif gubun="witakchulgo" then %>
			<font color="red"><strong>��Ź ��� ����޸�</strong></font>(���� ���׹� �߰� ������ ������ ���� ������ �Է��ϼ���)
			<% elseif gubun="witakoffshop" then %>
			<font color="red"><strong>��Ź �������� �Ǹ� ����޸�</strong></font>(���� ���׹� �߰� ������ ������ ���� ������ �Է��ϼ���)
			<% elseif gubun="witaksell" then %>
			<font color="red"><strong>��Ź �Ǹ� ����޸�</strong></font>(���� ���׹� �߰� ������ ������ ���� ������ �Է��ϼ���)
			<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->



<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="memofrm" method="post" action="dodesignerjungsan.asp">
	<input type="hidden" name="idx" value="<%= ojungsanmaster.FItemList(0).FID %>">
	<input type="hidden" name="gubun" value="<%= gubun %>">
	<input type="hidden" name="mode" value="memoedit">
	<tr align="center" bgcolor="#FFFFFF">
		<td>
			<textarea name="tx_memo" cols="90" rows="2"><%= ojungsanmaster.FItemList(0).Fub_comment %></textarea>
			<input type="button" value="�޸�����" onclick="savememo(memofrm)">
		</td>
	</tr>
	</form>
</table>

<br>

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
        	<img src="/images/icon_arrow_down.gif" align="absbottom">
			<% if gubun="upche" then %>
			<font color="red"><strong>��ü��۳���</strong>(�ִ� 5,000��ǥ��)</font>
			<% elseif gubun="maeip" then %>
			<font color="red"><strong>���� �԰���</strong></font>
			<% elseif gubun="witakchulgo" then %>
			<font color="red"><strong>��Ź �����</strong></font>(���꿡 ���Ե�)
			<% elseif gubun="witaksell" then %>
			<font color="red"><strong>��Ź �Ǹų���</strong></font>(���꿡 ���Ե� : �ִ� 5,000��ǥ��)
			<% elseif gubun="witakoffshop" then %>
			<font color="red"><strong>��Ź �������� �Ǹų���</strong></font>(���꿡 ���Ե�)
			<% end if %>

			<select name="rectorder" onchange="reOrder(this)">
				<option value="orderserial" <% if rectorder="orderserial" then response.write "selected" %> >�ֹ���ȣ��
				<option value="itemid" <% if rectorder="itemid" then response.write "selected" %> >�����ۼ�
			</select>
        </td>
        <td align="right">
        	<input type="button" value="��Ÿ�����߰�" onclick="addEtcList(<%= ojungsanmaster.FItemList(0).FID %>,'<%= gubun %>')">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
<% if (gubun="maeip") then %>
<td width="80">�����ڵ�</td>
<% else %>
<td width="80">�ֹ���ȣ</td>
<% end if %>
<td width="50">�Ǹ�ä��</td>
<td width="50">������</td>
<td width="50">������</td>
<td width="120">�����۸�</td>
<td width="80">�ɼǸ�</td>
<td width="40">����</td>
<td width="50">�ǸŰ�</td>
<td width="50">�Ǹ���</td>
<td width="50">������</td>
<td width="50">PG������</td>
<td width="50">���ް�</td>
<td width="30">����</td>
<td width="30">����</td>
<td width="30">����</td>
</tr>
<% for i=0 to ojungsan.FResultCount-1 %>
<form name="frmBuyPrcSell_<%= i %>" method="post" action="dodesignerjungsan.asp">
<input type="hidden" name="idx" value="<%= ojungsan.FItemList(i).Fid %>">
<input type="hidden" name="midx" value="<%= id %>">
<input type="hidden" name="gubun" value="<%= gubun %>">
<input type="hidden" name="mode" value="">
<tr bgcolor="#FFFFFF">
<td ><%= ojungsan.FItemList(i).Fmastercode %></td>
<td ><%= ojungsan.FItemList(i).FSitename %></td>
<td ><%= ojungsan.FItemList(i).FBuyname %></td>
<td ><%= ojungsan.FItemList(i).FReqname %></td>
<td ><%= ojungsan.FItemList(i).FItemName %></td>
<td ><%= ojungsan.FItemList(i).FItemOptionName %></td>
<td ><input type="text" size="3" name="itemno" value="<%= ojungsan.FItemList(i).FItemNo %>" style="text-align:center"></td>
<td ><input type="text" size="5" name="sellcash" value="<%= ojungsan.FItemList(i).Fsellcash %>" style="text-align:right"></td>
<td ><input type="text" size="5" name="reducedprice" value="<%= ojungsan.FItemList(i).Freducedprice %>" <%=CHKIIF(NOT IsCommissionTax,"readonly class='text_ro'","")%> style="text-align:right"></td>
<td ><input type="text" size="5" name="commission" value="<%= ojungsan.FItemList(i).Fcommission %>" <%=CHKIIF(TRUE or NOT IsCommissionTax,"readonly class='text_ro'","")%> style="text-align:right"></td>
<td ><input type="text" size="5" name="pgcommission" value="<%= ojungsan.FItemList(i).Fpgcommission %>" <%=CHKIIF(TRUE or NOT IsCommissionTax,"readonly class='text_ro'","")%> style="text-align:right"></td>
<td ><input type="text" size="5" name="suplycash" value="<%= ojungsan.FItemList(i).Fsuplycash %>" style="text-align:right"></td>
<td >
<%if ojungsan.FItemList(i).Fsellcash<>0 then %>
<%= 100-ojungsan.FItemList(i).Fsuplycash/ojungsan.FItemList(i).Fsellcash*100 %>
<% end if %>
</td>
<td ><a href="javascript:DelDetail(frmBuyPrcSell_<%= i %>)">����</a></td>
<td ><a href="javascript:ModiDetail(frmBuyPrcSell_<%= i %>)">����</a></td>
</tr>
</form>
<%
'' ���۱������� �ʰ��� �Ʒ� �ּ����� 
if (i mod 1000)=0 then 
    response.flush
end if 
%>
<% next %>
</table>

<br>
<% end if %>

<%
set ojungsan = Nothing
set ojungsanmaster = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->