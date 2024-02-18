<%@ language=vbscript %>
<% option explicit %>
<%
'Server.ScriptTimeOut = 60
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim yyyy1, mm1, gubun, page
dim yyyy_t, mm_t
yyyy1 = requestCheckvar(request("yyyy1"),10)
mm1 = requestCheckvar(request("mm1"),10)
gubun = requestCheckvar(request("gubun"),16)
page = requestCheckvar(request("page"),10)



if (gubun="") then gubun="chk0"
if (page="") then page=1

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

yyyy_t  = request("yyyy1")
mm_t    = request("mm1")

dim ojungsan
set ojungsan = new CUpcheJungsan
ojungsan.FPageSize = 3000
ojungsan.FCurrPage = page
ojungsan.FRectGubun = gubun
ojungsan.FRectYYYYMM = yyyy1 + "-" + mm1

' if (gubun="witakchulgo") or (gubun="witakchulgoJS") then
'     if (gubun="witakchulgoJS") then ojungsan.FRectNotIncDivcode999="on"
' 	ojungsan.SearchWitakMaeipChulgoJungsanList
' end if

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

function dobatch(frm,mode){



	frm.mode.value=mode;
	var ret = confirm('�ϰ�ó���� ���� �Ͻðڽ��ϱ�?');
	if(ret){
		frm.submit();
	}
}

function etcBrandCpnJungsan(comp){
	if (confirm("�귣������ ���� ���� �ۼ� �Ͻðڽ��ϱ�?")){
		comp.form.submit();
	}
}

function etcBrandCpnIdxJungsan(comp){
	if (confirm("�귣������ ���� ���� �ۼ� �Ͻðڽ��ϱ�?")){
		comp.form.submit();
	}
}

function jsAddextBeasongPay(comp){
	if (confirm("���޸� ������ۺ� ������ �Ͻðڽ��ϱ�?")){
		comp.form.submit();
	}
}

function popJungsanCheck(idifftp){
	var popwin = window.open("","popJungsanCheck","width=1200,height=800,scrollbars=yes,resizable=yes,status=yes");
	popwin.location.href="/admin/jungsan/popJungsanCheck.asp?difftp="+idifftp;

	popwin.focus();

}
</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" align="center" cellpadding="4" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">

	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" height="40" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
        <td align="left">
        	��������:<% DrawYMBox yyyy1,mm1 %>
        </td>

        <td rowspan="2" align="right" width="50">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td align="left">
			<input type="radio" name="gubun" value="chk0" <% if gubun="chk0" then response.write "checked" %> > ���� - ����
			&nbsp;
			<input type="radio" name="gubun" value="chk2" <% if gubun="chk2" then response.write "checked" %> > ���� - OFF
			&nbsp;
			<input type="radio" name="gubun" value="chk1" <% if gubun="chk1" then response.write "checked" %> > ���� - ON

			&nbsp;|&nbsp;

			<input type="radio" name="gubun" value="act0" <% if gubun="act0" then response.write "checked" %> > ���� BATCH ó��


			&nbsp;|&nbsp;
			<input type="radio" name="gubun" value="ext0" <% if gubun="ext0" then response.write "checked" %> > ��Ÿ���� - �߰���ۺ� / ��Ÿ��� / �귣������ / ���Ը�������


			&nbsp;|&nbsp;
			<input type="radio" name="gubun" value="chk9" <% if gubun="chk9" then response.write "checked" %> > �������
			&nbsp;
			<input type="radio" name="gubun" value="act9" <% if gubun="act9" then response.write "checked" %> > �������ó��
        </td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->
<p>

<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name=barchForm method=post action="dobatch.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="yyyy" value="<%= yyyy1 %>">
	<input type="hidden" name="mm" value="<%= mm1 %>">
	</form>

	<% if gubun="chk0" then %>
	<tr bgcolor="#FFFFFF">
		<td style="line-height:25px">
		<!-- ���� -->
		<strong>1. ���� ���곻�� ���� Ȯ��</strong>
		<br>&nbsp; - <a href="/admin/maechul/extjungsandata/extJungsanDataStatistic.asp?menupos=1656" target="_menu1656">[�濵]�����ڷ�>>���޸��������</a> : ���۽�Ʈ�� �����Ͽ� ������ ���űݾ��� �´��� �����Ѵ�.
		<br>&nbsp; - <a href="https://docs.google.com/spreadsheets/d/1MNrTeCz1RvLE-Neuoh7RCQR5NstxWmeg5BLJAKTqH2o/edit#gid=441301323" target="_441301323">[googlesheet]���޸������</a>
		<br>&nbsp; - <a href="/admin/maechul/extjungsandata/extJungsanDataList.asp?menupos=1652&mimap=on" target="_menu1652">[�濵]�����ڷ�>>���޸����곻��</a> : �̸��γ����� ������Ѵ�.(��ǰ�ǰ��)
		<br>&nbsp; - [�濵]�����ڷ�>>���޸�������� �� [����vs�ֹ��Է°���] �� ���������� ������ �����Ѵ�.
		<br>
		<br>
		<strong>2. ���� ���԰� Ȯ��</strong>
		<br>&nbsp; - <a href="/admin/etc/difforder/orderMarginErrList.asp?menupos=3956" target="_menu3956">[��������]���޸�����>>���޻� ���� üũ</a>
		: ���� �Է� ���԰��� ����� �Ǿ� �ִ��� Ȯ���Ѵ�. ����unit���� ��ü�� �����ؼ� ���� ���� �����ϴ°�쵵 �ִ�.
		<br>
		<br>
	<% if (FALSE) then %>
		<strong>3. </strong>
	<% end if%>

		</td>
	</tr>
	<% elseif gubun="chk1" then %>
	<tr bgcolor="#FFFFFF">
		<td style="line-height:25px">
		<strong>1. ���԰� Ȯ�� / �������밡Ȯ�� / �����鼼 / ���԰��Ҽ���</strong>
		<br>&nbsp; - [��������]���޸�����>>���޻� ���� üũ</a>
		: �ٹ����� ���԰� ���� ������ �ǵ� Ȯ��
		<br>
		&nbsp;&nbsp;&nbsp;&nbsp; - <a href="/admin/etc/difforder/buycashErrList.asp?menupos=3956&vTab=2" target="_menu3956_tab">���԰�����üũ</a>
		<br>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : ezwel�� ��� �ǸŰ��� 100�� ������ ������. �������� �ִ�.<br>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : 10x10�� ��� ������ȣ �ִ°�� (�ڻ�δ������ϰ��)�������� ���� �� �ִ�.<br>
		&nbsp;&nbsp;&nbsp;&nbsp; - <a href="/admin/etc/difforder/buycashOverList.asp?menupos=3956&vTab=3" target="_menu3956_tab">�����԰����� ��ǰ���� ���� ���԰��� ū���</a>
		<br>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : ��ǰ�ɼǺ���� �߸� ���°�찡 ��Ȥ �ִ�.<br>
		&nbsp;&nbsp;&nbsp;&nbsp; - <a href="/admin/etc/difforder/taxErrList.asp?menupos=3956&vTab=4" target="_menu3956_tab">�鼼 ����� üũ</a>
		<br>
		&nbsp;&nbsp;&nbsp;&nbsp; - <a href="/admin/etc/difforder/buycashPrimeList.asp?menupos=3956&vTab=5" target="_menu3956_tab">��ǰ/�ɼǰ��ް��Ҽ���</a>
		<br>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : ��ü/��Ź��ǰ�� ��� ���԰��� �Ҽ����� ���� ����ϴ� (�Ϻ���ġ�� ó��)<br>


		<br>
		<strong>2. �ֹ�����, ����Ȯ���� ����</strong>
		<% if (FALSE) then %><br>&nbsp; - �ֹ����� ���� ���� --> ����α� ���信���Ե� <% end if %>
		<br>&nbsp; - <a href="#" onClick="popJungsanCheck('900');return false;">����Ȯ���� ����</a>

		<br>
		<br>
		<strong>3. ����α� ����</strong>
		<br>&nbsp; -  <a href="#" onClick="popJungsanCheck('');return false;">pop</a>
		<br>
	<% if (FALSE) then %>
		<strong>4.</strong>
		<br>

		</td>
	<% end if %>
	</tr>
	<% elseif gubun="chk2" then %>
	<tr bgcolor="#FFFFFF">
		<td style="line-height:25px">
		<strong>1. ���� �����Է� ������ Ȯ��</strong>
		<br>&nbsp; - <a href="/admin/offshop/offshopjumun_error.asp?menupos=1183" target="_menu1183">[OFF]����_������>>���԰� ���� �Ǹų���</a>
		: ��� ������ �Ǹų��� Ȯ��
		<br>
		<br>
		<strong>2. ����α� ����</strong>
		<br>&nbsp; - <a href="#" onClick="popJungsanCheck('200');return false;">pop</a>
	<% if (FALSE) then %>
		<strong>3. ���� �ֹ� ����Ȯ��</strong>
		<br>&nbsp; -
		<br>

		<br>
		<br>
		<strong>3. </strong>
		<br>&nbsp; -
		<br>
	<% end if %>
		<br>
		<br>
		</td>
	</tr>
	<% elseif gubun="act0" then %>
	<tr bgcolor="#FFFFFF">
		<td style="line-height:25px">
		<strong>1. OFF ���� �ϰ�ó��</strong>
		<br>&nbsp; - <a href="/admin/upchejungsan/popjungsanmakebatch.asp?targetGbn=OF&yyyy1=<%=yyyy1%>&mm1=<%=mm1%>" target="_popjungsanmakebatch_onff">OFF �����ۼ�</a>
		<br>
		<br>

		<strong>2. ON ���� �ϰ�ó��</strong>
		<br>&nbsp; - <a href="/admin/upchejungsan/popjungsanmakebatch.asp?targetGbn=ON&yyyy1=<%=yyyy1%>&mm1=<%=mm1%>" target="_popjungsanmakebatch_onff">ON �����ۼ�</a>
		<br>
		<br>
		<strong>3. Class ���� �ϰ�ó��</strong>
		<br>&nbsp; - 2020/03 ���ε� ������ �������� �����
		<br>

		<br>
		<br>
		</td>
	</tr>
	<% elseif gubun="ext0" then %>
	<tr bgcolor="#FFFFFF">
		<td style="line-height:25px">
		<strong>1. �߰� ��ۺ�����(��ü�߰�����)</strong>
		<br>&nbsp; - <a href="/admin/upchejungsan/upchedeliverypay.asp?menupos=975" target="_menu975"> [����]���곻���ۼ�>>[ON]��ۺ�����</a>
		<br>
		<br>

		<strong>2. ��Ÿ�������</strong>
		<br>&nbsp; - <a href="/admin/upchejungsan/etcchulgojungsan.asp?menupos=321&gubun=witakchulgoJS" target="_menu321"> [����]���곻���ۼ�>>[ON]��Ÿ�������</a>
		<br>
		<br>
		<strong>3. �귣������ �������� </strong>
		<form name="frmbrandcpn" method="post" action="dobatch.asp">
		<input type="hidden" name="mode" value="brandcpn">
		<br>
		&nbsp; - ����� <input type="text" name="jyyyymm" value="<%=yyyy1%>-<%=mm1%>" size="5" maxlength="7">
		&nbsp; - ���� <input type="text" name="differencekey" value="0" size="1" maxlength="2">
		&nbsp; - �귣��ID <input type="text" name="makerid" value="<%= "ETUDEHOUSE" %>" size="12" maxlength="32">
		&nbsp; - ��ü�δ��� <input type="text" name="upchepro" value="<%= "70" %>" size="2" maxlength="3" style="text-align:right"> %
		&nbsp; <input type="button" value="���� ���� �ۼ�" onClick="etcBrandCpnJungsan(this)">
		</form>
		<form name="frmbrandcpnidx" method="post" action="dobatch.asp">
		<input type="hidden" name="mode" value="brandcpnidx">
		<br>
		&nbsp; - ����� <input type="text" name="jyyyymm" value="<%=yyyy1%>-<%=mm1%>" size="5" maxlength="7">
		&nbsp; - ���� <input type="text" name="differencekey" value="0" size="1" maxlength="2">
		&nbsp; - �귣��ID <input type="text" name="makerid" value="<%= "ETUDEHOUSE" %>" size="12" maxlength="32">
		&nbsp; - ������ȣ <input type="text" name="cpnidx" value="<%= "7777" %>" size="4" maxlength="32">
		&nbsp; - ��ü�δ��� <input type="text" name="upchepro" value="<%= "70" %>" size="2" maxlength="3" style="text-align:right"> %
		&nbsp; <input type="button" value="���� ���� �ۼ�" onClick="etcBrandCpnIdxJungsan(this)">
		</form>
		<br><br>
		<strong>4. ���Ը������� ��������</strong>
		<br>&nbsp; - <a href="/admin/shopmaster/sale/maeipSaleMarginList.asp?menupos=3967" target="_menu3967"> [ON]��ǰ����>>���� ��������(���Ի�ǰ)</a>
		<br><br>
		<strong>5. ��ۺ�д� ���θ�� �������</strong>
		<br>&nbsp; - <a href="/admin/sitemaster/halfDeliveryPay/index.asp?menupos=4155" target="_menu4155"> [ON]��ǰ����>>��ۺ�δ㼳��</a>
		<br><br>
		<strong>6. ���� ������ۺ� ������</strong>
		<form name="frmextbeasongPay" method="post" action="dobatch.asp">
		<input type="hidden" name="mode" value="addextbeasongPay">
		<br>
		&nbsp; - ���޸� <input type="text" name="sitename" value="" size="12" maxlength="32">
		&nbsp; - ������� <input type="text" name="yyyymmdd" value="" size="10" maxlength="10">
		&nbsp; - ��ۺ� �ݾ� <input type="text" name="beasongPay" value="" size="12" maxlength="32">
        &nbsp; <input type="button" value="��ۺ� ���� ���" onClick="jsAddextBeasongPay(this)">
        </form>
		<br>
		<br>

		<br>
		<br>
		</td>
	</tr>
	<% elseif gubun="chk9" then %>
	<tr bgcolor="#FFFFFF">
		<td style="line-height:25px">
		<strong>1. ����/����α� �� ON</strong>
		<br>&nbsp; - <a href="#" onClick="popJungsanCheck('300');return false;"> pop</a>
		<br>
		<br>

		<strong>2. ����/����α� �� OF</strong>
		<br>&nbsp; - <a href="#" onClick="popJungsanCheck('400');return false;"> pop</a>
		<br>

		<br>
		<br>
		</td>
	</tr>
	<% elseif gubun="act9" then %>
	<tr bgcolor="#FFFFFF">
		<td>
		 <input type="button" class="button" value="������->��üȮ���� �ϰ�ó�� ON" onClick="javascript:dobatch(barchForm,'finishflag1');"><br><br>

		 <input type="button" class="button" value="������->��üȮ���� �ϰ�ó�� OF" onClick="javascript:dobatch(barchForm,'finishflagoff1');"><br><br>


		</td>
	</tr>
	<% end if %>
</table>


<form name="frmArrupdate" method="post" action="dobatch.asp">
<input type="hidden" name="idx" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="yyyy" value="<%= yyyy1 %>">
<input type="hidden" name="mm" value="<%= mm1 %>">
</form>
<%
SET ojungsan = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
