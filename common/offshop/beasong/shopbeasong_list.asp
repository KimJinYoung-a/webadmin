<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : �������� ���
' Hieditor : 2011.02.22 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/checkPoslogin.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<% '<!-- #include virtual="/lib/checkAllowIPWithLog.asp" --> %>
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->

<%
dim showshopselect, loginidshopormaker ,research, checkblock, reqhp
dim ojumun , i , orderno , ipkumdiv , page
	page = requestcheckvar(request("page"),10)
	orderno = requestcheckvar(request("orderno"),16)
	ipkumdiv = requestcheckvar(request("ipkumdiv"),2)
	research = requestcheckvar(request("research"),10)
	research = requestcheckvar(request("research"),10)
	reqhp = requestcheckvar(request("reqhp"),16)

if page = "" then page = 1
if (research = "") then
	ipkumdiv = 2	'"99"
end if

checkblock = false
showshopselect = false
loginidshopormaker = ""

if C_ADMIN_USER then
	showshopselect = true
	loginidshopormaker = request("shopid")
elseif (C_IS_SHOP) then
	'����/������
	loginidshopormaker = C_STREETSHOPID
else
	if (C_IS_Maker_Upche) then
		loginidshopormaker = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
			loginidshopormaker = "--"		'ǥ�þ��Ѵ�. ����.
		else
			showshopselect = true
			loginidshopormaker = request("shopid")
		end if
	end if
end if

set ojumun = new cupchebeasong_list
	ojumun.FPageSize = 50
	ojumun.FCurrPage = page
	ojumun.frectorderno = orderno
	ojumun.frectipkumdiv = ipkumdiv
	ojumun.frectshopid = loginidshopormaker
	ojumun.frectreqhp = reqhp
	ojumun.fbeagsong_list()

%>

<script type="text/javascript">

	//���ε�� ����Ʈ
	function getOnload(){
	    frm.orderno.select();
	    frm.orderno.focus();
	}

	window.onload = getOnload;

	//������
	function gosubmit(page){
		frm.page.value=page;
		frm.action='/common/offshop/beasong/shopbeasong_list.asp';
		frm.submit();
	}

	//�ֹ�����
	function jumundetail(masteridx, orderno){
		//frmdetail.masteridx.value=masteridx;
		frmdetail.orderno.value=orderno;
		frmdetail.action='/common/offshop/beasong/shopbeasong_input.asp';
		frmdetail.submit();
	}

	//��ü ��� �뺸
	function beasonginput(upfrm){
		frminfo.masteridxarr.value='';
		frminfo.ordernoarr.value='';

		if (!CheckSelected()){
			alert('���þ������� �����ϴ�.');
			return;
		}

		<% if C_ADMIN_AUTH and not(C_logics_Part) then %>
			if (confirm('[�����ڱ���]���ִ� �������� ���� �մϴ�. ��� �����Ͻðڽ��ϱ�?')!=true){
				return;
			}

		<% elseif not(C_logics_Part) then %>
			alert('���ִ� �������� ���� �մϴ�.');
			return;
		<% end if %>

		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.masteridxarr.value = upfrm.masteridxarr.value + frm.masteridx.value + "," ;
					upfrm.ordernoarr.value = upfrm.ordernoarr.value + frm.orderno.value + "," ;
				}
			}
		}

		if (confirm('��� �뺸�� �Ͻðڽ��ϱ�?')){
			frminfo.mode.value='beasonginput';
			frminfo.action='/common/offshop/beasong/shopbeasong_process.asp';
			frminfo.submit();
		}
	}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page">
<input type="hidden" name="masteridx">
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* ����ID :
		<% if (showshopselect = true) then %>
			<% 'drawSelectBoxOffShop "shopid",loginidshopormaker %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",loginidshopormaker, "21") %>
		<% else %>
			<%= loginidshopormaker %>
		<% end if %>
		&nbsp;&nbsp;
		* �ֹ���ȣ : <input type="text" name="orderno" value="<%= orderno %>" size="16" onKeyPress="if(window.event.keyCode==13) gosubmit('');">
		&nbsp;&nbsp;
		* �޴�����ȣ : <input type="text" name="reqhp" value="<%= reqhp %>" size=16 maxlength=16 onKeyPress="if(window.event.keyCode==13) gosubmit('');">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="gosubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* ��ۻ��� : <% drawshopIpkumDivName "ipkumdiv",ipkumdiv," onchange=gosubmit('');" %>
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<br>

<form name="frminfo" method="post" style="margin:0px;">
<input type="hidden" name="mode">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="masteridxarr">
<input type="hidden" name="ordernoarr">
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<input type="button" value="�����ֹ�����뺸" class="button" onclick="beasonginput(frminfo);">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->
</form>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= ojumun.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %>/ <%= ojumun.FTotalPage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
	<td>IDX</td>
	<td>�ֹ���ȣ</td>
	<td>����ID</td>
	<td>�����</td>
	<td>������</td>
	<td width=110>�������޴�����ȣ</td>
	<td>��ۿ�û��</td>
	<td>�����</td>
	<td>��ۻ���</td>
	<td>���</td>
</tr>
<% if ojumun.FresultCount>0 then %>
<%
for i=0 to ojumun.FresultCount-1

checkblock=false
' ����� �Է¿Ϸ� ���� ���� ��� �뺸�� �Ǽ��� �ȵ�
if ojumun.FItemList(i).fipkumdiv < 2 then
	checkblock=true
end if

' �̹� ����뺸 ���� ��� ����
if ojumun.FItemList(i).fipkumdiv >= 5 then
	checkblock=true
end if
%>
<form action="" name="frmBuyPrc<%=i%>" method="get" style="margin:0px;">
<input type="hidden" name="orderno" value="<%= ojumun.FItemList(i).forderno %>">
<input type="hidden" name="masteridx" value="<%= ojumun.FItemList(i).fmasteridx %>">
<tr align="center" bgcolor="#FFFFFF" onmouseover=this.style.background="#f1f1f1"; onmouseout=this.style.background='#ffffff';>
	<td>
		<input type="checkbox" name="cksel" onClick="AnCheckClick(this);" <% if checkblock then %>disabled<% end if %> >
	</td>
	<td>
		<%= ojumun.FItemList(i).fmasteridx %>
	</td>
	<td>
		<%= ojumun.FItemList(i).forderno %>
	</td>
	<td>
		<%=ojumun.FItemList(i).fshopid %>
	</td>
	<td>
		<%=ojumun.FItemList(i).fshopname %>
	</td>
	<td>
		<%= ojumun.FItemList(i).freqname %>
	</td>
	<td>
		<%= ojumun.FItemList(i).freqhp %>
	</td>
	<td>
		<acronym title="<%= ojumun.FItemList(i).fregdate %>"><%= Left(ojumun.FItemList(i).fregdate, 10) %></acronym>
	</td>
	<td>
		<acronym title="<%= ojumun.FItemList(i).fbeadaldate %>"><%= Left(ojumun.FItemList(i).fbeadaldate, 10) %></acronym>
	</td>
	<td>
		<font color="<%= ojumun.FItemList(i).shopIpkumDivColor %>">
		<%= ojumun.FItemList(i).shopIpkumDivName %>
		</font>
	</td>
	<td>
		<input type="button" onclick="jumundetail('<%= ojumun.FItemList(i).fmasteridx %>','<%= ojumun.FItemList(i).forderno %>');" value="�ֹ�����" class="button">
	</td>
</tr>
</form>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
       	<% if ojumun.HasPreScroll then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= ojumun.StartScrollPage-1 %>');">[pre]</a></span>
		<% else %>
		[pre]
		<% end if %>
		<% for i = 0 + ojumun.StartScrollPage to ojumun.StartScrollPage + ojumun.FScrollCount - 1 %>
			<% if (i > ojumun.FTotalpage) then Exit for %>
			<% if CStr(i) = CStr(ojumun.FCurrPage) then %>
			<span class="page_link"><font color="red"><b><%= i %></b></font></span>
			<% else %>
			<a href="javascript:gosubmit('<%= i %>');" class="list_link"><font color="#000000"><%= i %></font></a>
			<% end if %>
		<% next %>
		<% if ojumun.HasNextScroll then %>
			<span class="list_link"><a href="javascript:gosubmit('<%= i %>');">[next]</a></span>
		<% else %>
		[next]
		<% end if %>
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="20" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>

<form name="frmdetail" method="get" action="">
<input type="hidden" name="masteridx" value="">
<input type="hidden" name="orderno" value="">
<input type="hidden" name="menupos" value="<%= menupos %>">
</form>
<%
set ojumun = nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->