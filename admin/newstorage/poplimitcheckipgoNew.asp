<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/newstoragecls.asp"-->
<%

// 97% -> 98%, 2014-07-25
const C_LIMIT_PERCENT = 1
'const C_LIMIT_PERCENT = 0.98
'//C_LIMIT_PERCENT(�������� ���� ����) 2019-01-03 ������
public Function GetDanJongStat(byval v)
	if v="Y" then
		GetDanJongStat="����"
	elseif v="S" then
		GetDanJongStat="�Ͻ�ǰ��"
	elseif v="M" then
		GetDanJongStat="MDǰ��"
	elseif v="N" then
		GetDanJongStat="������"
	else
	    GetDanJongStat=v
	end if
End Function

dim idx
idx = request("idx")





dim oipchul, oipchuldetail
set oipchul = new CIpChulStorage
oipchul.FRectId = idx
oipchul.GetIpChulMaster

set oipchuldetail = new CIpChulStorage
oipchuldetail.FRectStoragecode = oipchul.FOneItem.Fcode
oipchuldetail.GetIpChulDetailCheck

dim i
dim maylimitno, mayreallimtno, currlimitea
dim addmayno
%>
<script language='javascript'>
function PopItemSellEdit(iitemid){
	var popwin = window.open('/common/pop_simpleitemedit.asp?itemid=' + iitemid,'itemselledit','width=500 height=600')
}

function ReLimitSell(itemid,itemoption,maylimitno,mayreallimtno,currlimitno){
	if (confirm('�԰�� �������: ' + maylimitno + '\n\n�������� [   ' + mayreallimtno + '    ] \n\n�� �����Ͻðڽ��ϱ�?\n\n(����:' + currlimitno + '  �߰�:'+(mayreallimtno-currlimitno) + ')')){
		var popwin = window.open('/admin/newstorage/ipgoitemlimitcheckNew_process.asp?mode=addmaylimit&itemid=' + itemid + '&itemoption=' + itemoption + '&mayreallimtno=' + mayreallimtno,'doipgoitemlimitcheck','width=100 height=100');
	}
}

function ReLimitSellIMSI(itemid,itemoption,oldno,addno){
	if (confirm('�԰�� ��������: \n\ �������� [   ' + (oldno*1 + addno*1) + '    ] \n\n�� �����Ͻðڽ��ϱ�?\n\n(����:' + oldno + '  �߰�:'+(addno) + ')')){
		var popwin = window.open('/admin/newstorage/doipgoitemlimitcheck.asp?mode=addlimitno&itemid=' + itemid + '&itemoption=' + itemoption + '&addno=' + addno,'doipgoitemlimitcheck','width=100 height=100');
	}
}

// ============================================================================
// �ɼǼ���
function editItemOption(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/common/pop_itemoption.asp?' + param ,'editItemOption','width=900,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function fnEditRealStockNo(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;
    var isDasBalju = false;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('������ ���� ����� �����ϴ�.');
		return;
	}

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				upfrm.arrCheckinfo.value = upfrm.arrCheckinfo.value + "|" + frm.cksel.value;
				upfrm.arrStockNo.value = upfrm.arrStockNo.value + "|" + frm.mayreallimtno.value;
				//upfrm.arrLimitYN.value = upfrm.arrLimitYN.value + "|" + frm.limityn.value;
			}
		}
	}
	upfrm.submit();
}

function ckAll(icomp){
	var bool = icomp.checked;
	AnSelectAllFrame(bool);
}
</script>
<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
   	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
	        	�ѰǼ�:  <%= oipchuldetail.FResultCount %>
	        </td>
	        <td valign="top" align="right" bgcolor="F4F4F4">
	        	<input type="button" class="button" value="���� ��������" onClick="fnEditRealStockNo();">&nbsp;&nbsp;<input type="button" class="button" value="���ΰ�ħ" onClick="document.location.reload();">
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="100%" cellpadding="5" cellspacing="1" class="a" bgcolor="#BABABA">
	<tr bgcolor="#DDDDFF" align="center">
		<td width="50"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td width="120">��ǰ�ڵ�</td>
		<td>��ǰ��</td>
		<td>�ɼǸ�</td>
		<td width="50">��������</td>
		<td width="50">�ǸŰ�</td>
		<td width="40">�ֹ�<br>����</td>
		<td width="40">�԰�<br>����</td>
		<td width="50">����<br>�����</td>
		<td width="30">ǰ��</td>
		<td width="30">�Ǹ�</td>
		<td width="30">����</td>
		<td width="50">����<br>����</td>
		<td width="30">����<br>����</td>
		<td width="50">��� </td>
	</tr>
	<% for i=0 to oipchuldetail.FResultCount-1 %>
	<%
		''if (oipchuldetail.FItemList(i).FIsNewItem = "Y") then
		''    maylimitno = oipchuldetail.FItemList(i).FMaystockno + oipchuldetail.FItemList(i).Fitemno ''(�԰����)
		''else
		    maylimitno = oipchuldetail.FItemList(i).FMaystockno
		''end if

		mayreallimtno = Fix(maylimitno*C_LIMIT_PERCENT)

		currlimitea = oipchuldetail.FItemList(i).getOptionLimitEa

	%>
	<form name="frmBuyPrc_<%= oipchuldetail.FItemList(i).Fiitemgubun %>_<%= oipchuldetail.FItemList(i).FItemId %>_<%= oipchuldetail.FItemList(i).FItemOption %>" method="post" >
	<input type="hidden" name="mayreallimtno" value="<%=mayreallimtno%>">
	<tr align=center bgcolor="<%= oipchuldetail.FItemList(i).GetMayCheckColor %>">
		<td height="25">
		<% if (oipchuldetail.FItemList(i).FLimitYn="Y") then %>
		    <% if mayreallimtno<>0 then %>
        		<% if ((mayreallimtno<>currlimitea) or (currlimitea>mayreallimtno)) then %>
				<input type="checkbox" name="cksel" id="cksel" value="<%= oipchuldetail.FItemList(i).Fiitemgubun %>_<%= oipchuldetail.FItemList(i).FItemId %>_<%= oipchuldetail.FItemList(i).FItemOption %>">
				<% else %>
				<input type="checkbox" name="cksel" id="cksel" disabled>
		        <% end if %>
			<% else %>
			<input type="checkbox" name="cksel" id="cksel" disabled>
		    <% end if %>
		<% else %>
			<input type="checkbox" name="cksel" id="cksel" disabled>
		<% end if %>
		</td>
		<td>
			<font color="<%= mwdivColor(oipchuldetail.FItemList(i).FOnlineMwdiv) %>"><%= oipchuldetail.FItemList(i).Fiitemgubun %>-<%= oipchuldetail.FItemList(i).FItemId %>-<%= oipchuldetail.FItemList(i).FItemOption %></font>
		</td>
		<td align="left">
		    <a href="javascript:PopItemSellEdit('<%= oipchuldetail.FItemList(i).FItemID %>');"><%= oipchuldetail.FItemList(i).Fiitemname %></a>
		</td>
		<td <%= chkIIF(oipchuldetail.FItemList(i).FItemOption<>"0000" and oipchuldetail.FItemList(i).FOptUsing="N","bgcolor='#FF3333'","") %>>
		    <%= oipchuldetail.FItemList(i).Fiitemoptionname %>
		</td>
		<td><%= fncolor(oipchuldetail.FItemList(i).FDanJongYn,"dj") %></td>
		<td align=right><%= FormatNumber(oipchuldetail.FItemList(i).Fsellcash,0) %></td>
		<td><%= oipchuldetail.FItemList(i).Fbaljuitemno %></td>
		<td><%= oipchuldetail.FItemList(i).Fitemno %></td>
        <% if (oipchuldetail.FItemList(i).FIsNewItem = "Y") then %>
        <td><%= oipchuldetail.FItemList(i).FMaystockno %></td>
        <% else %>
        <td><%= oipchuldetail.FItemList(i).FMaystockno %></td>
        <% end if %>

		<td><font color="<%= oipchuldetail.FItemList(i).GetIsSlodOutColor %>"><%= oipchuldetail.FItemList(i).GetIsSlodOutText %></font></td>
		<td><font color="<%= oipchuldetail.FItemList(i).GetSellYnColor %>"><%= oipchuldetail.FItemList(i).Fsellyn %></font></td>

		<% if oipchuldetail.FItemList(i).Flimityn="Y" then %>
		<td bgcolor="#FFDDDD"><font color="<%= oipchuldetail.FItemList(i).GetLimitYnColor %>"><%= oipchuldetail.FItemList(i).Flimityn %></font></td>
		<% else %>
		<td><font color="<%= oipchuldetail.FItemList(i).GetLimitYnColor %>"><%= oipchuldetail.FItemList(i).Flimityn %></font></td>
		<% end if %>

		<td>
		<% if oipchuldetail.FItemList(i).FLimitYn="Y" then %>
			<%= oipchuldetail.FItemList(i).getOptionLimitEa %>
		<% end if %>
		</td>
		<td>
		<%=mayreallimtno%>/<%=currlimitea%>
		<% if (oipchuldetail.FItemList(i).FLimitYn="Y") then %>
		    <% if mayreallimtno<>0 then %>
        		<% if ((mayreallimtno<>currlimitea) or (currlimitea>mayreallimtno)) then %>
        		<input type="button" class="button" value="->" onclick="ReLimitSell('<%= oipchuldetail.FItemList(i).FItemID %>','<%= oipchuldetail.FItemList(i).FItemOption %>','<%= maylimitno %>','<%= mayreallimtno %>','<%= currlimitea %>')">
        		<% end if %>
		    <% end if %>
		<% end if %>
		</td>
		<td><% if oipchuldetail.FItemList(i).FDtComment<>"" then%><%= replace(oipchuldetail.FItemList(i).FDtComment," ","<br>") %><% end if %></td>
	</tr>
	</form>
	<% next %>

</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right" bgcolor="F4F4F4">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->
<form name="frmArrupdate" method="post" action="realStockNoEdit_process.asp">
<input type="hidden" name="arrCheckinfo" value="">
<input type="hidden" name="arrStockNo" value="">
<input type="hidden" name="arrLimitYN" value="">
</form>
<%
set oipchuldetail= Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
