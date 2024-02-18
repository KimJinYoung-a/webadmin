<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ������߼� ��ǰ ���
' History : 2020.11.24 ������ ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/baljucls.asp"-->
<%
dim orderserial,obalju,inputyn, frommakerid, tomakerid, ordby, i, ttlcount
dim makerid, warehouseCd
dim dplusFrom, dplusTo
    frommakerid 	= requestCheckVar(request("frommakerid"), 32)
    tomakerid 		= requestCheckVar(request("tomakerid"), 32)
    makerid 		= requestCheckVar(request("makerid"), 32)
    ordby 			= requestCheckVar(request("ordby"), 32)
    warehouseCd		= requestCheckVar(request("warehouseCd"), 32)
    dplusFrom		= requestCheckVar(request("dplusFrom"), 32)
    dplusTo			= requestCheckVar(request("dplusTo"), 32)

ttlcount=0

if (ordby = "") then
	ordby = "itemid"
end if

set obalju = New CBalju
	obalju.FStartdate = dateSerial(year(now()),month(now())-6,day(now()))
	obalju.FRectFromMakerid = frommakerid
	obalju.FRectToMakerid = tomakerid
    obalju.FRectMakerid = makerid
	obalju.FRectOrderBy = ordby
    obalju.FRectWarehouseCd = warehouseCd
    if IsNumeric(dplusFrom) then
        obalju.FRecDPlusFrom = dplusFrom
    end if
    if IsNumeric(dplusTo) then
        obalju.FRecDPlusTo = dplusTo
    end if

	obalju.GetMissingReSendOrderDetailAll

%>

<script type='text/javascript'>
function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500 height=600');
	popwin.focus();
}


function popResendJumunByItem(makerid, itemid, itemoption) {
	var popwin = window.open('resendmaster_top.asp?designer=' + makerid + '&itemid=' + itemid + '&itemoption=' + itemoption,'popmisendjumunbyitemid','width=1600 height=600 scrollbars=yes resizable=yes')
    popwin.focus();
}

function poppreorder(barcode, makerid){
	<%
	dim preorderdate
		preorderdate = dateadd("m",-3,date)
	%>
	var poppreorder = window.open('/admin/newstorage/orderlist.asp?menupos=537&barcode='+barcode+'&designer='+makerid+'&yyyy1=<%= year(preorderdate) %>&mm1=<%= Format00(2,month(preorderdate)) %>&dd1=<%= Format00(2,day(preorderdate)) %>&yyyy2=<%= year(date) %>&mm2=<%= Format00(2,month(date)) %>&dd2=<%= Format00(2,day(date)) %>&statecd=preorder','poppreorder','width=1280,height=960,scrollbars=yes,resizable=yes');
	poppreorder.focus();
}
</script>

<style type="text/css">
td,select,input { font-size:9pt; font-family:Verdana;}

.button {
	font-family: "����", "����";
	font-size: 10px;
	background-color: #E4E4E4;
	border: 1px solid #000000;
	color: #000000;
	height: 20px;
}
</style>
</div>

<!-- �˻� ���� -->
<form name="frm" method="get" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left" height="25">
        * �귣�� : <% drawSelectBoxDesignerwithName "makerid", makerid %>
        &nbsp;
		* �귣�� ù���� :
		<input type="text" class="text" name="frommakerid" value="<%= frommakerid %>" size="1" maxlength="1">
		~
		<input type="text" class="text" name="tomakerid" value="<%= tomakerid %>" size="1" maxlength="1">
		(�� : 0 ~ h)
        &nbsp;
		* D+Day(��������� ����) :
		<input type="text" class="text" name="dplusFrom" value="<%= dplusFrom %>" size="4" maxlength="4">
		~
		<input type="text" class="text" name="dplusTo" value="<%= dplusTo %>" size="4" maxlength="4">
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>"><input type="button" class="csbutton" value=" �˻� " onClick="document.frm.submit();"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" height="25">
		* ���ļ���:
		<input type="radio" name="ordby" value="itemid" <% if (ordby = "itemid") then %>checked<% end if %> > ��ǰ�ڵ�
		<input type="radio" name="ordby" value="makerid" <% if (ordby = "makerid") then %>checked<% end if %> > �귣��
		<input type="radio" name="ordby" value="rackcode" <% if (ordby = "rackcode") then %>checked<% end if %> > ����ȣ
        <input type="radio" name="ordby" value="ordercnt" <% if (ordby = "ordercnt") then %>checked<% end if %> > �ֹ��Ǽ�
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" height="25">
        * �����ġ :
        <select class="select" name="warehouseCd">
            <option value=""></option>
            <option value="AGV" <%= CHKIIF(warehouseCd="AGV", "selected", "") %>>AGV</option>
            <option value="BLK" <%= CHKIIF(warehouseCd="BLK", "selected", "") %>>BLK</option>
        </select>
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<p>

* ȸ��ǥ�� : ������ϼ����� �ǻ���� ����ġ<br />
* �ֹ��Ǽ� : ������߼� �ֹ��Ǽ�
</p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�귣��ID</td>
	<td width="80">��ǰ<br>�ڵ�</td>
	<td width="40">�ɼ�</td>
	<td width="70">���ڵ�</td>
	<td width="70">�������ڵ�</td>
	<td width="50">���ó</td>
	<td width="50">�̹���</td>
	<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
	<td width="35">�ֹ�<br />�Ǽ�</td>
	<td width="35">����<br />����</td>
	<td width="35">�ǻ�<br />���</td>
	<td width="35">AGV<br />���</td>
	<td width="60">���ֹ�</td>
	<td width="40">����<br>�ֹ�</td>
</tr>

<% for i=0 to Ubound(obalju.FBaljuDetailList) -1 %>
<%
ttlcount = ttlcount + obalju.FBaljuDetailList(i).FItemNo
%>
<% if obalju.FBaljuDetailList(i).FItemNo < obalju.FBaljuDetailList(i).Frealstock then %>
<tr bgcolor="<%= adminColor("dgray") %>" align="center">
<% else %>
<tr bgcolor="FFFFFF" align="center">
<% end if %>
	<td><%= obalju.FBaljuDetailList(i).Fmakerid %></td>
	<% if obalju.FBaljuDetailList(i).IsUpcheBeasong then %>
	<td><font color="red"><%= obalju.FBaljuDetailList(i).FItemID %></font></td>
	<% else %>
	<td><a href="javascript:PopItemSellEdit('<%= obalju.FBaljuDetailList(i).FItemID %>');"><%= obalju.FBaljuDetailList(i).FItemID %></a></td>
	<% end if %>
	<td><%= obalju.FBaljuDetailList(i).FItemOption %></td>
	<td><%= obalju.FBaljuDetailList(i).FItemRackCode %></td>
	<td><%= obalju.FBaljuDetailList(i).FItemsubrackcode %></td>
	<td><%= obalju.FBaljuDetailList(i).FwarehouseCd %></td>
	<td><img src="<%= obalju.FBaljuDetailList(i).FImageSmall %>" width="50" height="50"></td>
	<td align="left">
		<a href="/admin/stock/itemcurrentstock.asp?itemid=<%= obalju.FBaljuDetailList(i).FItemID %>&itemoption=<%= obalju.FBaljuDetailList(i).FItemOption %>" target=_blank ><%= obalju.FBaljuDetailList(i).FItemName %></a>
		<br>
		<% if obalju.FBaljuDetailList(i).FItemOptionName<>"" then %>
		<font color="blue">[<%= obalju.FBaljuDetailList(i).FItemOptionName %>]</font>
		<% end if %>
	</td>
	<td>
		<% if (obalju.FBaljuDetailList(i).Fordercnt <> 1) then %><font color="red"><% end if %>
		<%= obalju.FBaljuDetailList(i).Fordercnt %>
	</td>
	<td><%= obalju.FBaljuDetailList(i).FItemNo %></td>
	<td><%= obalju.FBaljuDetailList(i).Frealstock %></td>
	<td><%= obalju.FBaljuDetailList(i).Fagvstock %></td>
	<input type="hidden" id="itemgubun<%= i %>" name="itemgubun<%= i %>" value="10">
	<input type="hidden" id="itemid<%= i %>" name="itemid<%= i %>" value="<%= obalju.FBaljuDetailList(i).FItemID %>">
	<input type="hidden" id="itemoption<%= i %>" name="itemoption<%= i %>" value="<%= obalju.FBaljuDetailList(i).FItemOption %>">
	<td>
		<a href="#" onclick="poppreorder('<%= "10" & Format00(8,obalju.FBaljuDetailList(i).FItemID) & obalju.FBaljuDetailList(i).FItemOption %>','<%= obalju.FBaljuDetailList(i).Fmakerid %>'); return false;">
		<%= obalju.FBaljuDetailList(i).Fpreorderno %>
		<% if (obalju.FBaljuDetailList(i).Fpreorderno<>obalju.FBaljuDetailList(i).Fpreordernofix) then %>
			-&gt; <%= obalju.FBaljuDetailList(i).Fpreordernofix %>
		<% end if %>
		</a>
	</td>
	<td align=center><a href="" onClick="popResendJumunByItem('<%= obalju.FBaljuDetailList(i).Fmakerid %>', '<%= obalju.FBaljuDetailList(i).FItemID %>', '<%= obalju.FBaljuDetailList(i).FItemOption %>')">����</a></td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="17" align="left">
		�� �Է� ��ǰ�� : <b><%= ttlcount %></b>
	</td>
</tr>
</table>
</form>

<%
set obalju = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
