<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/shortagestockcls.asp"-->
<%

const C_STOCK_DAY=7

dim IsAvailDelete

dim yyyy1,yyyy2,mm1,mm2,dd1,dd2, nowdate, iStartDate, iEndDate
dim page
dim makerid
dim onlyusing,onlysell,onlyoptionusing, research
dim preorderinclude
dim hanjungsoldout
dim danjongnotinclude
dim mdsoldoutnotinclude
dim soldoutover7days

makerid = request("makerid")
page = request("page")
if page="" then page=1
onlyusing = request("onlyusing")
onlysell = request("onlysell")
onlyoptionusing = request("onlyoptionusing")
research = request("research")
preorderinclude = request("preorderinclude")
hanjungsoldout = request("hanjungsoldout")
danjongnotinclude = request("danjongnotinclude")
mdsoldoutnotinclude = request("mdsoldoutnotinclude")
soldoutover7days = request("soldoutover7days")

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")


if (yyyy1="") then
    nowdate = Left(CStr(DateAdd("d",now(),-2)),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
    
    nowdate = Left(CStr(DateAdd("d",now(),+2)),10)
	yyyy2 = Left(nowdate,4)
	mm2   = Mid(nowdate,6,2)
	dd2   = Mid(nowdate,9,2)
end if

iStartDate  = Left(CStr(DateSerial(yyyy1,mm1,dd1)),10)
iEndDate    = Left(CStr(DateSerial(yyyy2,mm2,dd2)),10)

if (research="") then
	if onlyusing="" then onlyusing="Y"
	'if onlysell="" then onlysell=""
	'if onlyoptionusing="" then onlyoptionusing="Y"
	'if preorderinclude="" then preorderinclude="Y"
	'if danjongnotinclude="" then danjongnotinclude="Y"
	'if mdsoldoutnotinclude="" then mdsoldoutnotinclude="Y"
	'if soldoutover7days="" then soldoutover7days=""
	'if hanjungsoldout="" then hanjungsoldout="Y"
end if

dim ostock
set ostock = new CShortageStock
ostock.FCurrPage = page
ostock.Fpagesize=500
ostock.FRectMakerid = makerid
ostock.FRectStartDate = iStartDate
ostock.FRectEndDate = iEndDate
ostock.FRectOnlyUsing = onlyusing
ostock.FRectOnlySell = onlysell
ostock.FRectOnlyOptionUsing = onlyoptionusing
ostock.FRectpreorderinclude = preorderinclude
ostock.FRectdanjongnotinclude = danjongnotinclude
ostock.FRectmdsoldoutnotinclude = mdsoldoutnotinclude
ostock.FRectsoldoutover7days = soldoutover7days
ostock.FRectSkipLimitSoldOut = hanjungsoldout

ostock.GetTempSoldOutOrderList

dim i


%>

<script language='javascript'>
function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ChangeReqDay(frm){
	if (!(IsDigit(frm.maxsellday.value))){
		alert('���ڸ� �����մϴ�.');
		return;
	}

	if (confirm('�ʿ� ��� �������� �����Ͻðڽ��ϱ�?')){
		frm.submit();
	}
}

function Research(page){
	document.frm.page.value= page;
	document.frm.submit();
}

function DeleteStockLog(itemgubun,itemid,itemoption){
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        frmdelstock.target="_blank";
        frmdelstock.itemgubun.value = itemgubun;
        frmdelstock.itemid.value = itemid;
        frmdelstock.itemoption.value = itemoption;
        frmdelstock.submit();
    }
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�귣�� : <% drawSelectBoxDesignerwithName "makerid", makerid %>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    ���԰� ������ : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		    &nbsp;
			<input type=checkbox name="onlyusing" value="Y" <% if onlyusing="Y" then response.write "checked" %> >����ǰ��
			<!--
			<input type=checkbox name="onlysell" value="Y" <% if onlysell="Y" then response.write "checked" %> >�ǸŻ�ǰ��
			<input type=checkbox name="onlyoptionusing" value="Y" <% if onlyoptionusing="Y" then response.write "checked" %> >���ɼǸ�
			-->
			<input type=checkbox name="Preorderinclude" value="Y" <% if preorderinclude="Y" then response.write "checked" %> >���ֹ�����
			
			<input type=checkbox name="danjongnotinclude" value="Y" <% if danjongnotinclude="Y" then response.write "checked" %> >��������
			<!--
			<input type=checkbox name="mdsoldoutnotinclude" value="Y" <% if mdsoldoutnotinclude="Y" then response.write "checked" %> >MDǰ������	
			<input type=checkbox name="soldoutover7days" value="S" <% if soldoutover7days="S" then response.write "checked" %> >����������
			<input type=checkbox name="hanjungsoldout" value="Y" <% if hanjungsoldout="Y" then response.write "checked" %> >�����Ǹ���������
			-->
	    </td>
	</tr>
	</form>
</table>

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<form name="frmshortage" method=post action="doshortagestock.asp">
	<input type="hidden" name="mode" value="maxsellday">
	<tr>
		<td align="left">
			<!--
			<input type="text" class="text" name="maxsellday" size="2" value="" maxlength=2>�� ��������
			<input type="button" class="button" value="����" onClick="ChangeReqDay(frmshortage);">
			-->
		</td>
		<td align="right">

		</td>
	</tr>
	</form>
</table>
<!-- �׼� �� -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			�˻���� : <b><%= ostock.FResultCount %></b>
			&nbsp;
			(�ִ�˻��Ǽ� : <%= ostock.Fpagesize %>)
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>�귣��ID</td>
		<td width="40">��ǰ<br>�ڵ�</td>
		<td width="50">�̹���</td>
		<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
		<td width="40">����</td>
		<td width="30">�԰�<br>��ǰ<br>(B)</td>
		<td width="30">ON<br>�Ǹ�<br>(D)</td>
		<td width="30">OFF<br>���<br>(C)</td>
		<td width="30">��Ÿ<br>���<br>(C)</td>
		<td width="30">CS<br>���<br>(C)</td>
		<td width="30">����<br>�ҷ�<br>(S)</td>
		<td width="30">����<br>����<br>(E)</td>
		<td width="30" bgcolor="#F3F3FF"><b>�ǻ�<br>��ȿ<br>���<br>(V)</b></td>

		<td width="40">ON(7)<br>�Ǹ�</td>
		<td width="40">OFF(7)<br>�Ǹ�</td>

		<td width="30" bgcolor="#F3F3FF"><b>��(<%= C_STOCK_DAY %>)<br>�ʿ�<br>����</b></td>
		<td width="30">�������<br>�ʿ���� <!-- OFF<br>�ֹ� --></td>
		<td width="30" bgcolor="#F3F3FF"><b>����<br>����</b></td>
		<td width="80">���</td>
	</tr>
<% for i=0 to ostock.FResultCount -1 %>
<%
    IsAvailDelete = (ostock.FItemList(i).Ftotipgono=0) and (ostock.FItemList(i).FtotSellNo=0) and (ostock.FItemList(i).Fshortageno=0) and (ostock.FItemList(i).Frealstock=0) and (ostock.FItemList(i).Fpreorderno=0)
%>

	<% if ostock.FItemList(i).IsInvalidOption then %>
	<tr align="center" bgcolor="#CCCCCC">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
		<td><a href="/admin/newstorage/orderinput.asp?suplyer=<%= ostock.FItemList(i).FMakerID %>" target="iorderinput"><%= ostock.FItemList(i).FMakerID %></a></td>
		<td><a href="javascript:PopItemSellEdit('<%= ostock.FItemList(i).FItemID %>');"><%= ostock.FItemList(i).FItemID %></a></td>
    	<td width="50" align=center><img src="<%= ostock.FItemList(i).Fimgsmall %>" width=50 height=50></td>
		<td align="left">
			<a href="/admin/stock/itemcurrentstock.asp?itemid=<%= ostock.FItemList(i).FItemID %>&itemoption=<%= ostock.FItemList(i).FItemOption %>" target=_blank ><%= ostock.FItemList(i).FItemName %></a>
			<% if ostock.FItemList(i).FItemOption <> "0000" then %>
				<% if ostock.FItemList(i).Foptionusing="Y" then %>
					<br><font color="blue">[<%= ostock.FItemList(i).FItemOptionName %>]</font>
				<% else %>
					<br><font color="#AAAAAA">[<%= ostock.FItemList(i).FItemOptionName %>]</font>
				<% end if %>
			<% end if %>
		</td>
		<td>
			<font color="<%= ostock.FItemList(i).getMwDivColor %>"><%= ostock.FItemList(i).getMwDivName %></font><br>
			<% if ostock.FItemList(i).Fbuycash<>0 then %>
			<%= 100-(CLng(ostock.FItemList(i).Fbuycash/ostock.FItemList(i).Fsellcash*10000)/100) %> %
			<% end if %>
		</td>
		<td><%= ostock.FItemList(i).Ftotipgono %></td>
		<td><%= ostock.FItemList(i).FtotSellNo %></td>
		<td><%= ostock.FItemList(i).Foffchulgono + ostock.FItemList(i).Foffrechulgono %></td>
		<td><%= ostock.FItemList(i).Fetcchulgono + ostock.FItemList(i).Fetcrechulgono %></td>
		<td></td>
		<td><%= ostock.FItemList(i).Ferrbaditemno %></td>
		<td>
			<% if ostock.FItemList(i).Ferrrealcheckno<0 then %>
			<font color="#cc3333"><%= ostock.FItemList(i).Ferrrealcheckno %></font>
			<% else %>
				<%= ostock.FItemList(i).Ferrrealcheckno %>
			<% end if %>
		</td>
		<td bgcolor="#F3F3FF"><b><%= ostock.FItemList(i).Frealstock %></b></td>

		<td><%= ostock.FItemList(i).Fsell7days %></td>
		<td><%= ostock.FItemList(i).Foffchulgo7days %></td>

		<td bgcolor="#F3F3FF"><b><%= ostock.FItemList(i).Frequireno %></b></td>
		<td>
		    <!-- ������� �ʿ���� -->
		    <%= (ostock.FItemList(i).Fipkumdiv5 + ostock.FItemList(i).Foffconfirmno+ostock.FItemList(i).Fipkumdiv4 + ostock.FItemList(i).Fipkumdiv2 + ostock.FItemList(i).Foffjupno)*-1 %>
		</td>
		<td bgcolor="#F3F3FF"><b><%= ostock.FItemList(i).Fshortageno %></b></td>
		<td>
			<%= ostock.FItemList(i).FreipgoMayDate %><br>
		
			<%= fnColor(ostock.FItemList(i).Fdanjongyn,"dj") %>
			<br>
			<% if ostock.FItemList(i).Foptionusing="N" then %>
			<font color="red">�ɼ�x</font><br>
			<% end if %>
			<% if ostock.FItemList(i).IsSoldOut then %>
			<font color="red">ǰ��</font><br>
			<% end if %>
			<% if ostock.FItemList(i).Flimityn="Y" then %>
			<font color="blue">����(<%= ostock.FItemList(i).GetLimitStr %>)</font><br>
			<% end if %>
	
			<% if ostock.FItemList(i).Fpreorderno<>0 then %>
			���ֹ�:
	    		<% if ostock.FItemList(i).Fpreorderno<>ostock.FItemList(i).Fpreordernofix then response.write "</br>" + CStr(ostock.FItemList(i).Fpreorderno) + " -> " %>
			<%= ostock.FItemList(i).Fpreordernofix %>
			<% end if %>
			
			<% if (False) and IsAvailDelete then %>
			<a href="javascript:DeleteStockLog('10','<%= ostock.FItemList(i).FItemID %>','<%= ostock.FItemList(i).FItemOption %>');"><img src="/images/icon_delete.gif" border="0"></a>
			<% end if %>
		</td>
	</tr>
<% next %>
</table>



<!-- ǥ �ϴܹ� ����-->
<!--
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
        	<% if ostock.HasPreScroll then %>
		<a href="javascript:Research('<%= ostock.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ostock.StartScrollPage to ostock.FScrollCount + ostock.StartScrollPage - 1 %>
			<% if i>ostock.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:Research('<%= i %>');">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ostock.HasNextScroll then %>
			<a href="javascript:Research('<%= i %>');">[next]</a>
		<% else %>
			[next]
		<% end if %>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
-->
<!-- ǥ �ϴܹ� ��-->

<%
set ostock = Nothing
%>
<form name="frmdelstock" method="post" action="dostockrefresh.asp">

<input type="hidden" name="mode" value="dellog">
<input type="hidden" name="itemgubun">
<input type="hidden" name="itemid">
<input type="hidden" name="itemoption">
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->