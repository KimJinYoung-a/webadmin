<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/upcheitemeditcls.asp"-->

<%
dim designerid, itemid
dim research,notfinish
dim page
page = request("page")
designerid = request("designerid")
itemid = request("itemid")
research = request("research")
notfinish = request("notfinish")

if page="" then page=1

if research<>"on" then
	notfinish = "on"
end if

dim isfinishStr
if notfinish="on" then
	isfinishStr="N"
end if

dim oupcheitemedit
set oupcheitemedit = New CUpCheItemEdit
oupcheitemedit.FPageSize = 20
oupcheitemedit.FCurrPage = page
oupcheitemedit.FRectDesignerID =  designerid
oupcheitemedit.FRectItemId = itemid
oupcheitemedit.FRectNotFinish = isfinishStr

oupcheitemedit.GetReqList

dim i
%>
<script language='javascript'>

function NextPage(page){
	document.frm.page.value = page;
	document.frm.submit();
}

function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500 height=600')
	popwin.focus();
}

function PopItemDetail(itemid, itemoption){
	var popwin = window.open('/admin/stock/itemcurrentstock.asp?itemid=' + itemid + '&itemoption=' + itemoption,'popitemdetail','width=1000, height=600, scrollbars=yes');
	popwin.focus();
}

function SelectCk(opt){
	var bool = opt.checked;
	AnSelectAllFrame(bool)
}

function DelThis2(){

}

function DelThis(frm){
	if (frm.rejectstr.value.length<1){
		alert('�ź� ������ �Է��� �ּ���.');
		frm.rejectstr.focus();
		return;
	}

	var ret = confirm('���� �ź� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.mode.value="del";
		frm.submit();
	}
}

function AccThis(frm){
	frm.mode.value="acct";

	if ((frm.limitSetno)&&(!IsDigit(frm.limitSetno.value))){
		alert('���ڸ� �����մϴ�.');
		frm.limitno.focus();
		return;
	}

//	if (!IsDigit(frm.limitsold.value)){
//		alert('���ڸ� �����մϴ�.');
//		frm.limitno.focus();
//		return;
//	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret){
		frm.submit();
	}
}


// ============================================================================
function ChangeDispYN(frm, divdispyn) {
    if (frm.dispyn.value == "Y") {
        frm.dispyn.value = "N";
        divdispyn.innerHTML = "<font color=red>����</font>";
    } else {
        frm.dispyn.value = "Y";
        divdispyn.innerHTML = "����";
    }
}

function ChangeSellYN(frm, divsellyn) {
    if (frm.sellyn.value == "Y") {
        frm.sellyn.value = "N";
        divsellyn.innerHTML = "<font color=red>ǰ��</font>";
    } else {
        frm.sellyn.value = "Y";
        divsellyn.innerHTML = "�Ǹ�";
    }
}

function ChangeLimitYN(frm, divlimityn) {
    if (frm.limityn.value == "Y") {
        frm.limityn.value = "N";
        divlimityn.innerHTML = "�Ϲ�";
        frm.limitSetno.disabled = true;
        frm.limitSetno.style.background = "#CCCCCC";
    } else {
        frm.limityn.value = "Y";
        divlimityn.innerHTML = "<font color=red>����</font>";
        frm.limitSetno.disabled = false;
        frm.limitSetno.style.background = "#FFFFFF";
    }
}

</script>
 
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�귣�� :
			<% drawSelectBoxDesigner "designerid",designerid %>
			&nbsp;
			��ǰ�ڵ� :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="10" maxlength="9">
			&nbsp;
			<input type="checkbox" name="notfinish" <% if notfinish="on" then response.write "checked" %> >��ó����ϸ�
			<br>
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<!--<input type="button" class="button" value="���þ���������" onClick="alert('�ϰ�ó���� �غ����Դϴ�.');">-->
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- �׼� �� -->

<p>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="20"><input type="checkbox" name="ckall" onClick="SelectCk(this)"></td>
		<td width="50">�̹���</td>
		<td width="50">��ǰ�ڵ�</td>
		<td width="80">�귣��ID</td>
		<td>�����۸�</td>
		<td width="30">�ŷ�<br>����</td>
		<td width="70">�����</td>
		<td width="200">��û����<br>(��û������ �����Ϸ��� ���콺�� Ŭ���ϼ���)</td>
		<td width="100">��û��������</td>
		<td width="30">�ź�</td>
		<td width="30">����</td>
	</tr>
	<% for i=0 to oupcheitemedit.FResultCount -1 %>
	<form name="frmBuyPrc_<%= i %>" method="post" action="do_upche_req_itemmodify.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="idx" value="<%= oupcheitemedit.FItemList(i).Fidx %>">
	<input type="hidden" name="itemid" value="<%= oupcheitemedit.FItemList(i).FItemId %>">
	<input type="hidden" name="itemoption" value="<%= oupcheitemedit.FItemList(i).FItemOption %>">
	<input type="hidden" name="sellyn" value="<%= oupcheitemedit.FItemList(i).FSellYn %>">
	<input type="hidden" name="limityn" value="<%= oupcheitemedit.FItemList(i).FLimitYn %>">
	<tr align="center" bgcolor="#FFFFFF">
		<td rowspan="2"><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td rowspan="2"><img src="<%= oupcheitemedit.FItemList(i).FImageSmall %>" width="50" height="50" ></td>
		<td rowspan="2">
		    <a href="javascript:PopItemSellEdit('<%= oupcheitemedit.FItemList(i).FItemId %>');"><%= oupcheitemedit.FItemList(i).FItemId %></a>
		    <br>
		    (<%= oupcheitemedit.FItemList(i).FItemOption %>)
		</td>
		<td rowspan="2"><%= oupcheitemedit.FItemList(i).FMakerId %></td>
		<td rowspan="2" align="left">
			<%= oupcheitemedit.FItemList(i).FItemName %>
		<% if oupcheitemedit.FItemList(i).FItemOptionName<>"" then %>
			<br><%= oupcheitemedit.FItemList(i).FItemOptionName %>
		<% end if %>
		</td>
		<td rowspan="2"><%= fnColor(oupcheitemedit.FItemList(i).Fmwdiv,"mw") %></td>
		<td><acronym title="<%= oupcheitemedit.FItemList(i).FRegDate %>"><%= left(oupcheitemedit.FItemList(i).FRegDate,10) %></acronym></td>
		<td>
        <% if (oupcheitemedit.FItemList(i).FItemOption = "0000") or (oupcheitemedit.FItemList(i).FItemOption = "XXXX") then %>
        <!-- �ɼ��� ������� -->
        	<table width="100%" border="0" class="a">
        		<tr>
        			<td>
					  	<table width="100%" border="0" class="a">
						    <tr>
						    	<td width="30">�Ǹ�:</td>
							    <td width="40"><%= oupcheitemedit.FItemList(i).GetOldSellYnName %>-&gt;</td>
							    <td width="30"><a href="javascript:ChangeSellYN(frmBuyPrc_<%= i %>, divsellyn<%= i %>)"><div id="divsellyn<%= i %>"><% if (oupcheitemedit.FItemlist(i).FSellYn = "Y") then %>�Ǹ�<% else %><font color=red>ǰ��</font><% end if %></div></a></td>
							    <td>(����:<%= oupcheitemedit.FItemList(i).GetCurrSellYnName %>)</td>
						    </tr>
					  	</table>
					</td>
				</tr>

				<tr>
        			<td>
					  	<table width="100%" border="0" class="a">
						    <tr>
						    	<td width="30">����:</td>
						        <td width="40"><%= oupcheitemedit.FItemList(i).GetOldLimitYnName %>-&gt;</td>
						        <td width="30"><a href="javascript:ChangeLimitYN(frmBuyPrc_<%= i %>, divlimityn<%= i %>)"><div id="divlimityn<%= i %>"><% if (oupcheitemedit.FItemlist(i).FLimitYn = "N") then %>�Ϲ�<% else %><font color=red>����</font><% end if %></div></a></td>
						        <td>(����:<%= oupcheitemedit.FItemList(i).GetCurrLimitYnName %>)</td>
						    </tr>
					    </table>
					</td>
				</tr>

				<tr>
        			<td>
					  	<table width="100%" border="0" class="a">
						    <tr>
						        <td>
						        	������������:<input type="text" class="text" name="limitSetno" value="<%= oupcheitemedit.FItemList(i).FLimitNo %>" size="2" <%= chkIIF(oupcheitemedit.FItemlist(i).FLimitYn= "N","disabled style='background-color:#CCCCCC'","") %> >
						     	</td>
						    </tr>
					    </table>
					</td>
				</tr>
		  </table>
		<% else %>
		<!-- �ɼ��� �������  -->
        	<table width="100%" border="0" class="a">
        		<tr>
        			<td>
					  	<table width="100%" border="0" class="a">
						    <tr>
						    	<td width="30"><strong>�ɼǻ��</strong>:</td>
							    <td width="40"><%= oupcheitemedit.FItemList(i).GetOldOptUsingYnName %>-&gt;</td>
							    <td width="30"><a href="javascript:ChangeSellYN(frmBuyPrc_<%= i %>, divsellyn<%= i %>)"><div id="divsellyn<%= i %>"><% if (oupcheitemedit.FItemlist(i).FSellYn = "Y") then %>�Ǹ�<% else %><font color=red>ǰ��</font><% end if %></div></a></td>
							    <td>(����:<%= oupcheitemedit.FItemList(i).GetCurrSellYnName %>)</td>
						    </tr>
					  	</table>
					</td>
				</tr>
				<!--
				<tr>
        			<td>
					  	<table width="100%" border="0" class="a">
						    <tr>
						    	<td width="30">����:</td>
						        <td width="40"><%= oupcheitemedit.FItemList(i).GetOldLimitYnName %>-&gt;</td>
						        <td width="30"><a href="javascript:ChangeLimitYN(frmBuyPrc_<%= i %>, divlimityn<%= i %>)"><div id="divlimityn<%= i %>"><% if (oupcheitemedit.FItemlist(i).FLimitYn = "N") then %>�Ϲ�<% else %><font color=red>����</font><% end if %></div></a></td>
						        <td>(����:<%= oupcheitemedit.FItemList(i).GetCurrLimitYnName %>)</td>
						    </tr>
					    </table>
					</td>
				</tr>
				-->
				<% if oupcheitemedit.FItemList(i).FLimitYn="Y" then %>
				<tr>
        			<td>
					  	<table width="100%" border="0" class="a">
						    <tr>
						        <td>
						        	������������:<input type="text" class="text" name="limitSetno" value="<%= oupcheitemedit.FItemList(i).FLimitNo %>" size="2">
						     	</td>
						    </tr>
					    </table>
					</td>
				</tr>
				<% end if %>
		  </table>

        <% end if %>
		</td>

		<td>
			<%= oupcheitemedit.FItemList(i).FLimitNo %>
			-
			<%= oupcheitemedit.FItemList(i).FLimitSold %>
			=
			<%= oupcheitemedit.FItemList(i).GetRemainEa %>
			<br>
			(����:<%= oupcheitemedit.FItemList(i).GetCurrRemainEa %>)
     		<br>
     		(�����:<%= oupcheitemedit.FItemList(i).GetLimitStockNo %>)
			<!-- (<%= oupcheitemedit.FItemList(i).FCurrLimitNo %>-<%= oupcheitemedit.FItemList(i).FCurrLimitSold %>) -->
		</td>
		<td rowspan="2"><a href="javascript:DelThis(frmBuyPrc_<%= i %>)">�ź�</a></td>
		<td rowspan="2"><a href="javascript:AccThis(frmBuyPrc_<%= i %>)">����</a></td>
	</tr>
	<tr>
		<td colspan="3" bgcolor="#FFFFFF">
			��û���� : <%= oupcheitemedit.FItemList(i).FEtcStr %><br>
			�źλ��� : <input type="text" class="text" name="rejectstr" value="" size="36" maxlength="64">
		</td>
	</tr>
	</form>
	<% next %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if oupcheitemedit.HasPreScroll then %>
        		<a href="javascript:NextPage('<%= oupcheitemedit.StartScrollPage-1 %>')">[pre]</a>
        	<% else %>
        		[pre]
        	<% end if %>

        	<% for i=0 + oupcheitemedit.StartScrollPage to oupcheitemedit.FScrollCount + oupcheitemedit.StartScrollPage - 1 %>
        		<% if i>oupcheitemedit.FTotalpage then Exit for %>
        		<% if CStr(page)=CStr(i) then %>
        		<font color="red">[<%= i %>]</font>
        		<% else %>
        		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
        		<% end if %>
        	<% next %>

        	<% if oupcheitemedit.HasNextScroll then %>
        		<a href="javascript:NextPage('<%= i %>')">[next]</a>
        	<% else %>
        		[next]
        	<% end if %>
		</td>
	</tr>
</table>


<%
set oupcheitemedit = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->