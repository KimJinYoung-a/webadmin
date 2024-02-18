<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ���̘� �÷��� ��ǰ
' History : 2010.11.09 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/DIYitemCls.asp"-->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/PlusDIYItemCls.asp"-->
<%
dim itemid, page , oitem ,oitemoption ,i ,oPsItemList
	itemid = requestCheckvar(request("itemid"),9)
	page   = requestCheckvar(request("page"),9)

	if page="" then page=1

set oitem = new CItem
	oitem.FRectItemID = itemid

	if itemid<>"" then
		oitem.GetOneItem
	end if

set oitemoption = new CItemOption
	oitemoption.FRectItemID = itemid
	if itemid<>"" then
		oitemoption.GetItemOptionInfo
	end if

set oPsItemList = new CPlusSaleItem
	oPsItemList.FCurrPage = page
	oPsItemList.FPageSize = 20
	oPsItemList.FRectItemID = itemid
    oPsItemList.GetPlusSaleSubItemListByMainLinkItemID
%>

<script language='javascript'>

function PlusDIYItem_Main_Reg(iitemid){

	var acURL = "<%= server.UrlEncode("/academy/itemmaster/PlusDIYItem/PlusDIYItem_Process.asp?mode=PlusMainAddArr&PlusSaleItemid=") %>" + iitemid;
	var popwin = window.open("/academy/itemmaster/plusdiyitem/pop_plusdiyitem_list.asp?sellyn=Y&usingyn=Y&saleYn=&defaultmargin=0&plusSaleLinkItemID=<%=ItemID%>&acURL=" + acURL, "PlusDIYItem_Main_Reg", "width=1024,height=768,scrollbars=yes,resizable=yes");

    popwin.focus();
}

function delPlusSaleMainItem(){
    var frm;
    frm = document.frmList;
    var chkExists = false;

    if (!frm.chkitem){
        return;
    }else{
        if(!frm.chkitem.length){
            chkExists = frm.chkitem.checked;
        }else{
            for (var i=0;i<frm.chkitem.length;i++){
                if (frm.chkitem[i].checked){
                    chkExists = true;
                }
            }
        }
    }

    if (!chkExists){
        alert('���õ� ������ �����ϴ�.');
        return;
    }

    if (confirm('���õ� �߰����� ��ǰ�� ���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

//��ü ����
function jsChkAll(){
    var frm;
    frm = document.frmList;
	if (frm.chkAll.checked){
	   if(typeof(frm.chkitem) !="undefined"){
	   	   if(!frm.chkitem.length){
		   	 	frm.chkitem.checked = true;
		   }else{
				for(i=0;i<frm.chkitem.length;i++){
					frm.chkitem[i].checked = true;
			 	}
		   }
	   }
	} else {
	  if(typeof(frm.chkitem) !="undefined"){
	  	if(!frm.chkitem.length){
	   	 	frm.chkitem.checked = false;
	   	}else{
			for(i=0;i<frm.chkitem.length;i++){
				frm.chkitem[i].checked = false;
			}
		}
	  }

	}

}

</script>

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		��ǰ�ڵ� :
		<input type="text" class="text" name="itemid" value="<%= itemid %>" size="7" maxlength="7">
		<!--<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">-->
	</td>
</tr>
</form>
</table>

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oitem.FResultCount>0 then %>
<tr bgcolor="#FFFFFF">
	<td rowspan="<%= 6 + oitemoption.FResultCount -1 %>" width="100" valign="top" align="center">
		<img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100" border="0">
	</td>
  	<td width="60" bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
  	<td width="300"><%= oitem.FOneItem.FItemID %></td>
  	<td width="80" bgcolor="<%= adminColor("tabletop") %>"><!-- PlusSale���� --></td>
  	<td>
  	    �÷��� D.I.Y ��ǰ
  	</td>
</tr>
<tr bgcolor="#FFFFFF">
  	<td bgcolor="<%= adminColor("tabletop") %>">�귣��ID</td>
  	<td><%= oitem.FOneItem.FMakerid %></td>
  	<td bgcolor="<%= adminColor("tabletop") %>">�Ǹſ���</td>
  	<td><%= oitem.FOneItem.FSellyn %></td>
</tr>
<tr bgcolor="#FFFFFF">
  	<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��</td>
  	<td><%= oitem.FOneItem.FItemName %></td>
  	<td bgcolor="<%= adminColor("tabletop") %>">��뿩��</td>
  	<td><%= oitem.FOneItem.FIsUsing %></td>
</tr>
<tr bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td colspan=4>
		<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuyCash,0) %>
		&nbsp;
		<% if oitem.FOneItem.FSellcash<>0 then %>
		<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %> %
		<% end if %>
		&nbsp;&nbsp;
		<!-- ���ο���/�������뿩�� -->
		<% if (oitem.FOneItem.FsaleYn="Y") then %>
		    <font color=red>
		    <% if (oitem.FOneItem.Forgprice<>0) then %>
		        <%= CLng((oitem.FOneItem.Forgprice-oitem.FOneItem.Fsellcash)/oitem.FOneItem.Forgprice*100) %> %
		    <% end if %>
		     ����
		    </font>
		<% end if %>

		<% if (oitem.FOneItem.Fitemcouponyn="Y") then %>

		    <font color=green><%= oitem.FOneItem.GetCouponDiscountStr %> ����
		    (<%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %>)</font>
		<% end if %>

		&nbsp;&nbsp;
		<%= fnColor(oitem.FOneItem.FMwDiv,"mw") %>
	</td>
</tr>
    <% if oitemoption.FResultCount>1 then %>
    <!-- �ɼ��� �ִ°�� -->
    <% for i=0 to oitemoption.FResultCount -1 %>
	    <% if oitemoption.FITemList(i).FOptIsUsing<>"Y" then %>
	    <tr bgcolor="#FFFFFF">
	      	<td bgcolor="<%= adminColor("tabletop") %>"><font color="#AAAAAA">�ɼǸ� :</font></td>
	      	<td><font color="#AAAAAA"><%= oitemoption.FITemList(i).FOptionName %></font></td>
	      	<td bgcolor="<%= adminColor("tabletop") %>"><font color="#AAAAAA">�������� : </font></td>
	      	<td><font color="#AAAAAA"><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</font></td>
	    </tr>
	    <% else %>


	    <tr bgcolor="#FFFFFF">
	      	<td bgcolor="<%= adminColor("tabletop") %>">�ɼǸ�</td>
	      	<td><%= oitemoption.FITemList(i).FOptionName %></td>
	      	<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
	      	<td><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</td>
	    </tr>
	    <% end if %>
    <% next %>
    <% else %>
	<tr bgcolor="#FFFFFF">
      	<td bgcolor="<%= adminColor("tabletop") %>">�ɼǸ�</td>
      	<td>-</td>
      	<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
      	<td><font color="<%= ynColor(oitem.FOneItem.Flimityn) %>"><%= oitem.FOneItem.Flimityn %> (<%= oitem.FOneItem.GetLimitEa %>)</font></td>
    </tr>
    <% end if %>

<% else %>
<tr bgcolor="#FFFFFF"><td align="center"> ��ǰ �˻� ����� �����ϴ�. </td></tr>
<% end if %>
</table>
<br>

<% if (oitem.FResultCount>0) then %>
<!-- <b>���� ��ũ��ǰ�� ���.....</b> -->
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
	<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td>
				<img src="/images/icon_star.gif" border="0" align="absbottom">
				<b>�߰� ���� ��ǰ ����Ʈ</b>
				&nbsp;
				�˻���� : <b><%= oPsItemList.FTotalCount %></b>
			</td>
			<td align="right">
				<input type="button" class="button" value="���û�ǰ ����" onClick="delPlusSaleMainItem('<%= itemid %>');">
				<input type="button" class="button" value="�߰�������ǰ �߰�" onClick="PlusDIYItem_Main_Reg('<%= itemid %>');">
			</td>
		</tr>
	</table>
	</td>
</tr>
<form name="frmList" method="post" action="PlusDIYItem_Process.asp">
<input type="hidden" name="mode" value="PlusSaleDellArr">
<input type="hidden" name="PlusSaleLinkItemid" value="<%= itemid %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
	<td width="60">��ǰ�ڵ�</td>
	<td width="50">�̹���</td>
  	<td>�귣��ID</td>
  	<td>��ǰ��</td>
  	<td width="50">�ǸŰ�</td>
	<td width="50">���԰�</td>
	<td width="40">����</td>
	<td width="80">��౸��</td>
	<td width="40">�÷���<br>������</td>
	<td width="35">�Ǹ�<br>����</td>
	<!-- td width="35">PS<br>���</td -->
</tr>
<% if oPsItemList.FResultCount>0 then %>
<% for i=0 to oPsItemList.FResultCount-1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td rowspan="2"><input type="checkbox" name="chkitem" value="<%= oPsItemList.FItemList(i).FPlusSaleItemID %>"></td>
	<td rowspan="2"><%= oPsItemList.FItemList(i).FPlusSaleItemID %></td>
  	<td rowspan="2"><img src="<%= oPsItemList.FItemList(i).FImageSmall %>" width="50" height="50" border="0"></td>
  	<td rowspan="2"><%= oPsItemList.FItemList(i).FMakerID %></td>
  	<td align="left"><%= oPsItemList.FItemList(i).FItemName %></td>
  	<td align="right"><%= FormatNumber(oPsItemList.FItemList(i).FSellCash,0) %></td>
  	<td align="right"><%= FormatNumber(oPsItemList.FItemList(i).FBuyCash,0) %></td>
  	<td><%= fnPercent(oPsItemList.FItemList(i).FBuyCash,oPsItemList.FItemList(i).FSellCash,1) %></td>
  	<td><%= oPsItemList.FItemList(i).FMwdiv %></td>
  	<td rowspan="2"><%= oPsItemList.FItemList(i).FPlusSalePro %>%</td>
  	<td rowspan="2"><%= fnColor(oPsItemList.FItemList(i).FSellyn,"sellyn") %></td>
  	<!-- td rowspan="2">Y</td -->
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="left"><%= Left(oPsItemList.FItemList(i).FPlusSaleStartDate,10) %> ~ <%= Left(oPsItemList.FItemList(i).FPlusSaleEndDate,10) %> <font color="<%= oPsItemList.FItemList(i).getCurrstateColor %>">(<%= oPsItemList.FItemList(i).getCurrstateName %>)</font></td>
	<td align="right"><font color="#CC33FF"><%= FormatNumber(oPsItemList.FItemList(i).getPlusSalePrice,0) %></font></td>
	<td align="right"><font color="#CC33FF"><%= FormatNumber(oPsItemList.FItemList(i).getPlusSaleBuycash,0) %></font></td>
	<td><%= fnPercent(oPsItemList.FItemList(i).getPlusSaleBuycash,oPsItemList.FItemList(i).getPlusSalePrice,1) %></td>
	<td><%= oPsItemList.FItemList(i).getMaginFlagName %></td>
</tr>
<% next %>
<tr>
    <td colspan="20" align="center" bgcolor="#FFFFFF">
        <% if oPsItemList.HasPreScroll then %>
			<a href="javascript:NextPage('<%= oPsItemList.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oPsItemList.StartScrollPage to oPsItemList.FScrollCount + oPsItemList.StartScrollPage - 1 %>
			<% if i>oPsItemList.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oPsItemList.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
    </td>
</tr>
<% else %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
		���� �������� �÷��� D.I.Y ��ǰ�� �����ϴ�.
	</td>
</tr>
<% end if %>
</form>
</table>
<% end if %>

<%
	set oitem = Nothing
	set oitemoption = Nothing
	set oPsItemList = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
