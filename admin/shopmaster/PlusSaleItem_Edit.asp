<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<!-- #include virtual="/lib/classes/items/PlusSaleItemCls.asp"-->

<%

dim itemid, page, pagem
itemid = requestCheckvar(request("itemid"),9)
page   = requestCheckvar(request("page"),9)
pagem   = requestCheckvar(request("pagem"),9)

if page="" then page=1
if pagem="" then pagem=1

dim oitem
set oitem = new CItem
oitem.FRectItemID = itemid

if itemid<>"" then
	oitem.GetOneItem
end if

dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
end if

dim oPlusSaleItem
set oPlusSaleItem = new CPlusSaleItem
oPlusSaleItem.FRectItemID = itemid

if itemid<>"" then
	oPlusSaleItem.GetOnePlusSaleSubItem
end if

dim i
dim IsPlusSaleItem        '' �÷������� ��ǰ����
IsPlusSaleItem = (oPlusSaleItem.FResultCount>0)


'' ���� IsLinkedItem �ΰ��
dim IsLinkedItem
if itemid<>"" then
    IsLinkedItem = oPlusSaleItem.IsPlusSaleLinkItem
end if

'// ���λ�ǰ ��� ����
dim oPsItemList
set oPsItemList = new CPlusSaleItem
oPsItemList.FCurrPage = page
oPsItemList.FPageSize = 20
oPsItemList.FRectItemID = itemid
if (IsPlusSaleItem) then
    oPsItemList.FRectPlusSaleItemID = itemid
    oPsItemList.GetPlusSaleMainItemList
end if

'// �߰���ǰ ��� ����
dim oPmItemList
set oPmItemList = new CPlusSaleItem
oPmItemList.FCurrPage = pagem
oPmItemList.FPageSize = 20
oPmItemList.FRectItemID = itemid
if (IsLinkedItem) then
    oPmItemList.GetPlusSaleSubItemListByMainLinkItemID
end if
%>

<script language='javascript'>

function PlusSaleItem_Main_Reg(iitemid){
	//var popwin = window.open('PlusSaleItem_Main_Reg.asp','PlusSaleItem_Main_Reg','width=900,height=600,scrollbars=yes,resizable=yes')
	//popwin.focus();
	var acURL = "<%= server.UrlEncode("/admin/shopmaster/PlusSaleItem_Process.asp?mode=PlusMainAddArr&PlusSaleItemid=") %>" + iitemid;
	var popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&sailyn=&defaultmargin=0&acURL=" + acURL, "PlusSaleItem_Main_Reg", "width=800,height=500,scrollbars=yes,resizable=yes");

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

    if (confirm('���õ� ���θ�ũ ��ǰ�� ���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}

function delPlusSaleItem(){
    var frm;
    frm = document.frmListm;
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

    if (confirm('���õ� �÷��� ���� ��ǰ�� ���� �Ͻðڽ��ϱ�?')){
        frm.submit();
    }
}

function PlusSaleItem_Sub_Reg(iitemid){
	var popwin = window.open('PlusSaleItem_Sub_Reg.asp?PlusSaleLinkItemid=' + iitemid,'PlusSaleItem_Sub_Reg','width=900,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function PlusSaleItem_Sub_Direct_Reg(iitemid){
	var acURL = "<%= server.UrlEncode("/admin/shopmaster/PlusSaleItem_Process.asp?mode=PlusSubDirectAddArr&PlusSaleLinkItemid=") %>" + iitemid;
	var popwin = window.open("/admin/itemmaster/pop_itemAddInfo.asp?sellyn=Y&usingyn=Y&sailyn=&defaultmargin=0&acURL=" + acURL, "PlusSaleItem_Sub_Reg", "width=800,height=500,scrollbars=yes,resizable=yes");

    popwin.focus();
}

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function NextPagem(page){
    document.frm.pagem.value = page;
    document.frm.submit();
}

//��ü ����
function jsChkAll(frm){
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
	<input type="hidden" name="pagem" value="">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			��ǰ�ڵ� :
			<input type="text" class="text" name="itemid" value="<%= itemid %>" size="9" maxlength="9">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<p>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if oitem.FResultCount>0 then %>
    <tr bgcolor="#FFFFFF">
        <% if  (IsPlusSaleItem) then %>
        <td rowspan="<%= 7 + oitemoption.FResultCount -1 %>" width="100" valign="top" align="center">
        <% else %>
    	<td rowspan="<%= 6 + oitemoption.FResultCount -1 %>" width="100" valign="top" align="center">
    	<% end if %>
    		<img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100" border="0">
		</td>
      	<td width="60" bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
      	<td width="300"><%= oitem.FOneItem.FItemID %></td>
      	<td width="80" bgcolor="<%= adminColor("tabletop") %>"><!-- PlusSale���� --></td>
      	<td colspan=2>
      	    �÷�������
			<% if (IsPlusSaleItem) then %>
      	    * �߰�������ǰ
      		<% end if %>

      	    <% if (IsLinkedItem) then %>
      	    * ���� ��ũ ��ǰ
      	    <% end if %>
      	</td>
    </tr>
    <% if  (IsPlusSaleItem) then %>
    <tr bgcolor="#FFFFFF">
        <td bgcolor="<%= adminColor("tabletop") %>"><font color="red">�÷������� ����</font></td>
        <td colspan=4>
            <%= oPlusSaleItem.FOneItem.FPlusSalePro %>% ����
            <%= FormatNumber(oPlusSaleItem.FOneItem.getPlusSalePrice,0) %>
            /
            <%= FormatNumber(oPlusSaleItem.FOneItem.getPlusSaleBuycash,0) %>
            (<%= oPlusSaleItem.FOneItem.FPlusSaleMargin %>%)

            <%= oPlusSaleItem.FOneItem.getMaginFlagName %>
        </td>
    </tr>
    <% end if %>
    <tr bgcolor="#FFFFFF">
      	<td bgcolor="<%= adminColor("tabletop") %>">�귣��ID</td>
      	<td><%= oitem.FOneItem.FMakerid %></td>
      	<td bgcolor="<%= adminColor("tabletop") %>">�Ǹſ���</td>
      	<td colspan=2><%= oitem.FOneItem.FSellyn %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      	<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��</td>
      	<td><%= oitem.FOneItem.FItemName %></td>
      	<td bgcolor="<%= adminColor("tabletop") %>">��뿩��</td>
      	<td colspan=2><%= oitem.FOneItem.FIsUsing %></td>
    </tr>
    <tr bgcolor="#FFFFFF">
	    <td bgcolor="<%= adminColor("tabletop") %>">��������</td>
		<td>
			<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuyCash,0) %>
			&nbsp;
			<% if oitem.FOneItem.FSellcash<>0 then %>
			<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %> %
			<% end if %>
			&nbsp;&nbsp;
			<!-- ���ο���/�������뿩�� -->
			<% if (oitem.FOneItem.FSailYn="Y") then %>
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
      	<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
      	<td colspan="2">
      		<%= fncolor(oitem.FOneItem.Fdanjongyn,"dj") %>
      		<% if oitem.FOneItem.Fdanjongyn="N" then %>
			������
			<% end if %>
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
		      	<td>���� ����� (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
		    </tr>
		    <% else %>


		    <tr bgcolor="#FFFFFF">
		      	<td bgcolor="<%= adminColor("tabletop") %>">�ɼǸ�</td>
		      	<td><%= oitemoption.FITemList(i).FOptionName %></td>
		      	<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
		      	<td><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</td>
		      	<td>���� ����� (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
		    </tr>
		    <% end if %>
	    <% next %>
        <% else %>
    	<tr bgcolor="#FFFFFF">
	      	<td bgcolor="<%= adminColor("tabletop") %>">�ɼǸ�</td>
	      	<td>-</td>
	      	<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
	      	<td><font color="<%= ynColor(oitem.FOneItem.Flimityn) %>"><%= oitem.FOneItem.Flimityn %> (<%= oitem.FOneItem.GetLimitEa %>)</font></td>
	      	<td>���� ����� (<b><%= oitem.FOneItem.GetLimitStockNo %></b>)</td>
	    </tr>
        <% end if %>

    <% else %>
    <tr bgcolor="#FFFFFF"><td align="center"> ��ǰ �˻� ����� �����ϴ�. </td></tr>

    <% end if %>
</table>
<p>



<!-- �÷������ϻ�ǰ(�߰�������ǰ)�� ��ϵǾ��� ���, ����Ʈ�� ǥ���Ѵ� -->
<% if  (IsPlusSaleItem) then %>
<br>
<!-- <b>�÷��� ���� ��ǰ�� ���.....</b> -->
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_star.gif" border="0" align="absbottom">
					<b>���� ��ũ ��ǰ ����Ʈ</b>
					&nbsp;
					�˻���� : <b><%= oPsItemList.FTotalCount %></b>
				</td>
				<td align="right">
					<input type="button" class="button" value="���û�ǰ ����" onClick="delPlusSaleMainItem('<%= itemid %>');">
					<input type="button" class="button" value="���θ�ũ��ǰ �߰�" onClick="PlusSaleItem_Main_Reg('<%= itemid %>');">
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<form name="frmList" method="post" action="PlusSaleItem_Process.asp">
	<input type="hidden" name="mode" value="PlusMainDellArr">
	<input type="hidden" name="PlusSaleItemID" value="<%= itemid %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="20"><input type="checkbox" name="chkAll" onClick="jsChkAll(document.frmList);"></td>
	   	<td width="60">��ǰ�ڵ�</td>
    	<td width="50">�̹���</td>
      	<td width="100">�귣��ID</td>
      	<td>��ǰ��</td>
      	<td width="60">�ǸŰ�</td>
		<td width="60">���԰�</td>
		<td width="40">����</td>
		<td width="40">�ŷ�<br>����</td>
		<td width="40">�Ǹ�<br>����</td>
    </tr>
    <% if oPsItemList.FResultCount<1 then %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="20" align="center">
			���� �������� ���� ��ǰ�� �����ϴ�.
		</td>
	</tr>
	<% else %>
    <% for i=0 to oPsItemList.FResultCount-1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><input type="checkbox" name="chkitem" value="<%= oPsItemList.FItemList(i).FPlusSaleLinkItemID %>"></td>
    	<td><%= oPsItemList.FItemList(i).FPlusSaleLinkItemID %></td>
    	<td><img src="<%= oPsItemList.FItemList(i).FSmallImage %>" width="50" height="50" border="0"></td>
      	<td><%= oPsItemList.FItemList(i).FMakerID %></td>
      	<td><%= oPsItemList.FItemList(i).FItemName %></td>
		<td align="right">
		    <% if oPsItemList.FItemList(i).Fsailyn="Y" then %>
      		<%= FormatNumber(oPsItemList.FItemList(i).FOrgPrice,0) %>
      		<br><font color=#F08050>(��)<%= FormatNumber(oPsItemList.FItemList(i).FSellcash,0) %></font>
		    <% else %>
		    <%= FormatNumber(oPsItemList.FItemList(i).FSellcash,0) %>
      		<% end if %>
      	</td>
      	<td align="right">
      	    <% if oPsItemList.FItemList(i).Fsailyn="Y" then %>
      		<%= FormatNumber(oPsItemList.FItemList(i).Forgsuplycash,0) %>
      		<br><font color=#F08050>(��)<%= FormatNumber(oPsItemList.FItemList(i).FBuycash,0) %></font>
      	    <% else %>
      	    <%= FormatNumber(oPsItemList.FItemList(i).FBuycash,0) %>
      		<% end if %>
      	</td>
      	<td>
      	    <% if oPsItemList.FItemList(i).Fsailyn="Y" then %>
      		<%= fnPercent(oPsItemList.FItemList(i).Forgsuplycash,oPsItemList.FItemList(i).FOrgPrice,1) %>
      		<br><font color=#F08050><%= fnPercent(oPsItemList.FItemList(i).Forgsuplycash,oPsItemList.FItemList(i).FOrgPrice,1) %></font>
      	    <% else %>
      	    <%= fnPercent(oPsItemList.FItemList(i).FBuycash,oPsItemList.FItemList(i).FSellcash,1) %>
      		<% end if %>
      	</td>
      	<td><%= fnColor(oPsItemList.FItemList(i).FMwDiv,"mw") %></td>
      	<td><%= fnColor(oPsItemList.FItemList(i).FSellyn,"sellyn") %></td>
    </tr>
    <% next %>
    </form>
    <tr>
        <td colspan="20" align="center" bgcolor="#FFFFFF">
            <% if oPsItemList.HasPreScroll then %>
    			<a href="javascript:NextPage('<%= oPsItemList.StarScrollPage-1 %>')">[pre]</a>
    		<% else %>
    			[pre]
    		<% end if %>

    		<% for i=0 + oPsItemList.StarScrollPage to oPsItemList.FScrollCount + oPsItemList.StarScrollPage - 1 %>
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
    <% end if %>
</table>
<% end if %>

<% if (oitem.FResultCount>0) then %>
    <% if IsLinkedItem then %>
    <br>
    <!-- <b>���� ��ũ��ǰ�� ���.....</b> -->
    <!-- ����Ʈ ���� -->
    <table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    	<tr height="25" bgcolor="FFFFFF">
    		<td colspan="20">
    		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
    			<tr>
    				<td>
    					<img src="/images/icon_star.gif" border="0" align="absbottom">
    					<b>�÷��� ���� ��ǰ ����Ʈ</b>
    					&nbsp;
    					�˻���� : <b><%= oPmItemList.FTotalCount %></b>
    				</td>
    				<td align="right">
    					<input type="button" class="button" value="���û�ǰ ����" onClick="delPlusSaleItem('<%= itemid %>');">
    					<input type="button" class="button" value="�߰�������ǰ �߰�" onClick="PlusSaleItem_Sub_Reg('<%= itemid %>');">
						<input type="button" class="button_auth" value="���� �߰�������ǰ �߰�" onClick="PlusSaleItem_Sub_Direct_Reg('<%= itemid %>');">
    				</td>
    			</tr>
    		</table>
    		</td>
    	</tr>
    	<form name="frmListm" method="post" action="PlusSaleItem_Process.asp">
    	<input type="hidden" name="mode" value="PlusSaleDellArr">
    	<input type="hidden" name="PlusSaleLinkItemid" value="<%= itemid %>">
    	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        	<td width="20"><input type="checkbox" name="chkAll" onClick="jsChkAll(document.frmListm);"></td>
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
        <% if oPmItemList.FResultCount>0 then %>
        <% for i=0 to oPmItemList.FResultCount-1 %>
        <tr align="center" bgcolor="#FFFFFF">
        	<td rowspan="2"><input type="checkbox" name="chkitem" value="<%= oPmItemList.FItemList(i).FPlusSaleItemID %>"></td>
        	<td rowspan="2"><%= oPmItemList.FItemList(i).FPlusSaleItemID %></td>
          	<td rowspan="2"><img src="<%= oPmItemList.FItemList(i).FImageSmall %>" width="50" height="50" border="0"></td>
          	<td rowspan="2"><%= oPmItemList.FItemList(i).FMakerID %></td>
          	<td align="left"><%= oPmItemList.FItemList(i).FItemName %></td>
          	<td align="right"><%= FormatNumber(oPmItemList.FItemList(i).FSellCash,0) %></td>
          	<td align="right"><%= FormatNumber(oPmItemList.FItemList(i).FBuyCash,0) %></td>
          	<td><%= fnPercent(oPmItemList.FItemList(i).FBuyCash,oPmItemList.FItemList(i).FSellCash,1) %></td>
          	<td><%= oPmItemList.FItemList(i).FMwdiv %></td>
          	<td rowspan="2"><%= oPmItemList.FItemList(i).FPlusSalePro %>%</td>
          	<td rowspan="2"><%= fnColor(oPmItemList.FItemList(i).FSellyn,"sellyn") %></td>
          	<!-- td rowspan="2">Y</td -->
        </tr>
        <tr align="center" bgcolor="#FFFFFF">
        	<td align="left"><%= Left(oPmItemList.FItemList(i).FPlusSaleStartDate,10) %> ~ <%= Left(oPmItemList.FItemList(i).FPlusSaleEndDate,10) %> <font color="<%= oPmItemList.FItemList(i).getCurrstateColor %>">(<%= oPmItemList.FItemList(i).getCurrstateName %>)</font></td>
        	<td align="right"><font color="#CC33FF"><%= FormatNumber(oPmItemList.FItemList(i).getPlusSalePrice,0) %></font></td>
        	<td align="right"><font color="#CC33FF"><%= FormatNumber(oPmItemList.FItemList(i).getPlusSaleBuycash,0) %></font></td>
        	<td><%= fnPercent(oPmItemList.FItemList(i).getPlusSaleBuycash,oPmItemList.FItemList(i).getPlusSalePrice,1) %></td>
        	<td><%= oPmItemList.FItemList(i).getMaginFlagName %></td>
        </tr>
        <% next %>
        <tr>
            <td colspan="20" align="center" bgcolor="#FFFFFF">
                <% if oPmItemList.HasPreScroll then %>
        			<a href="javascript:NextPagem('<%= oPmItemList.StarScrollPage-1 %>')">[pre]</a>
        		<% else %>
        			[pre]
        		<% end if %>

        		<% for i=0 + oPmItemList.StarScrollPage to oPmItemList.FScrollCount + oPmItemList.StarScrollPage - 1 %>
        			<% if i>oPmItemList.FTotalpage then Exit for %>
        			<% if CStr(pagem)=CStr(i) then %>
        			<font color="red">[<%= i %>]</font>
        			<% else %>
        			<a href="javascript:NextPagem('<%= i %>')">[<%= i %>]</a>
        			<% end if %>
        		<% next %>

        		<% if oPmItemList.HasNextScroll then %>
        			<a href="javascript:NextPagem('<%= i %>')">[next]</a>
        		<% else %>
        			[next]
        		<% end if %>
            </td>
        </tr>
        <% else %>
        <tr height="25" bgcolor="FFFFFF">
    		<td colspan="20" align="center">
    			���� �������� �÷������λ�ǰ�� �����ϴ�.
    		</td>
    	</tr>
    	<% end if %>
    	</form>
    </table>
    <% end if %>
<% end if %>
<% if (Not IsPlusSaleItem) and (Not IsLinkedItem) then %>
<!--
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="25" bgcolor="FFFFFF">
        <td align="center"> �÷��� ���� ��ǰ�� �ƴմϴ�. </td>
    </tr>
</table>
-->
<% end if %>
<%
set oitem = Nothing
set oitemoption = Nothing
set oPlusSaleItem = Nothing
set oPsItemList = Nothing
set oPmItemList = Nothing
%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
