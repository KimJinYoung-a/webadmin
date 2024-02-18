<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/realjaegocls.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<!-- #include virtual="/lib/classes/offshop_dailystock.asp"-->
<%

dim BasicMonth
BasicMonth = Left(CStr(DateSerial(Year(now()),Month(now())-1,1)),7)

const C_STOCK_DAY=7

dim sqlStr
dim barcode
dim itemgubun, itemid, itemoption

barcode = request("barcode")

'������ڵ� �˻�
if Len(barcode)>=12 then
        sqlStr = "select top 1 b.* " + VbCrlf
        sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option_stock b " + VbCrlf
        sqlStr = sqlStr + " where b.barcode='" + CStr(barcode) + "' " + VbCrlf
        'response.write sqlStr
        rsget.Open sqlStr,dbget,1
        if Not rsget.Eof then
        	itemgubun = rsget("itemgubun")
        	itemid = rsget("itemid")
        	itemoption = rsget("itemoption")
        else
        	itemgubun = Left(barcode,2)
        	itemid = CLng(Mid(barcode,3,6))
        	itemoption = Right(barcode,4)
        end if
        rsget.Close
else
        itemgubun="10"
        itemid = barcode
        itemoption="0000"
end if


if itemgubun="" then itemgubun="10"
if itemoption="" then itemoption="0000"

dim ojaegoitem
set ojaegoitem = new CRealJaeGo
ojaegoitem.FRectItemID = itemid
if itemid<>"" then
	ojaegoitem.GetItemDefaultDataStock
end if

dim oitemoption
set oitemoption = new CItemOptionInfo
oitemoption.FRectItemID =  itemid
if itemid<>"" then
	oitemoption.getOptionList
end if

if (oitemoption.FResultCount<1) then
	itemoption = "0000"
end if

dim osummarystock
set osummarystock = new CSummaryItemStock
osummarystock.FRectStartDate = BasicMonth + "-01"
osummarystock.FRectItemGubun = itemgubun
osummarystock.FRectItemID =  itemid
osummarystock.FRectItemOption =  itemoption
if itemid<>"" then
	osummarystock.GetCurrentItemStock
	osummarystock.GetDaily_Logisstock_Summary
end if

dim i, j, k, found
dim starti, endi


dim x, y, max_x, max_y, size_x, size_y
dim bgcolors(9)

bgcolors(0) = "ff8040"
bgcolors(1) = "fcf00e"
bgcolors(2) = "baf30c"
bgcolors(3) = "green"
bgcolors(4) = "teal"
bgcolors(5) = "0612f2"
bgcolors(6) = "blue"
bgcolors(7) = "purple"
bgcolors(8) = "gray"

max_x = 23
max_y = 38
size_x = 16
size_y = 4




%>
<script>
function FrameControl(imagesrc){
	 document.getElementById("imgtarget").src = imagesrc;
}

function BTNChange(id,Max){
 	/*
 	var idnum = id.substring(3,4);
 	try{
	  	for(i=0;i<Max;i++){
	      if (idnum == i){
			  eval('document.getElementById("btn' + i + '").src  ="/images/shopping/add_0' + (i + 1) + '.gif"');
	      }
		  else{
			  eval('document.getElementById("btn' + i + '").src  ="/images/shopping/add_b0' + (i + 1) + '.gif"');
		  }
	  	}
	 }catch(e){

	 }
	 */
}

function Research(){
	document.frm.submit();
}



function itemrackcodereg(itemrackcode){
	var popwin = window.open('/pop/popitemrackcode_input.asp?itemrackcode=' + itemrackcode,'popitemrackcode_input','width=500,height=400,resizabled=yes,scrollbars=yes');
	popwin.focus();
}

</script>

<!-- TOP -->
<table width="280" border="0" align="center" cellpadding="1" cellspacing="1" class="a">
  <tr height="20">
  	<td>
    	<img src="/images/icon_star.gif" align="absbottom"><font color="red"><strong>��ǰ�˻�</strong></font>
	</td>
	<td align="right">
    	<a href="/PDAadmin/index.asp">HOME</a>
	</td>
  </tr>
</table>
<!-- TOP -->


<!-- ǥ ��ܰ˻� ����-->
<table width="280" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
	<form name="frm" method=get>
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td valign="top" bgcolor="F4F4F4">
	        	<input type="text" class="text" name="barcode" value="<%= barcode %>" size=14 maxlength=14>
	        	<input type="button" class="button" value="�˻�" onclick="Research()">
	        </td>
	</tr>
<!--
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
		<td>
			<% if (itemid<>"" and ojaegoitem.FResultCount > 0) then %>
	    		<input type="button" value="������ڵ� ���/����" onclick="publicbarreg('<%= barcode %>');">
	    	<% end if %>
		</td>
	<tr>

	<tr height="25" valign="bottom" bgcolor="F4F4F4">
		<td>
			<% if (ojaegoitem.Fresultcount>0) then %>
	    	<input type="button" value="��ǰ ���ڵ� ���/����" onclick="itemrackcodereg('<%= ojaegoitem.FItemList(0).Fitemrackcode %>');">
	    	<% else %>
	    	<input type="button" value="��ǰ ���ڵ� ���/����" onclick="itemrackcodereg('');">
	    	<% end if %>
	    </td>
	</tr>
-->


	</form>
</table>
<!-- ǥ ��ܰ˻� ��-->


<% if (itemid<>"" and ojaegoitem.FResultCount > 0) then %>
<table width="280" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
			<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>��ǰ������</strong></font>
			&nbsp;&nbsp;
			<%= ojaegoitem.FItemList(0).FItemName %>
			&nbsp;&nbsp;
			<% if oitemoption.FResultCount>0 then %>
			�ɼǼ��� :
			<select name="itemoption">
			<option value="0000">----
			<% for i=0 to oitemoption.FResultCount-1 %>
			<option value="<%= oitemoption.FItemList(i).FItemOption %>" <% if itemoption=oitemoption.FItemList(i).FItemOption then response.write "selected" %> ><%= oitemoption.FItemList(i).FItemOptionName %>
			<% next %>
			</select>
			<% end if %>
			&nbsp;
			<input type=button value="�˻�" onclick="document.frm.submit();">
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td bgcolor="#FFFFFF">
    		<table width="100%" border="0" cellspacing=2 cellpadding=0 class=a>
    			<tr>
    				<td width="400" align="left" valign="top">
						<table width="100%" border=0 cellspacing=2 cellpadding=0 class=a>
							<tr>
								<td colspan="10" width="410" valign="top"><img src="<%= ojaegoitem.FItemList(0).FImageSmall %>"  id="imgtarget" onError="this.src='/images/sampleimage_400.jpg'"></td>
							</tr>

						</table>
					</td>
					<td valign="top">
						<table width="100%" border=1 cellspacing=2 cellpadding=0 class=a>
			    			<tr>
			    				<td height="30" colspan="2"><font size="3" style="line-height:100%"><b>��ǰ�� : <%= ojaegoitem.FItemList(0).FItemName %></b><%= ojaegoitem.FItemList(0).FItemOptionName %></font></td>
			    			</tr>
			    			<tr>
								<td height="1" colspan="3" bgcolor="#CCCCCC"></td>
							</tr>
			    			<tr>
			    				<td height="30" colspan="2"><font size="3" style="line-height:100%">
			    					<b>����ľ���� : <%= osummarystock.FOneItem.Frealstock + osummarystock.FOneItem.Fipkumdiv5 + osummarystock.FOneItem.Foffconfirmno %>
			    				 	��ǰ ���ڵ� : <font color=red><%= ojaegoitem.FItemList(0).Fitemrackcode %></font>)</b></font>
			    				</td>
			    			</tr>
			    			<tr>
								<td height="1" colspan="3" bgcolor="#CCCCCC"></td>
							</tr>
			    			<tr>
			    				<td width="80">�ɼǸ�:</td>
			    				<td>
        <% for i=0 to ojaegoitem.FResultCount -1 %>
                <% if ojaegoitem.FItemList(i).Foptionusing<>"N" then %>
                                                          <%= ojaegoitem.FItemList(i).FItemOptionName %><br>
                <% end if %>
        <% next %>
			    				</td>
			    			</tr>
			    			<tr>
								<td height="1" colspan="3" bgcolor="#CCCCCC"></td>
							</tr>
			    			<tr>
			    				<td width="80">��ǰ�ڵ�:</td>
			    				<td>10 <b><%= Format00(6,ojaegoitem.FItemList(0).FItemID) %></b> <%= itemoption %></td>
			    			</tr>
			    			<tr>
								<td height="1" colspan="3" bgcolor="#CCCCCC"></td>
							</tr>
			    			<tr>
			    				<td>�귣��ID:</td>
			    				<td><%= ojaegoitem.FItemList(0).FMakerid %>(�귣�� ���ڵ� : <font color=red><b><%= Format00(4,ojaegoitem.FItemList(0).FRackCode) %></b></font>)</td>
			    			</tr>
			    			<tr>
								<td height="1" colspan="3" bgcolor="#CCCCCC"></td>
							</tr>
			    			<tr>
			    				<td>�Ǹſɼ�:</td>
			    				<td>�Ǹ�(<%= ojaegoitem.FItemList(0).FSellyn %>)&nbsp;&nbsp;���(<%= ojaegoitem.FItemList(0).FIsUsing %>)&nbsp;&nbsp;����(<%= ojaegoitem.FItemList(0).FLimitYn %>/<%= ojaegoitem.FItemList(0).GetLimitStr %>)
			    			</tr>
			    			<tr>
								<td height="1" colspan="3" bgcolor="#CCCCCC"></td>
							</tr>
			    			<tr>
			    				<td>��ۿɼ�:</td>
			    				<td><%= ojaegoitem.FItemList(0).GetDeliveryName %></td>
			    			</tr>

			    			<tr>
								<td height="1" colspan="3" bgcolor="#CCCCCC"></td>
							</tr>

							<tr>
								<td colspan="3" ></td>
							</tr>
							<!--
			    			<tr>
			    				<td><b>*���⳻��</b></td>
			    				<td align="right"><input type="button" value="����" onclick="Research()"></td>
			    			</tr>
			    			<tr>
								<td height="1" colspan="3" bgcolor="#CCCCCC"></td>
							</tr>
			    			<tr>
			    				<td>����ľ����</td>
			    				<td><%= osummarystock.FOneItem.Frealstock + osummarystock.FOneItem.Fipkumdiv5 + osummarystock.FOneItem.Foffconfirmno %>&nbsp;&nbsp;[<%= osummarystock.FOneItem.Flastupdate %>]</td>
			    			</tr>
			    			<tr>
								<td height="1" colspan="3" bgcolor="#CCCCCC"></td>
							</tr>
			    			<tr>
			    				<td>�ѽǻ����</td>
			    				<td><%= osummarystock.FOneItem.Ferrrealcheckno %></td>
			    			</tr>
			    			<tr>
								<td height="1" colspan="3" bgcolor="#CCCCCC"></td>
							</tr>
			    			<tr>
			    				<td>�ǻ����</td>
			    				<td><%= osummarystock.FOneItem.Frealstock %></td>
			    			</tr>
			    			<tr>
								<td height="1" colspan="3" bgcolor="#CCCCCC"></td>
							</tr>
			    			<tr>
			    				<td>[ON]<%= osummarystock.FOneItem.Fmaxsellday %>���Ǹż���</td>
			    				<td><%= osummarystock.FOneItem.Fsell7days*-1 %></td>
			    			</tr>
			    			<tr>
			    				<td>[OFF]������</td>
			    				<td><%= osummarystock.FOneItem.Foffchulgo7days*-1 %></td>
			    			</tr>
			    			<tr>
								<td height="1" colspan="3" bgcolor="#CCCCCC"></td>
							</tr>
			    			<tr>
			    				<td>����ּ���</td>
			    				<td><%= osummarystock.FOneItem.Fpreorderno %></td>
			    			</tr>
			    			<tr>
								<td height="1" colspan="3" bgcolor="#CCCCCC"></td>
							</tr>
			    			<tr>
			    				<td>�����������</td>
			    				<td><%= osummarystock.FOneItem.Fshortageno %></td>
			    			</tr>
			    			<tr>
								<td height="1" colspan="3" bgcolor="#CCCCCC"></td>
							</tr>
			    			<tr>
			    				<td>���</td>
			    				<td></td>
			    			</tr>
			    			-->
			    		</table>
					</td>
					<td>
						<table width="200" border=0 cellspacing=2 cellpadding=0 class=a>
							<tr align="center">
								<td>
									�ʳ��� ��
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td height="10"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td height="10" background="/images/tbl_blue_round_08.gif"></td>
		<td height="10"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>

<% end if %>
<%
set oitemoption = Nothing
set ojaegoitem = Nothing
set osummarystock = Nothing
%>

<script language='javascript'>
function getOnLoad(){
    document.frm.barcode.select();
    document.frm.barcode.focus();
}

window.onload=getOnLoad;
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->