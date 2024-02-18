<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  상품검색
' History : 2009.04.07 서동석 생성
'			2012.08.29 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_dailystock.asp"-->
<!-- #include virtual="/lib/classes/items/new_itemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/stock/offitemstock_cls.asp"-->
<%
dim i,BasicMonth ,sqlStr ,barcode, itemgubun, itemid, itemoption
	barcode = request("barcode")

BasicMonth = Left(CStr(DateSerial(Year(now()),Month(now())-1,1)),7)
const C_STOCK_DAY=7

'바코드 검색
if barcode <> "" then
	if Len(barcode) >= "12" then
	    sqlStr = "select top 1 b.itemgubun ,b.itemid ,b.itemoption"
	    sqlStr = sqlStr + " from [db_item].[dbo].tbl_item_option_stock b"
	    sqlStr = sqlStr + " where b.barcode='" & barcode & "'"

	    'response.write sqlStr & "<br>"
	    rsget.Open sqlStr,dbget,1

	    if Not rsget.Eof then
	    	itemgubun = rsget("itemgubun")
	    	itemid = rsget("itemid")
	    	itemoption = rsget("itemoption")
	    end if

	    rsget.Close

	    if itemid = "" then
			sqlStr = "select top 1 i.itemgubun, i.shopitemid , i.itemoption"
			sqlStr = sqlStr & " from db_shop.dbo.tbl_shop_item i"
			sqlStr = sqlStr + " where i.extbarcode='" & barcode & "'"

		    'response.write sqlStr & "<br>"
		    rsget.Open sqlStr,dbget,1

		    if Not rsget.Eof then
		    	itemgubun = rsget("itemgubun")
		    	itemid = rsget("shopitemid")
		    	itemoption = rsget("itemoption")
		    end if

		    rsget.Close

		    if itemid = "" then
	            IF (Len(barcode)=12) and ((Left(barcode,2)="10") or (Left(barcode,2)="90") or (Left(barcode,2)="70") or (Left(barcode,2)="80")) then
	                itemgubun = Left(barcode,2)
	                itemid = CLng(Mid(barcode,3,6))
	                itemoption = Right(barcode,4)
	            end if

	            IF (Len(barcode)=14) and ((Left(barcode,2)="10") or (Left(barcode,2)="90") or (Left(barcode,2)="70") or (Left(barcode,2)="80")) then
	                itemgubun = Left(barcode,2)
	                itemid = CLng(Mid(barcode,3,8))
	                itemoption = Right(barcode,4)
	            end if
		    end if
	    end if

	else
		response.write "<script language='javascript'>"
		response.write "	alert('바코드 길이가 짧습니다. 12자리 이상으로 입력하세요.');"
		response.write "	history.go(-1);"
		response.write "</script>"
		response.end	:	dbget.close()
	end if
end if

dim oitem
if itemgubun = "10" then
	set oitem = new CItemInfo
		oitem.FRectItemID = itemid

		if itemid<>"" then
			oitem.GetOneItemInfo
		end if

else
	set oitem = new CoffstockItemlist	'//온라인 인크루드파일 클래스명 충돌, 오프라인 따로 생성
		oitem.frectitemgubun = itemgubun
		oitem.FRectItemID = itemid
		oitem.frectitemoption = itemoption

		if itemid<>"" then
			oitem.GetoffItemDefaultData
		end if
end if

dim oitemoption
set oitemoption = new CItemOption
	oitemoption.FRectItemID = itemid

	if itemid<>"" then
		oitemoption.GetItemOptionInfo
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
%>
<script language='javascript'>

function Research(){
    document.frm.submit();
}

function GetOnLoad(){

	window.resizeTo(1024,500);

    document.frm.barcode.select();
    document.frm.barcode.focus();
}
window.onload=GetOnLoad;

function chbarcode(itemgubun,itemid){
	var itemoption;

	itemoption = frm.itemoption.value;
	frm.barcode.value=itemgubun+itemid+itemoption;
	frm.submit();
}

</script>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method=get>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="3">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong>상품검색</strong></font>
			    </td>
			    <td align="right">
					바코드: <input type="text" class="text" name="barcode" value="<%= barcode %>" size=14 maxlength=14>
    				<input type="button" class="button" value="검색" onclick="Research()">
				</td>
			</tr>
		</table>
	</td>
</tr>
<% if itemid <> "" then %>
<tr bgcolor="#FFFFFF">
	<td colspan=3 align="center">
		<% if oitem.fresultcount > 0 then %>
			<% if itemgubun="10" then %>
				<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
				<tr bgcolor="#FFFFFF">
					<td rowspan=<%= 6 + oitemoption.FResultCount -1 %> width="110" valign=top align=center><img src="<%= oitem.FOneItem.FListImage %>" width="100" height="100"></td>
				  	<td width="60">상품코드</td>
				  	<td width="300">
				  		<%= itemgubun %> <b><%= CHKIIF(oitem.FOneItem.FItemID>=1000000,Format00(8,oitem.FOneItem.FItemID),Format00(6,oitem.FOneItem.FItemID)) %></b> <%= itemoption %>
				  		&nbsp;&nbsp;
			    		<% if oitemoption.FResultCount>0 then %>
			    			<select class="select" name="itemoption">
								<option value="0000">----

								<% for i=0 to oitemoption.FResultCount-1 %>
									<option value="<%= oitemoption.FItemList(i).FItemOption %>" <% if itemoption=oitemoption.FItemList(i).FItemOption then response.write "selected" %> ><%= oitemoption.FItemList(i).Foptionname %>
								<% next %>
							</select>
							<input type="button" class="button" value="옵션검색" onclick="chbarcode('<%= itemgubun %>','<%= CHKIIF(oitem.FOneItem.FItemID>=1000000,Format00(8,oitem.FOneItem.FItemID),Format00(6,oitem.FOneItem.FItemID)) %>');">
						<% end if %>
				  	</td>
				  	<td width="60">상품랙코드</td>
				  	<td ><%= oitem.FOneItem.Fitemrackcode %></td>
				  	<td>평균배송소요일 :
					<% if (oitem.FOneItem.FavgDLvDate>-1) then %>
					    <a href="javascript:popItemAvgDlvList('<%= itemid %>');">D+<%= oitem.FOneItem.FavgDLvDate+1 %></a>
					<% else %>
					    <a href="javascript:popItemAvgDlvList('<%= itemid %>');">데이터 없음</a>
					<% end if %>
					</td>

				</tr>
				<tr bgcolor="#FFFFFF">
				  	<td>브랜드ID</td>
				  	<td><%= oitem.FOneItem.FMakerid %> (브랜드 랙코드 : <font color="red"><b><%= Format00(4,oitem.FOneItem.FRackCode) %></b></font>)</td>
				  	<td>판매여부</td>
				  	<td colspan=2><font color="<%= ynColor(oitem.FOneItem.FSellyn) %>"><%= oitem.FOneItem.FSellyn %></font></td>
				</tr>
				<tr bgcolor="#FFFFFF">
				  	<td>상품명</td>
				  	<td><%= oitem.FOneItem.FItemName %></td>
				  	<td>사용여부</td>
				  	<td colspan=2><font color="<%= ynColor(oitem.FOneItem.FIsUsing) %>"><%= oitem.FOneItem.FIsUsing %></font></td>
				</tr>
				<tr bgcolor="#FFFFFF">
				  	<td>판매가</td>
				  	<td>
				  		<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
				  		&nbsp;&nbsp;
				  		<font color="<%= oitem.FOneItem.getMwDivColor %>"><%= oitem.FOneItem.getMwDivName %></font>
				  	    <% if oitem.FOneItem.FSellcash<>0 then %>
						<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %> %
						<% end if %>
						&nbsp;&nbsp;
						<!-- 할인여부/쿠폰적용여부 -->
						<% if (oitem.FOneItem.FSailYn="Y") then %>
						    <font color=red>
						    <% if (oitem.FOneItem.Forgprice<>0) then %>
						        <%= CLng((oitem.FOneItem.Forgprice-oitem.FOneItem.Fsellcash)/oitem.FOneItem.Forgprice*100) %> %
						    <% end if %>
						     할인
						    </font>
						<% end if %>

						<% if (oitem.FOneItem.Fitemcouponyn="Y") then %>

						    <font color=green><%= oitem.FOneItem.GetCouponDiscountStr %> 쿠폰
						    (<%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %>)</font>
						<% end if %>

				  	</td>
				  	<td>단종여부</td>
				  	<td colspan="2">
				  		<%= fncolor(oitem.FOneItem.Fdanjongyn,"dj") %>
				  		<% if oitem.FOneItem.Fdanjongyn="N" then %>
						생산중
						<% end if %>
					</td>

				</tr>

				<% if oitemoption.FResultCount>1 then %>
				    <!-- 옵션이 있는경우 -->
				    <% for i=0 to oitemoption.FResultCount -1 %>
					    <% if oitemoption.FITemList(i).FOptIsUsing<>"Y" then %>
					    <tr bgcolor="#FFFFFF">
					      	<td><font color="#AAAAAA">옵션명 :</font></td>
					      	<td><font color="#AAAAAA"><%= oitemoption.FITemList(i).FOptionName %></font></td>
					      	<td><font color="#AAAAAA">한정여부 : </font></td>
					      	<td><font color="#AAAAAA"><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</font></td>
					      	<td>한정 비교재고 (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)</td>
					    </tr>
					    <% else %>

					    <% if oitemoption.FITemList(i).Fitemoption=itemoption then %>
					    <tr bgcolor="#EEEEEE">
					    <% else %>
					    <tr bgcolor="#FFFFFF">
					    <% end if %>
					      	<td>옵션명</td>
					      	<td><%= oitemoption.FITemList(i).FOptionName %></td>
					      	<td>한정여부</td>
					      	<td><font color="<%= ynColor(oitemoption.FITemList(i).Foptlimityn) %>"><%= oitemoption.FITemList(i).Foptlimityn %></font> (<%= oitemoption.FITemList(i).GetOptLimitEa %>)</td>
					      	<td>
					      	  한정 비교재고 (<b><%= oitemoption.FITemList(i).GetLimitStockNo %></b>)
						      <% if (oitem.FOneItem.Fdanjongyn = "S") then %>
						      (재입고예정일 : <%= oitemoption.FITemList(i).Frestockdate %>)
						      <% end if %>
					      	</td>
					    </tr>
					    <% end if %>
				    <% next %>
				<% else %>
					<tr bgcolor="#FFFFFF">
				      	<td>옵션명</td>
				      	<td>-</td>
				      	<td>한정여부</td>
				      	<td><font color="<%= ynColor(oitem.FOneItem.Flimityn) %>"><%= oitem.FOneItem.Flimityn %> (<%= oitem.FOneItem.GetLimitEa %>)</font></td>
				      	<td>
				      		한정 비교재고 (<b><%= oitem.FOneItem.GetLimitStockNo %></b>)
						<% if ((oitem.FOneItem.Fdanjongyn="S") and (oitemoption.FResultCount<1)) then %>
						(재입고예정일 : <%= restockdate %>)
						<% end if %>
				      	</td>
				    </tr>
				<% end if %>
				</table>
			<%
			'//온라인 제외 나머지
			else
			%>
				<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
				<tr bgcolor="#FFFFFF">
					<td rowspan=<%= 5 + oitem.FResultCount -1 %> width="110" valign="top" align="center">
						<img src="<%= oitem.foneitem.FImageList %>" width="100" height="100">
					</td>
				  	<td width="60"><b>*상품정보</b></td>
				  	<td width="300">
				  	</td>
				  	<td width="60">브랜드ID :</td>
				  	<td colspan=2><%= oitem.foneitem.FMakerid %></td>
				</tr>
				<tr bgcolor="#FFFFFF">
				  	<td>상품코드 :</td>
				  	<td><%= oitem.foneitem.fitemgubun %> <b><%= CHKIIF(oitem.foneitem.FItemID>=1000000,Format00(8,oitem.foneitem.FItemID),Format00(6,oitem.foneitem.FItemID)) %></b> <%= oitem.foneitem.fitemoption %></td>
				  	<td>사용여부 : </td>
				  	<td colspan=2><%= oitem.foneitem.FIsUsing %></td>
				</tr>
				<tr bgcolor="#FFFFFF">
				  	<td>상품명 :</td>
				  	<td colspan=4><%= oitem.foneitem.FItemName %></td>
				</tr>
			    <tr bgcolor="#FFFFFF">
			      	<td><font color="#AAAAAA">옵션명 :</font></td>
			      	<td><font color="#AAAAAA"><%= oitem.foneitem.FItemOptionName %></font></td>
			      	<td><font color="#AAAAAA">센터재고 : </font></td>
			      	<td>
			      		<%= oitem.foneitem.GetCheckStockNo %> : (NEW)
			      	</td>
			    </tr>
				</table>
			<% end if %>

			<Br>
			<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
			<tr align="center"  bgcolor="#FFFFFF">
				<td>입고</td>
				<td>ON<br>판매</td>
				<td>OFF<br>출고</td>
				<td>기타<br>출고</td>
				<td>CS<br>출고</td>
				<td>불량</td>
				<td>오차</td>
				<td bgcolor="<%= adminColor("tabletop") %>">실사<br>재고</td>
				<td>상품<br>준비</td>
				<td bgcolor="<%= adminColor("tabletop") %>">재고<br>파악<br>재고</td>
				<td>출고지시<br>이전</td>
				<td bgcolor="<%= adminColor("tabletop") %>">예상<br>재고</td>
			</tr>
			<tr align="center" bgcolor="#FFFFFF">
				<td><%= osummarystock.FOneItem.Ftotipgono %></td>
				<td><%= osummarystock.FOneItem.Ftotsellno*(-1) %></td>
				<td><%= osummarystock.FOneItem.Foffchulgono + osummarystock.FOneItem.Foffrechulgono %></td>
				<td><%= osummarystock.FOneItem.Fetcchulgono + osummarystock.FOneItem.Fetcrechulgono %></td>
				<td><%= osummarystock.FOneItem.Ferrcsno %></td>
				<td><%= osummarystock.FOneItem.Ferrbaditemno %></td>
				<td><%= osummarystock.FOneItem.Ferrrealcheckno %></td>
				<td bgcolor="<%= adminColor("tabletop") %>"><%= osummarystock.FOneItem.Frealstock %></td>
				<td><%= osummarystock.FOneItem.Fipkumdiv5 + osummarystock.FOneItem.Foffconfirmno %></td>
				<td bgcolor="<%= adminColor("tabletop") %>"><%= osummarystock.FOneItem.GetCheckStockNo %></td>
				<td><%= osummarystock.FOneItem.Fipkumdiv4 + osummarystock.FOneItem.Fipkumdiv2 + osummarystock.FOneItem.Foffjupno %></td>
				<td bgcolor="<%= adminColor("tabletop") %>"><%= osummarystock.FOneItem.GetMaystock %></td>
			</tr>
			</table>

		<% else %>
			검색결과가 없습니다.
		<% end if %>
	</td>
</tr>

<% else %>
<tr bgcolor="ffffff" align="center">
	<td>
		바코드를 입력하세요.
	</td>
</tr>
<% end if %>
</form>
</table>

<%
	set oitem = Nothing
	set oitemoption = Nothing
	set osummarystock = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->