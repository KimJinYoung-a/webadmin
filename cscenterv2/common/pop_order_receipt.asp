<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<!-- #include virtual="/cscenter/lib/popheader_xhtml.asp"-->

<!-- #include virtual="/cscenterv2/lib/classes/order/ordercls.asp"--> 
<!-- #include virtual="/lib/util/md5.asp"--> 
<%
dim i, j
dim getPdfDownLinkUrlAdm
dim orderserial
orderserial = requestCheckvar(request("orderserial"),16)

dim oordermaster, oorderdetail
dim addparam

set oordermaster = new COrderMaster
oordermaster.FRectOrderSerial = orderserial
oordermaster.QuickSearchOrderMaster

if (oordermaster.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    oordermaster.FRectOldOrder = "on"
    oordermaster.QuickSearchOrderMaster
end if

set oorderdetail = new COrderMaster
oorderdetail.FRectOldOrder = oordermaster.FRectOldOrder
oorderdetail.FRectOrderSerial = orderserial
oorderdetail.QuickSearchOrderDetail
addparam = "vos="&orderserial&"&ekey="&md5(orderserial)
if (application("Svr_Info")	= "Dev") then
  getPdfDownLinkUrlAdm = "/pdf/dnOrderReceiptPdf.asp?"&addparam
else
  getPdfDownLinkUrlAdm = "http://apps.10x10.co.kr/pdf/dnOrderReceiptPdf.asp?"&addparam
end if
%>	
<script type="text/javascript">
	function jsGoPDF(iUri){
		  var popwin = window.open(iUri,'dnPdf','width=1024,height=768,scrollbars=yes,resizable=yes'); 
	}
</script>
	<div class="heightgird" id="orderPrint"><!-- 2013.09.24 : id="mediaPrint" �߰� -->
		<div class="popWrap">
			<div class="popHeader">
				<h1><img src="http://fiximage.10x10.co.kr/web2013/my10x10/tit_order_receipt.gif" alt="�ֹ�Ȯ�μ�" /></h1>
			</div>
			<div class="popContent">
				<!-- content -->
				<div class="mySection">
					<div class="orderDetail">
						<div class="title">
							<h2 class="ftLt" style="margin-top:0;">�ֹ�����</h2>
							<!--p class="ftRt"><img src="http://fiximage.10x10.co.kr/web2013/@temp/img_barcode.gif" alt="���ڵ��̹���" /></p-->
						</div>
						<table class="baseTable rowTable">
						<caption>�ֹ����� ����</caption>
						<colgroup>
							<col width="15%" /> <col width="35%" /> <col width="15%" /> <col width="35%" />
						</colgroup>
						<tbody>
						<tr>
							<th scope="row">�ֹ���ȣ</th>
							<td><%= oordermaster.FOneItem.FOrderSerial %>
							 <% If oordermaster.FOneItem.IsForeignDeliver Then %>
									  (<strong>�ؿܹ��</strong>)
								  <% End If %>
							</td>
							<th scope="row">�ֹ�����</th>
							<td><%= left(oordermaster.FOneItem.FRegDate,10) %></td>
						</tr>
						<tr>
							<th scope="row">�������</th>
							<td><%= oordermaster.FOneItem.JumunMethodName %></td>
							<th scope="row">��������</th>
							<td><% if IsNULL(oordermaster.FOneItem.FIpkumDate) then %>
                                <strong class="crRed">�Ա���</strong>
                                <% else %>
                                <%= left(oordermaster.FOneItem.FIpkumDate,10) %>
                                <% end if %>
							</td>
						</tr>
						<tr>
							<% if oordermaster.FOneItem.FAccountdiv = 7 then %>
							<th scope="row"><%= CHKIIF(IsNULL(oordermaster.FOneItem.FIpkumDate),"�����ϽǱݾ�","�����ݾ�") %></th>
							<td><em class="crRed"><strong><%= FormatNumber(oordermaster.FOneItem.TotalMajorPaymentPrice,0) %></strong>��</em></td>
							<th scope="row">�Ա��Ͻ� ����</th>
							<td><%= Replace(oordermaster.FOneItem.Faccountno,"X","") %></td>
						 <% else %>
							<th scope="row"><%= CHKIIF(IsNULL(oordermaster.FOneItem.FIpkumDate),"�����ϽǱݾ�","�����ݾ�") %></th>
							<td colspan="3"><em class="crRed"><strong><%= FormatNumber(oordermaster.FOneItem.TotalMajorPaymentPrice,0) %></strong>��</em>
							 <% if (oordermaster.FOneItem.FAccountDiv="100") or (oordermaster.FOneItem.FAccountDiv="110") then %>
		                        <% if (oordermaster.FOneItem.FokcashbagSpend<>0) then %>
			                        : <span class="red_11px">�ſ�ī�� <%= FormatNumber(oordermaster.FOneItem.TotalMajorPaymentPrice-oordermaster.FOneItem.FokcashbagSpend,0) %> ��
			                        , OKĳ���� ��� : <%= FormatNumber(oordermaster.FOneItem.FokcashbagSpend,0) %> ��
			                	   <% end if %>
		                         </span>
		                    <% end if %>
                    		</td>
                    	<% end if %>
						</tr> 
						 <% if oordermaster.FOneItem.FspendTenCash<>0 then %>
		                <tr>
		                  <th scope="row">��ġ�ݻ��</th>
		                  <td colspan="3"><em class="crRed"><strong><%= FormatNumber(oordermaster.FOneItem.FspendTenCash,0) %></strong> ��</em></td>
		                </tr>
		                 <% end if %>
		                 <% if oordermaster.FOneItem.Fspendgiftmoney<>0 then %>
		                <tr>
		                  <th scope="row">Giftī����</th>
		                     <td colspan="3"><em class="crRed"><strong><%= FormatNumber(oordermaster.FOneItem.Fspendgiftmoney,0) %></strong> ��</em></td>
		                </tr>
		                 <% end if %>
						<tr>
							<th scope="row">�ֹ��� ����</th>
							<td colspan="3"><%= oordermaster.FOneItem.FBuyName %> (�޴���ȭ��ȣ : <%= oordermaster.FOneItem.FBuyHp %> /  ��ȭ��ȣ : <%= oordermaster.FOneItem.FBuyPhone %>)</td>
						</tr>
						<tr>
							<th scope="row">������ ����</th>
							<td colspan="3">
								<div><%= oordermaster.FOneItem.FReqName %> (�޴���ȭ��ȣ : <%= oordermaster.FOneItem.FReqHp %> /  ��ȭ��ȣ : <%=oordermaster.FOneItem.FReqPhone  %>)</div>
								<div><%= oordermaster.FOneItem.Freqzipaddr %></div>
								<div><%= oordermaster.FOneItem.Freqaddress %></div>
							</td>
						</tr>
						<% if (oordermaster.FOneItem.IsReceiveSiteOrder) then %>
						<tr>
							<th scope="row">���� ��¥</th>
							<td colspan="3"><%= oordermaster.FOneItem.Freqdate %></td>
						</tr>
						<tr>
							<th scope="row">���� ���</th>
							<td colspan="3">
							    <!--
				                  ����� ���ı� ���̵� 88-2 �ø��� ü������� 2-1�� ����Ʈ �� �ٹ����� �����Ǹ� ������� �ν�
				                  <br>* ������ ��¥�� ������ ��ҿ����� ��ǰ���ɰ���
				                  -->
							</td>
						</tr>
						<% End If %>
						</tbody>
						</table>

						<div class="title">
							<h2>�ֹ���ǰ����</h2>
						</div>
						<table class="baseTable btmLine">
						<caption>�ֹ���ǰ���� ���</caption>
						<colgroup>
							<col width="98" /> <col width="70" /> <col width="*" /> <col width="90" /> <col width="50" /> <col width="90" /> <col width="80" />
						</colgroup>
						<thead>
						<tr>
							<th scope="col">��ǰ�ڵ�/���</th>
							<th scope="col" colspan="2">��ǰ����</th>
							<th scope="col">�ǸŰ�</th>
							<th scope="col">����</th>
							<th scope="col">�Ұ�ݾ�</th>
							<th scope="col">�ֹ�����</th>
						</tr>
						</thead>
						<tfoot>
						<tr>
							<td colspan="7">
								<div class="orderSummary">
									<span>�ֹ���ǰ�� <strong><%=oorderdetail.FTotItemKind%>�� (<%=oorderdetail.FTotItemNo%>��)</strong></span>
									<span>���� ���ϸ��� <strong><%=FormatNumber(oordermaster.FOneItem.Ftotalmileage,0)%>P</strong></span>
									<span>��ǰ���� �Ѿ� <strong><%= FormatNumber((oordermaster.FOneItem.Ftotalsum - oorderdetail.BeasongPay),0) %>��</strong></span>
								</div>
								<div class="orderTotal">
									�� �����ݾ� : ��ǰ�����Ѿ� <strong><%= FormatNumber((oordermaster.FOneItem.Ftotalsum - oorderdetail.BeasongPay),0) %></strong>�� 
									+ ��ۺ� <%= FormatNumber(oorderdetail.BeasongPay,0) %>��  

			    					<% IF (oordermaster.FOneItem.Fmiletotalprice<>0) then %>
			    					- ���ϸ��� <%= FormatNumber(oordermaster.FOneItem.Fmiletotalprice,0) %>P
			    					<% end if %>
									<% IF (oordermaster.FOneItem.Ftencardspend<>0) then %>
									- ���ʽ��������� <%= FormatNumber(oordermaster.FOneItem.Ftencardspend,0) %>��
									<% end if %> 
			    					<% if (oordermaster.FOneItem.Fallatdiscountprice + oordermaster.FOneItem.Fspendmembership<>0) then %>
			    					- ��Ÿ���� <%= FormatNumber((oordermaster.FOneItem.Fallatdiscountprice + oordermaster.FOneItem.Fspendmembership),0) %>��
			    					<% end if %> 
									= <strong class="crRed"><%= FormatNumber(oordermaster.FOneItem.FsubtotalPrice,0) %></strong>��
								</div> 
							</td>
						</tr>
						</tfoot>
						<tbody>
						 <% for i=0 to oorderdetail.FResultCount-1 %> 
						 <% if oorderdetail.FItemList(i).Fitemid <>0 then %>
						<tr>
							<td>
								<div> <%= oorderdetail.FItemList(i).FItemid %></div>
								<div>
									<% if oorderdetail.FItemList(i).Fisupchebeasong="N" then %>
									�ٹ�����
									<% elseif oorderdetail.FItemList(i).Fisupchebeasong="Y" then %>
									<font color="red">��ü����</font>
									<% end if %>
								</div>
							</td>
							<td><a href="<%=wwwUrl%>/shopping/category_prd.asp?itemid=<%= oorderdetail.FItemList(i).FItemid %>"  target="_blank"><img src="<%= oorderdetail.FItemList(i).FSmallImage %>" width="50" height="50" alt="<%= oorderdetail.FItemList(i).FItemName %>" /></a></td>
							<td class="lt">
								<div><a href="<%=wwwUrl%>/shopping/category_prd.asp?itemid=<%= oorderdetail.FItemList(i).FItemid %>" target="_blank"><%= oorderdetail.FItemList(i).FItemName %></a></div>
								<div><font color="blue"><%= oorderdetail.FItemList(i).FItemoptionName %></font></div>
							</td>
							<td>
						<% if (oorderdetail.FItemList(i).Fcancelyn <> "Y")  then %>
							<% if (oorderdetail.FItemList(i).IsSaleItem) then %>
                                    <strike><%= FormatNumber(oorderdetail.FItemList(i).Forgitemcost,0) %></strike><%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","��") %><br>
                                    <strong class="crRed"><%= FormatNumber(oorderdetail.FItemList(i).getItemcostCouponNotApplied,0) %></strong><%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","��") %>
                                <% else %>
                                    <% if (oorderdetail.FItemList(i).IsItemCouponAssignedItem) then %>
                                    <strike><%= FormatNumber(oorderdetail.FItemList(i).getItemcostCouponNotApplied,0) %></strike><%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","��") %>
                                    <% else %>
                                    <%= FormatNumber(oorderdetail.FItemList(i).getItemcostCouponNotApplied,0) %><%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","��") %>
                                    <% end if %>
                                <% end if %>

                                <% if (oorderdetail.FItemList(i).IsItemCouponAssignedItem) then %>
                                    <br><strong class="crGrn"><%= FormatNumber(oorderdetail.FItemList(i).FItemCost,0) %>��</strong>
                                <% else %>

                                <% end if %>

                                <% if (oorderdetail.FItemList(i).IsSaleBonusCouponAssignedItem) then %>
                                <p class="crRed"><img src='http://fiximage.10x10.co.kr/web2008/shoppingbag/coupon_icon.gif' width='10' height='10' > <%= FormatNumber(oorderdetail.FItemList(i).FreducedPrice,0) %><%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","��") %></p>
                                <% end if %>
                        <% else %>
                        	<font color="red">���</font>        
                        <% end if %>
							</td>
							<td><%= oorderdetail.FItemList(i).FItemNo %></td>
							<td>
						<% if (oorderdetail.FItemList(i).Fcancelyn <> "Y")  then %>
							<%= FormatNumber(oorderdetail.FItemList(i).FItemCost*oorderdetail.FItemList(i).FItemNo,0) %> <%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","��") %>
							<% if (oorderdetail.FItemList(i).IsSaleBonusCouponAssignedItem) then %>
							<p class="crRed"><img src='http://fiximage.10x10.co.kr/web2008/shoppingbag/coupon_icon.gif' width='10' height='10' > <%= FormatNumber(oorderdetail.FItemList(i).FreducedPrice*oorderdetail.FItemList(i).FItemNo,0) %><%= CHKIIF(oorderdetail.FItemList(i).IsMileShopSangpum,"Pt","��") %></p>
							<% end if %>
						 <% else %>
                        	<font color="red">���</font>        
                        <% end if %>	
							</td>
							<td><%= oorderdetail.FItemList(i).GetItemDeliverStateName(oordermaster.FOneItem.FIpkumDiv, oordermaster.FOneItem.FCancelyn) %></td>
						</tr> 
						<%end if%>
						 <% next %>
						</tbody>
						</table>
					</div>

					<div class="companyInfo">
						<p><img src="http://fiximage.10x10.co.kr/web2020/my10x10/img_company_info.png" alt="�ٹ����� 10X10 / �Ǹ�ó �ȳ� : (��)�ٹ����� ����ڵ�Ϲ�ȣ : 211-87-00620 / ��ǥ�̻� : ������ / ������ : ��)110-510 ����� ���α� ������ 1-45 �������� 5�� / �ٹ����� �����;ȳ� TEL : 1644-6030 / AM 09 :00~PM 06:00 ���ɽð� PM 12:00~01:00 �ָ�,������ �޹� / E-mail : customer@10x10.co.kr " /></p>
					</div>

					<div class="btnArea tMar30 ct">
						<button type="button" class="btn btnB1 btnWhite btnW185 lMar10" onClick="window.print();">�μ��ϱ�</button>
						<button type="button" class="btn btnB1 btnWhite btnW185 lMar10" onClick="jsGoPDF('<%=getPdfDownLinkUrlAdm%>');">PDF ��ȯ</button>
					</div>
				</div>
				<!-- //content -->
			</div>
		</div>
		<div class="popFooter">
			<div class="btnArea">
				<button type="button" class="btn btnS1 btnGry2" onclick="window.close();">�ݱ�</button>
			</div>
		</div>
	</div>
<%
set oorderdetail = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->