<% option Explicit %>

<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/academy/lib/classes/DIYShopItem/waitDIYitemCls.asp"-->

<%
'response.write "�������Դϴ�. "
'dbget.close()	:	response.End

Dim oupcheitemedit,ix,page
page = requestCheckvar(request("page"),10)

if page="" then page=1


dim i
dim cdl_mnu,cdm_mnu,cdn_mnu
dim cdl,cdm,cdn
cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cdn = requestCheckvar(request("cdn"),10)

dim cmi,cmj,k

dim itemid, oitem, designer
itemid = requestCheckvar(request("itemid"),10)

'// �߰� �̹���
dim oADD
set oADD = new CWaitItemAddImage
oADD.FRectItemId = itemid
oADD.GetOneItemAddImageList 


set oitem = New CWaitItem
oitem.FRectItemid = itemid

if (C_IS_Maker_Upche) then
    oitem.FRectMakerid = session("ssBctID")
end if

oitem.GetOneItem

'designer = oitem.FOneItem.Fmakerid

if (oitem.FResultCount<1) then
    response.write "�˻� ����� �����ϴ�."
    response.End
end if

function ImageExists(byval iimg)
	if (IsNull(iimg)) or (trim(iimg)="") or (Right(trim(iimg),1)="\") or (Right(trim(iimg),1)="/") then
		ImageExists = false
	else
		ImageExists = true
	end if
end function

'// �߰� �̹���-���� �̹���
Function getFirstAddimage()
	if ImageExists(oitem.FOneItem.FBasicImage) then
		getFirstAddimage= oitem.FOneItem.FBasicImage

	elseif (oAdd.FResultCount>0) then
		if ImageExists(oAdd.FItemList(0).FADDIMAGE_400) then
			getFirstAddimage= oAdd.FItemList(0).FADDIMAGE_400
		end if
	else
		getFirstAddimage= oitem.FOneItem.FMainImage
	end if
end Function

dim iOptionBoxHTML

iOptionBoxHTML = getOptionBoxHTML_FrontType(oitem.FOneItem.FWaititemid)

%>
<html>
<head>
<title> ���� ��ſ� ���θ�, ����ä�� �ٹ����� 10X10</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<meta name="description" content="Ŀ��10x10������ Ŀ�õ��� ���� ����, ������ ��ǰ,����, ���̵���ǰ ���� �������� �ϴ� �����̳ʵ��� ��Ÿ�ϼ��Դ� ">
<meta name="keywords" content="Ŀ��, ����, Ŀ�ü���, ����������, ������,���̵���ǰ, �����ο�ǰ, ����, ��Ÿ��, 10x10, �ٹ�����, ť��">
<meta name="classification" content="����Ͻ��� ����,���ΰ� ����(B2C, C2C),����, Ư����ǰ">
<link rel=stylesheet type="text/css" href="http://www.10x10.co.kr/lib/css/2008ten.css">
<!-- script language="JavaScript" SRC="http://www.10x10.co.kr/js/tenbytencommon.js"></script -->

<script language="JavaScript">
// �߰� �̹��� ����
	var btnT='btnB'; // ���õ� �߰��̹��� ������
	function TnSwitchImageBG(imagesrc,btn){
		document.getElementById("IimageMain").style.backgroundImage = "url('" + imagesrc + "')";

		// ���õ� �̹����� #FF0000 �׿� #B7B7B7
		try {
			//document.getElementById(btnT).innerHTML='';
		}	catch(e) {
			//document.getElementById('btn0').innerHTML='<img src="http://fiximage.10x10.co.kr/web2008/category/today_redline.gif" border="0" width="38" height="38">';
		}

		//document.getElementById(btn).innerHTML='<img src="http://fiximage.10x10.co.kr/web2008/category/today_redline.gif" border="0" width="38" height="38">';

		btnT=btn;
	}
	
<!--�귣����ü�˻�����-->
	function GoToBrandShop(designerid){
		if (designerid == "")
		{
			alert("�귻�尡 �����ϴ�.")
			}
				var popup = window.open('http://www.10x10.co.kr/street/brandshop.asp?designerid=' + designerid , 'popupedit' , 'width=1024 height=768 scrollbars=yes');		
				popup.focus();					
			}
<!--�귣����ü�˻���-->

function CLargeMenuOpen(){
	ctg.style.visibility="visible";
}
function CLargeMenuClose(){
	ctg.style.visibility="hidden";
}

function OverStay() {
	var i,x,a=document.sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function search_clear(){
	document.searchForm.rect.value = "";
}
</script>
<STYLE type="text/css">
	#ctg{
		Z-INDEX: 1; BACKGROUND: white; FILTER: alpha(opacity=95); LEFT: 7px; VISIBILITY: hidden; OVERFLOW: hidden; WIDTH: 180px; POSITION: absolute; TOP: 25 px;
	}
</STYLE>
</head>


<script language="JavaScript">
	function AddFavorite(){
		// nothing
	}
	function GoJorgi(itemname,maker,sellcash,listimg,itemid){
		// // nothing
	}
function FrameControl(imagesrc,imgstr){
	 itemimgview.TnFrameChangeImage(imagesrc);
	 document.getElementById("IimageText").innerHTML = imgstr;
}

function TnBTNChange(id,Max){

 var idnum = id.substring(3,4);

  for(i=0;i<=Max;i++){
      if (idnum == i){
		  eval('document.getElementById("btn' + i + '").src  ="http://fiximage.10x10.co.kr/images/shopping/add_0' + (i + 1) + '.gif"');
      }
	  else{
		  eval('document.getElementById("btn' + i + '").src  ="http://fiximage.10x10.co.kr/images/shopping/add_b0' + (i + 1) + '.gif"');
	  }
  }
}
</script>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" rightmargin="0" bottommargin="0" bgcolor="#FFFFFF">
<center>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td style="padding:20px 0 0 0">
    <!----- ��ǰ�⺻���� START ---->
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
    	<td valign="top" style="padding-left:20px;">
    		<table border="0" cellspacing="0" cellpadding="0">
    		<tr>
    			<td style="background-repeat: no-repeat;background-position: center;" id="IimageMain" background="<% = getFirstAddimage %>">
    				<img src="http://fiximage.10x10.co.kr/images/spacer.gif" id="IimageAdd" width="400" height="400"></td>
    				
    		</tr>
    		<tr>
    			<td style="padding-top:7px;">
    				<!-- //  ��ǰ �̹��� 50x50 -->
    				<table border="0" cellspacing="0" cellpadding="0">
    				<tr>
    					<% if ImageExists(oitem.FOneItem.FBasicImage) then %>
    					<td style="padding-right:5px;">
    						<table border="0" cellspacing="1" cellpadding="1"  onclick="TnSwitchImageBG('<%= oitem.FOneItem.FBasicImage %>','btnB');" style="cursor:pointer" bgcolor="#CCCCCC">
    						<tr>
    							<td width="38" height="38"  id="btnB"><img src="<%= oitem.FOneItem.FBasicImage %>" width=38 height=38></td>
    						</tr>
    						</table>
    					</td>
    					<% end if %>
    					<% IF oAdd.FResultCount > 0 THEN %>
    					<% FOR i= 0 to oAdd.FResultCount-1  %>
    					<%IF oAdd.FItemList(i).FIMGTYPE=0 and ImageExists(oAdd.FItemList(i).FADDIMAGE_400) THEN %>
    					<td style="padding-right:5px;">
    						<table border="0" cellspacing="1" cellpadding="1" onclick="TnSwitchImageBG('<%= oAdd.FItemList(i).FADDIMAGE_400 %>','btn<%=i %>');" style="cursor:pointer" bgcolor="#CCCCCC">
    						<tr>
    							<td width="38" height="38" align="center" id="btn<%=i %>" ><img src="<%= oAdd.FItemList(i).FADDIMAGE_400 %>" width=38 height=38></td>
    						</tr>
    						</table>
    					</td>
    					<%End IF %>
    					<% NEXT %>
    					<% END IF %>
    				</tr>
    				</table>
    				<!--  ��ǰ �̹��� 50x50 //-->
    			</td>
    		</tr>
    		<tr>
    		    <td>(������ �̹����� �� ��� �� ���� �˴ϴ�.)</td>
    		</tr>
    		</table>
    	</td>
    	<td align="right" valign="top">
    		<table width="310" border="0" cellspacing="0" cellpadding="0">
    		
    		<tr>
    			<td height="350" valign="top">
    				<table width="100%" border="0" cellspacing="0" cellpadding="0">
    				<tr>
    					<td>
    						<table width="100%" border="0" cellspacing="0" cellpadding="0">
    						<!-- �귣��� -->
    						<tr>
    							<td>
    								<table border="0" cellspacing="0" cellpadding="0">
    								<tr>
    									<td style="padding-right:5px;" class="eng14pxgray"><%= UCase(oitem.FOneItem.FBrandName) %></td>
    									<% If oitem.FOneItem.IsStreetAvail Then %>
    										<td valign="bottom" style="padding-bottom:2px;" valign="bottom"><a href="#>" onFocus="blur()"><img src="http://fiximage.10x10.co.kr/web2008/category/btn_brandmore02.gif" width="90" height="16" border="0"></a></td>
    									<% End If %>	
    								</tr>
    								</table>
    							</td>
    						</tr>
    						<!-- ��ǰ�� -->
    						<tr>
    							<td class="prodtitle" style="padding-top:2px;"><%= oitem.FOneItem.FItemName %></td>
    						</tr>
    						<!-- ����/����.. ������ -->
    						<tr>
    							<td style="padding-top:4px;">
    								<table border="0" cellspacing="0" cellpadding="0">
    								<tr>
    									<% if oitem.FOneItem.IsSoldOut then %>
    										<td style="padding-right:5px;"><img src="http://fiximage.10x10.co.kr/web2008/category/icon_soldout02.gif" width="30" height="12"></td>
    									<% else %>
    										<% if oitem.FOneItem.IsSaleItem then %>
    											<td style="padding-right:5px;"><img src="http://fiximage.10x10.co.kr/web2008/category/icon_sale02.gif" width="30" height="12"></td>
    										<% end if %>
    
    										<% if oitem.FOneItem.IsLimitItem then %>
    											<td style="padding-right:5px;"><img src="http://fiximage.10x10.co.kr/web2008/category/icon_limit02.gif" width="30" height="12"></td>
    										<% end if %>
    										
    									<% end if %>
    								</tr>
    								</table>
    							</td>
    						</tr>
    						</table>
    					</td>
    				</tr>
    				<tr>
    					<td style="padding:15px 0 9px 0;border-bottom:1px solid #eaeaea;">
    						<table width="100%" border="0" cellspacing="0" cellpadding="0">
    						<!-- �Һ��ڰ� -->
    						<tr>
    							<td width="92" height="21"><img src="http://fiximage.10x10.co.kr/web2008/category/text_price.gif" width="37" height="10"></td>
    							<td>
    								<table border="0" cellspacing="0" cellpadding="0">
    								<tr>
    									<td class="black11pxb" style="padding:2px 5px 0 0;"><%= FormatNumber(oitem.FOneItem.getOrgPrice,0) %> ��</td>
    									<td><%= oitem.FOneItem.getInterestFreeImg %></td>
    								</tr>
    								</table>
    							</td>
    						</tr>
    						<!-- ���� �ǸŰ� -->
    						<% IF oitem.FOneItem.IsSaleItem THEN %>
    						<tr>
    							<td height="21"><img src="http://fiximage.10x10.co.kr/web2008/category/text_sale.gif" width="46" height="10"></td>
    							<td class="sale11px01" style="padding-top:2px;"><%= FormatNumber(oitem.FOneItem.getRealPrice,0) %>�� [<% = oitem.FOneItem.getSalePro %>]</td>
    						</tr>
    						<% End If %>
    						
    						<!-- ���ϸ��� -->
    						<tr>
    							<td height="21"><img src="http://fiximage.10x10.co.kr/web2008/category/text_point.gif" width="35" height="10"></td>
    							<td class="gray11px02" style="padding-top:2px;"><strong><% = oitem.FOneItem.FMileage %> Point</strong></td>
    						</tr>
    						</table>
    					</td>
    				</tr>
    				
    				<tr>
    					<td style="padding:10px 0 9px 0;border-bottom:1px solid #eaeaea;">
    						<!-- ��ǰ�ڵ� & ������/������ & ��۱��� -->
    						<table width="100%" border="0" cellspacing="0" cellpadding="0">
    						<tr>
    							<td width="92" height="21"><img src="http://fiximage.10x10.co.kr/web2008/category/text_code.gif" width="37" height="10"></td>
    							<td class="gray11px02" style="padding-top:2px;"></td>
    						</tr>
    						<tr>
    							<td height="21"><img src="http://fiximage.10x10.co.kr/web2008/category/text_company.gif" width="59" height="10"></td>
    							<td class="gray11px02" style="padding-top:2px;"><% = oitem.FOneItem.FMakerName %> / <% = oitem.FOneItem.FSourceArea %></td>
    						</tr>
    						<tr>
    							<td height="21"><img src="http://fiximage.10x10.co.kr/web2008/category/text_delivery.gif" width="36" height="10"></td>
    							<td class="gray11px02" style="padding-top:2px;">
    								<table border="0" cellpadding="0" cellspacing="0">
    									<tr>
    										<td><% = oitem.FOneItem.GetDeliveryName %></td>
    										<td style="padding-left:10;">
    										<% if oitem.FOneItem.IsFreeBeasong then %>
    										<img src="http://fiximage.10x10.co.kr/web2008/category/icon_free02.gif" width="38" height="12">
    										<% end if %>
    										
    										<% if (oitem.FOneItem.IsUpcheParticleDeliverItem) then %>
    										<img src="http://fiximage.10x10.co.kr/web2008/category/btn_chargeinfo.gif" border="0" width="57" height="13" onClick="fnShowDeliveryNotice();" onmouseout="fnHideDeliveryNotice();" style="cursor:pointer">
    										<div id="layer_dlv" style="Display:none;Position:absolute; width:190px;margin-top:13px;margin-left:-185px "  onmouseover="fnShowDeliveryNotice();" onmouseout="fnHideDeliveryNotice();">
    											<table width="100%" border="0" cellspacing="0" cellpadding="0">
    											<tr>
    												<td bgcolor="#FFFFFF" style="padding:12px;border:4px solid #eeeeee;" class="gray11px02"><%= oitem.FOneItem.getDeliverNoticsStr %></td>
    											</tr>
    											</table>
    										</div>
    										<% end if %>
    										</td>
    									</tr>
    								</table>
    							</td>
    						</tr>
    						</table>
    					</td>
    				</tr>
    				<tr>
    					<td style="padding:8px 0 9px 0;border-bottom:1px solid #eaeaea;">
    						<!-- �ֹ� ���� -->
    						<table width="100%" border="0" cellspacing="0" cellpadding="0">
    						<tr>
    							<td width="92" height="21"><img src="http://fiximage.10x10.co.kr/web2008/category/text_ordernum.gif" width="37" height="10"></td>
    							<td class="gray11px02" style="padding-top:1px;"><label>
    								<input name="itemea" type="text" class="input_01" style="width:37px;height:17px;" value="1">
    								ea</label>
    							</td>
    						</tr>
    						<!-- �ɼ� -->
    						<% IF oitem.FOneItem.FOptionCnt>0 then %>
    						<tr>
    							<td valign="top"  style="padding-top:8px;"><img src="http://fiximage.10x10.co.kr/web2008/category/text_option.gif" width="36" height="10"></td>
    							<td style="padding:3px 0 1px 0;"><%= ioptionBoxHtml %></td>
    						</tr>
    						<% End If %>
    						<% if oitem.FOneItem.FItemDiv = "06" then %>
    						<tr>
    							<td valign="top"  style="padding-top:8px;"><img src="http://fiximage.10x10.co.kr/web2008/category/text_text.gif" width="46" height="10"></td>
    							<td style="padding-top:1px;">
    								<textarea name="requiredetail" id="requiredetail" cols="45" rows="3" class="input_01" style="width:200px;"></textarea>
    							</td>
    						</tr>
    						<% End If %>
    						
    						</table>
    					</td>
    				</tr>
    				<% if oitem.FOneItem.isLimitItem Then %>
    				<tr>
    					<td style="padding:10px 0 8px 0;">
    						<!-- ���� �Ǹ� ��ǰ -->
    						<table width="100%" border="0" cellspacing="0" cellpadding="0">
    						<tr>
    							<td width="92" ><img src="http://fiximage.10x10.co.kr/web2008/category/text_limit.gif" width="55" height="10"></td>
    							<td class="gray11px02" style="padding-top:2px;"><strong><% = oitem.FOneItem.FRemainCount %></strong>�� ���ҽ��ϴ�.</td>
    						</tr>
    						
    						</table>
    					</td>
    				</tr>
    				<% End If %>
    				<% IF oitem.FOneItem.FAvailPayType="1" THEN %>
    				<tr>
    					<td style="padding:10px 0 8px 0;">
    						<!-- ������ �ǸŻ�ǰ  -->
    						<table width="100%" border="0" cellspacing="0" cellpadding="0">
    						<tr>
    							<td width="92" height="21" valign="top"><img src="http://fiximage.10x10.co.kr/web2008/category/text_speedsale.gif" width="64" height="10"></td>
    							<td class="gray11px02" valign="top" style="padding-top:0px;">������ �Ǹ� ��ǰ <br>- �ǽð�(ī������)���θ� �����ϽǼ� �ֽ��ϴ�.</td>
    						</tr>
    						</table>
    					</td>
    				</tr>	
    					
    				<% End IF %>
    				</table>
    			</td>
    		</tr>
    		<tr>
    			<td>
    				<!-- ���� ��ư -->
    				<table border="0" cellspacing="0" cellpadding="0">
    				<tr>
    					<% if oitem.FOneItem.isSoldout then %>
    					<td style="padding-right:6px;"><img src="http://fiximage.10x10.co.kr/web2008/category/btn_soldout.gif" width="178" height="50"></td>
    					<% Else %>
    					<td style="padding-right:6px;"><a href="#" onFocus="blur()"><img src="http://fiximage.10x10.co.kr/web2008/category/btn_noworder.gif" width="86" height="50" border="0"></a></td>
    					<td style="padding-right:6px;"><a href="#" onFocus="blur()"><img src="http://fiximage.10x10.co.kr/web2008/category/btn_cart.gif" width="86" height="50" border="0"></a></td>
    					<% End If %>
    					<td style="padding-right:6px;"><a href="#" onFocus="blur()"><img src="http://fiximage.10x10.co.kr/web2008/category/btn_wishlist.gif" width="60" height="50" border="0"></a></td>
    					<td><a href="#;" onFocus="blur()"><img src="http://fiximage.10x10.co.kr/web2008/category/btn_letter.gif" width="60" height="50" border="0"></a></td>
    				</tr>
    				</table>	
    			</td>
    		</tr>
    		<tr>
    			<td height="59" style="padding-top:3px;">
    				<!--  ��ȹ �ڳ� & �̺�Ʈ -->
    			</td>
    		</tr>
    		</table>
    	</td>
    </tr>
    </table>
    <!----- ��ǰ�⺻���� END ---->
    </td>
</tr>
<tr>
	<td style="padding-top:20px;">
		<!-- tab START -->
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td height="30" align="right" valign="bottom" bgcolor="#f8f8f8" style="border-bottom:1px solid #dddddd;padding-right:9px">
				<table border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td><a href="#det" onFocus="blur()"><img src="http://fiximage.10x10.co.kr/web2008/category/view_tab01.gif" width="94" height="21" border="0"></a></td>
					<td><a href="#not" onFocus="blur()"><img src="http://fiximage.10x10.co.kr/web2008/category/view_tab02.gif" width="94" height="21" border="0"></a></td>
					<td width="94" background="http://fiximage.10x10.co.kr/web2008/category/view_tabbg.gif">
						<table border="0" cellspacing="0" cellpadding="0">
		  				<tr>
		    				<td style="padding:2px 4px 0 7px"><a href="#eva" onFocus="blur()"><img src="http://fiximage.10x10.co.kr/web2008/category/view_tabtext03.gif" width="35" height="10" border="0"></a></td>
		    				<td style="padding-top:2px;"><a href="#eva">(<strong>0</strong>)</a></td>
		  				</tr>
						</table>
					</td>
					<td width="94" background="http://fiximage.10x10.co.kr/web2008/category/view_tabbg.gif">
						<table border="0" cellspacing="0" cellpadding="0">
		    			<tr>
		      				<td style="padding:2px 4px 0 7px"><a href="#qa" onFocus="blur()"><img src="http://fiximage.10x10.co.kr/web2008/category/view_tabtext04.gif" width="28" height="10" border="0"></a></td>
		      				<td style="padding-top:2px;"><a href="#qa">(<strong>0</strong>)</a></td>
		    			</tr>
						</table>
					</td>
					<td><a href="#bb" onFocus="blur()"><img src="http://fiximage.10x10.co.kr/web2008/category/view_tab05.gif" width="95" height="21" border="0"></a></td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
		<!-- tab END -->
	</td>
</tr>
<tr>
	<td style="padding:20px 20px 10px 20px;">
		
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td style="padding-bottom:30px;">
				<!-- // �ֹ� ���ǻ��� -->
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td style="padding-bottom:10px;"><img src="http://fiximage.10x10.co.kr/web2008/category/view_text_01.gif" width="74" height="13"></td>
				</tr>
				<% IF (oItem.FOneItem.getDeliverNoticsStr<>"") THEN %>
				<tr>
					<td class="gray11px02" style="padding-top:2px;padding-bottom:10">
                    <%= oItem.FOneItem.getDeliverNoticsStr %>
					</td>
				</tr>
				<% End IF %>
				<tr>
					<td class="gray11px02" style="padding-top:2px;"><%= nl2br(oItem.FOneItem.FOrderComment) %></td>
				</tr>
				
				</table>
				<!-- �ֹ� ���ǻ��� //-->
			</td>
		</tr>
		<tr>
			<td style="padding-bottom:30px;">
				<!-- // ��ǰ ��ǰ���� & �� �̹��� Start-->
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td style="padding-bottom:10px;"><a name="det"></a><img src="http://fiximage.10x10.co.kr/web2008/category/view_text_02.gif" width="50" height="13"></td>
				</tr>
				<tr>
					<td>
						<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<!-- ��� -->
						<tr>
							<td width="56" height="21" align="center" style="border-bottom:1px solid #eaeaea;"><img src="http://fiximage.10x10.co.kr/web2008/category/view_text_02a.gif" width="18" height="10"></td>
							<td class="gray11px02" style="border-bottom:1px solid #eaeaea;padding-top:3px;"><% = oItem.FOneItem.FItemSource %></td>
						</tr>
						<!-- ������ -->
						<tr>
							<td height="21" align="center" style="border-bottom:1px solid #eaeaea;"><img src="http://fiximage.10x10.co.kr/web2008/category/view_text_02b.gif" width="19" height="10"></td>
							<td class="gray11px02" style="border-bottom:1px solid #eaeaea;padding-top:3px;"><% = oItem.FOneItem.FItemSize %></td>
						</tr>
						<!-- �߷� -->
						<tr>
							<td height="21" align="center" style="border-bottom:1px solid #eaeaea;"><img src="http://fiximage.10x10.co.kr/academy2010/diyshop/view_text_02c.gif" width="19" height="10"></td>
							<td class="gray11px02" style="border-bottom:1px solid #eaeaea;padding-top:3px;"><% = oItem.FOneItem.FItemWeight %>g</td>
						</tr>
						<!-- �±� -->
						<tr>
							<td height="21" align="center" style="border-bottom:1px solid #eaeaea;"><img src="http://fiximage.10x10.co.kr/web2008/category/view_text_02c.gif" width="19" height="10"></td>
							<td class="gray11px02" style="border-bottom:1px solid #eaeaea;padding-top:3px;"><% = oItem.FOneItem.FKeyWords %>&nbsp;</td>
						</tr>
						</table>
					</td>
				</tr>
				
				<tr>
					<td style="padding-top:30px;" class="gray11px02">
						<% 
						IF oItem.FOneItem.FUsingHTML="Y" THEN 
							Response.write oItem.FOneItem.FItemContent
						ELSEIF oItem.FOneItem.FUsingHTML="H" THEN
							Response.write "<span class=""gray11px02"">" & nl2br(oItem.FOneItem.FItemContent) & "</span>"
						ELSE
							Response.write "<span class=""gray11px02"">" & nl2br(ReplaceBracket(oItem.FOneItem.FItemContent)) & "</span>"
						END IF 
						%>
						<% IF oAdd.FResultCount > 0 THEN %>
							<% FOR i= 0 to oAdd.FResultCount-1  %>
								<%IF oAdd.FItemList(i).FIMGTYPE=1 THEN %>
									<img src="<%= oAdd.FItemList(i).FADDIMAGE_400 %>" border="0"><br>
								<%End IF %>
							<% NEXT %>
						<% END IF %>
						<% if ImageExists(oItem.FOneItem.FMainImage) then %>
							<img src="<% = oItem.FOneItem.FMainImage %>" border="0" id="filemain" style="cursor:pointer;" onclick="ViewOrgImage('<% = oItem.FOneItem.FMainImage %>');" >
						<% end if %>
						<!--<img src="http://fiximage.10x10.co.kr/web2008/category/img_detail.jpg" width="620" height="1087">-->
					</td>
				</tr>
				</table>
				<!-- ��ǰ ��ǰ���� & �� �̹��� End //-->
			</td>
		</tr>
		
		<tr>
			<td style="padding-bottom:30px;">
				<!-- // ��� ��ȯ ȯ�� ���� -->
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td style="padding-bottom:10px;"><a name="not"></a><img src="http://fiximage.10x10.co.kr/web2008/category/view_text_03.gif" width="80" height="14"></td>
				</tr>
				<tr>
					<td><img src="http://fiximage.10x10.co.kr/web2008/category/view_delivery.gif" width="720" height="245"></td>
				</tr>
				</table>
				<!-- ��� ��ȯ ȯ�� ���� //-->
			</td>
		</tr>
        </table>
    </td>
    </tr>
</table>

</center>
</body>
</html>
<%
set oitem = Nothing
set oADD = Nothing
%>
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->