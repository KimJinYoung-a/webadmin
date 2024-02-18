<%
session.codePage = 65001
%>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 온라인 바코드 출력
' Hieditor : 2016.12.15 한용민 생성
'###########################################################
%>
<%
dim divcd ,companyid ,userid, itemname, isfixed, jumunwait, IsForeign_confirmed, IsForeignOrder
dim defaultlocationid ,barcodetype ,barcodetypestring, sellyn, usingyn
dim locationidfrom ,locationnamefrom ,locationidto ,locationnameto
dim IsOneOrderOnly ,siteSeq ,innerboxidx ,innerboxweight ,cartonboxweight , shopseq
	listgubun 		= requestCheckVar(request("listgubun"), 32)
	divcd = requestCheckVar(request("divcd"),32)
	'companyid = requestCheckVar(trim(request("companyid")),32)
	companyid = requestCheckVar(session("ssBctID"), 32)
	barcodetype = requestCheckVar(request("barcodetype"),32)
	isforeignprint = requestCheckVar(request("isforeignprint"),1)
	page 			= requestCheckVar(request("page"),32)
	makerid = requestCheckVar(request("makerid"),32)
	itemid      = requestCheckvar(request("itemid"),255)
	itemname    = requestCheckvar(request("itemname"),64)
	prdcode 		= requestCheckVar(request("prdcode"),32)
	generalbarcode 	= requestCheckVar(request("generalbarcode"),32)
	sellyn      = requestCheckvar(request("sellyn"),10)
	usingyn     = requestCheckvar(request("usingyn"),10)
	research 		= requestCheckVar(request("research"),32)
	itembarcodearr = request("itembarcodearr")
	printpriceyn = requestCheckVar(request("printpriceyn"),1)
	makeriddispyn = requestCheckVar(request("makeriddispyn"),1)
	papername = requestCheckVar(request("papername"),2)
	itemoptionyn = requestCheckVar(request("itemoptionyn"),1)

isdispsql=true
isdispconfirm=true
if papername = "" then papername = "BQ"
jumunwait = false
IsForeignOrder = false		'/업체접수주문
IsForeign_confirmed = false		'/업체접수주문 컨펌완료여부

'/매장일경우
if (C_IS_SHOP) then
	'/가맹점 일경우
	if getoffshopdiv(C_STREETSHOPID) = "3" then
		isdispconfirm=false
	end if
else
	if (C_IS_Maker_Upche) then
		makerid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
		else
		end if
	end if
end if
if page = "" then page = 1
iPageSize=50
'if sellyn = "" and research <> "on" then sellyn = "Y"
if listgubun = "" then listgubun = "ITEM"
if listgubun="ITEM" then
	if usingyn = "" and research <> "on" then usingyn = "Y"	
end if
if printpriceyn = "" then printpriceyn = "Y"
if itemoptionyn = "" then itemoptionyn = "Y"
siteSeq = "10"

if itemid<>"" then
  itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

set oproduct = new CStorageDetail

'/상품리스트
if listgubun = "ITEM" then
	oproduct.FPageSize = iPageSize
	oproduct.FCurrPage = page
	oproduct.FRectMakerid = makerid
	oproduct.FRectItemid       = itemid
	oproduct.FRectItemName     = itemname
	oproduct.FRectPrdCode = prdcode
	oproduct.FRectGeneralBarcode = generalbarcode
	oproduct.FRectisforeignprint = isforeignprint
	oproduct.FRectSellYN       = sellyn
	oproduct.FRectIsUsing      = usingyn
	oproduct.frectitembarcodearr = itembarcodearr

	if makerid<>"" or itemname<>"" or prdcode<>"" or itemid<>"" or generalbarcode<>"" then
		oproduct.GetProductListOnline
	else
		isdispsql=false
	end if
end if

if isforeignprint="" then isforeignprint="N"
if currencyunit="" then
	if isforeignprint="N" then
		currencyunit = "KRW"
	else
		currencyunit = "USD"
	end if
end if
if currencyChar="" then
	if isforeignprint="N" then
		currencyChar = "￦"
	else
		currencyChar = "$"
	end if
end if

wd = 80
ht = 80
qt = "M"

%>

<script type="text/javascript">

function reg(page){
	frm.page.value=page;
	frm.action='';
	frm.target='';
	frm.method="post"
	frm.submit();
}

</script>

<table align="left" valign="top" cellpadding="0" cellspacing="0" border="0">

<% if not(isdispconfirm) then %>
	<tr bgcolor="#FFFFFF">
		<td align="center">
			<font color="red"><strong>온라인 상품정보 조회 권한이 없습니다.</strong></font>
		</td>
	</tr>
<% elseif oproduct.FresultCount > 0 then %>
	<%
	'/물류코드, 범용바코드
	if papername="T" or papername="G" then
	%>
		<tr bgcolor='#FFFFFF'>
			<% for i=0 to oproduct.FresultCount-1 %>
			<% tmptdcnt = tmptdcnt + 1 %>
			<td style='width:208px; height:133px;' valign='top'>
				<table align='left' cellpadding='0' cellspacing='0' border='0' valign='top' width='100%' height='100%'>
				<tr valign='top' align='left'>
					<td height=20>
						<% if makeriddispyn="Y" then %>
							<strong><span class="currencychar10X10" style='font-size:18px;'><%= oproduct.FItemList(i).fsocname %></span></strong>
						<% end if %>
					</td>			
				</tr>
				<tr valign='top' align='left'>
					<td height=20 style="vetical-align:top; line-height:6px;">
						<% if makeriddispyn="Y" then %>
							<strong><span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).fsocname_kor %></span></strong>
						<% end if %>
					</td>
				</tr>
				<tr valign='top' align='left'>
					<td style="vetical-align:top; line-height:10px;">
						<% if isforeignprint = "Y" then %>
							<span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).Flcitemname %></span>

							<% if itemoptionyn="Y" then %>
								<% if oproduct.FItemList(i).Flcitemoptionname <> "" then %>
									<br><span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).Flcitemoptionname %></span>
								<% end if %>
							<% end if %>
						<% else %>
							<span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).Fprdname %></span>

							<% if itemoptionyn="Y" then %>
								<% if oproduct.FItemList(i).Fitemoptionname <> "" then %>
									<br><span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).Fitemoptionname %></span>
								<% end if %>
							<% end if %>
						<% end if %>
					</td>
				</tr>
				<tr valign='top' align='left'>
					<td style="vetical-align:top; line-height:6px;">
					</td>
				</tr>
				<tr valign='top' align='left'>
					<td height=40>
						<table align='left' cellpadding='0' cellspacing='0' border='0' valign='top' width='100%' height='100%'>
						<tr style='padding-bottom:5px'>
							<td align='left' valign='bottom' style="vetical-align:bottom; line-height:15px;>
								<% if printpriceyn = "Y" or printpriceyn = "C" or printpriceyn="R" or printpriceyn="S" then %>
									<% if isforeignprint = "Y" then %>
										<strong><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= round(oproduct.FItemList(i).Flcprice,2) %></span></strong>
									<% else %>
										<%
										'//할인가 표시
										if printpriceyn="C" then
										%>
											<%
											'/할인 처리
											if oproduct.FItemList(i).Fsaleyn="Y" and oproduct.FItemList(i).Fcustomerprice>oproduct.FItemList(i).Flcprice then
											%>
												<strong><span class="currencychardefault" style='font-size:8px; text-decoration:line-through;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:8px; text-decoration:line-through;'><%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %></span>
												<br><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).Flcprice,0) %></span></strong>
											<%
											'/쿠폰 처리
											elseif oproduct.FItemList(i).FitemCouponYn="Y" and oproduct.FItemList(i).Fcustomerprice>oproduct.FItemList(i).GetCouponAssignPrice then
											%>
												<strong><span class="currencychardefault" style='font-size:8px; text-decoration:line-through;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:8px; text-decoration:line-through;'><%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %></span>
												<br><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).GetCouponAssignPrice,0) %></span></strong>
											<% else %>
												<strong><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %></span></strong>
											<% end if %>
										<%
										'/판매가 표시
										elseif printpriceyn="R" then
										%>
											<strong><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).Fsellprice,0) %></span></strong>
										<%
										'/심플금액표시
										elseif printpriceyn="S" then
										%>
											<strong><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %></span></strong>
										<%
										'//소비자가 표시
										else
										%>
											<strong><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %></span></strong>
										<% end if %>
									<% end if %>
								<% end if %>
							</td>
							<td align='right' valign='bottom' style='padding-right:8px'>
								<%
								'/물류코드
								if papername="T" then
								%>
									<img src="http://company.10x10.co.kr/barcode/barcode.asp?image=3&type=25&height=25&barwidth=1&ClearAreaLeft=0&ClearAreaRight=0&ClearAreaTop=0&ClearAreaBottom=0&ClearAreaMiddle=0&Size=6&data=<%= BF_MakeTenBarcode(oproduct.FItemList(i).Fitemgubun, oproduct.FItemList(i).Fitemid, oproduct.FItemList(i).Fitemoption) %>&caption=<%= BF_GetItemGubun(oproduct.FItemList(i).Fitemgubun) %>-<%= BF_GetFormattedItemId(oproduct.FItemList(i).Fitemid) %>-<%= BF_GetItemOption(oproduct.FItemList(i).Fitemoption) %>" alt="<%= BF_MakeTenBarcode(oproduct.FItemList(i).Fitemgubun, oproduct.FItemList(i).Fitemid, oproduct.FItemList(i).Fitemoption) %>" />
								<%
								'/범용바코드
								elseif papername="G" then
								%>
									<% if oproduct.FItemList(i).Fgeneralbarcode then %>
										<img src="http://company.10x10.co.kr/barcode/barcode.asp?image=3&type=25&height=25&barwidth=1&ClearAreaLeft=0&ClearAreaRight=0&ClearAreaTop=0&ClearAreaBottom=0&ClearAreaMiddle=0&Size=6&data=<%= oproduct.FItemList(i).Fgeneralbarcode %>&caption=<%= oproduct.FItemList(i).Fgeneralbarcode %>" alt="<%= oproduct.FItemList(i).Fgeneralbarcode %>" />
									<% end if %>
								<% end if %>
							</td>
						</tr>
						</table>
					</td>			
				</tr>
				</table>
			</td>
			<%
			'/ 세로칸 중간에 공백 줌
			if tmptdcnt=1 or tmptdcnt=2 then
				response.write "<td style='width:19px; height:133px;'>&nbsp;</td>"
			end if
			%>
			<%
			'/ 3개 넘으면 줄내림
			if tmptdcnt >= 3 then
				tmptrcnt = tmptrcnt + 1

				if (oproduct.FresultCount/3) <> tmptrcnt then
					response.write "</tr>" & vbcrlf
					response.write "<tr bgcolor='#FFFFFF'><td colspan=5 style='height:15px;'>&nbsp;</td></tr>" & vbcrlf
					response.write "<tr bgcolor='#FFFFFF'>" & vbcrlf
				end if

				tmptdcnt = 0
			end if
			%>
			<% next %>
		</tr>

	<%
	'/장바구니 쇼카드(QR코드), QR코드, 이미지, 쇼카드만
	'elseif papername="BQ" or papername="Q" or papername="I" or papername="" then
	else
	%>
		<tr bgcolor='#FFFFFF'>
			<% for i=0 to oproduct.FresultCount-1 %>
			<%
			'/ 장바구니 쇼카드(QR코드)
			if papername="BQ" then
				if oproduct.FItemList(i).Fitemgubun="10" then
					'msg = "http://m.10x10.co.kr/offshop/view/category_prd.asp?barcode=" & BF_MakeTenBarcode(oproduct.FItemList(i).Fitemgubun,oproduct.FItemList(i).Fitemid,oproduct.FItemList(i).Fitemoption)
					msg = "http://m.10x10.co.kr/offshop/view/category_prd.asp?barcode=" & oproduct.FItemList(i).Fprdbarcode

					'// 구글 Chart API - QRCode URL (반드시 UTF-8로 전송)
					imgPath = "http://chart.apis.google.com/chart?cht=qr&chl=" & URLEncodeUTF8(msg) & "&choe=UTF-8&chs=" & wd & "x" & ht & "&chld=" & qt & "|1"
				else
					IF application("Svr_Info")="Dev" THEN
						imgPath = oproduct.FItemList(i).Fmainimageurl
					else
						imgPath = getThumbImgFromURL(oproduct.FItemList(i).Fmainimageurl,wd,ht,"true","false")
					end if
				end if

			'/QR코드
			elseif papername="Q" then
				if oproduct.FItemList(i).Fitemgubun="10" then
					'msg = "http://m.10x10.co.kr/category/category_itemPrd.asp?itemid=" & oproduct.FItemList(i).Fitemid
					msg = "http://m.10x10.co.kr/offshop/view/iteminfo.asp?itemid=" & oproduct.FItemList(i).Fitemid


					'// 구글 Chart API - QRCode URL (반드시 UTF-8로 전송)
					imgPath = "http://chart.apis.google.com/chart?cht=qr&chl=" & URLEncodeUTF8(msg) & "&choe=UTF-8&chs=" & wd & "x" & ht & "&chld=" & qt & "|1"
				else
					IF application("Svr_Info")="Dev" THEN
						imgPath = oproduct.FItemList(i).Fmainimageurl
					else
						imgPath = getThumbImgFromURL(oproduct.FItemList(i).Fmainimageurl,wd,ht,"true","false")
					end if
				end if

			'/이미지
			elseif papername="I" THEN
				if oproduct.FItemList(i).Fitemgubun="10" then
					IF application("Svr_Info")="Dev" THEN
						imgPath = oproduct.FItemList(i).Fmainimageurl
					else
						imgPath = getThumbImgFromURL(oproduct.FItemList(i).Fmainimageurl,wd,ht,"true","false")
					end if
				else
					IF application("Svr_Info")="Dev" THEN
						imgPath = oproduct.FItemList(i).Fmainimageurl
					else
						imgPath = getThumbImgFromURL(oproduct.FItemList(i).Fmainimageurl,wd,ht,"true","false")
					end if
				end if

			'/쇼카드만
			else
				imgPath = ""
			end if
			%>
			<% tmptdcnt = tmptdcnt + 1 %>
			<td style='width:208px; height:133px;' valign='top'>
				<table align='left' cellpadding='0' cellspacing='0' border='0' valign='top' width='100%' height='100%'>
				<tr valign='top' align='left'>
					<td height=20>
						<% if makeriddispyn="Y" then %>
							<strong><span class="currencychar10X10" style='font-size:18px;'><%= oproduct.FItemList(i).fsocname %></span></strong>
						<% end if %>
					</td>
				</tr>
				<tr valign='top' align='left'>
					<td height=20 style="vetical-align:top; line-height:6px;">
						<% if makeriddispyn="Y" then %>
							<strong><span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).fsocname_kor %></span></strong>
						<% end if %>
					</td>			
				</tr>
				<tr valign='top' align='left'>
					<td height=93>
						<table align='left' cellpadding='0' cellspacing='0' border='0' valign='top' width='100%' height='100%'>
						<tr align='left' valign='top'>
							<td style="vetical-align:top; line-height:10px;">
								<% if isforeignprint = "Y" then %>
									<span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).Flcitemname %></span>

									<% if itemoptionyn="Y" then %>
										<% if oproduct.FItemList(i).Flcitemoptionname <> "" then %>
											<br><span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).Flcitemoptionname %></span>
										<% end if %>
									<% end if %>
								<% else %>
									<span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).Fprdname %></span>

									<% if itemoptionyn="Y" then %>
										<% if oproduct.FItemList(i).Fitemoptionname <> "" then %>
											<br><span class="currencychar10X10" style='font-size:7px;'><%= oproduct.FItemList(i).Fitemoptionname %></span>
										<% end if %>
									<% end if %>
								<% end if %>
							</td>

							<%
							'/QR코드, 이미지
							if imgPath <> "" then
							%>
								<td align='right' valign='top' width=85 rowspan=3 style='padding-right:5px'>
									<img src="<%= imgPath %>" width="<%= wd %>" height="<%= ht %>" />
								</td>
							<% end if %>
						</tr>
						<tr align='left' valign='bottom'>
							<td style="vetical-align:top; line-height:6px;">
							</td>
						</tr>
						<tr align='left' valign='top'>
							<td height=20>
								<table align='left' cellpadding='0' cellspacing='0' border='0' valign='top' width='100%' height='100%'>
								<tr align='left' valign='bottom' style='padding-bottom:5px'>
									<% if printpriceyn = "Y" or printpriceyn = "C" or printpriceyn="R" or printpriceyn="S" then %>
										<% if isforeignprint = "Y" then %>
											<td style="vetical-align:bottom; line-height:15px;">
												<strong><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= round(oproduct.FItemList(i).Flcprice,2) %></span></strong>
											</td>
										<% else %>
											<%
											'//할인가 표시
											if printpriceyn="C" then
											%>
												<%
												'/할인 처리
												if oproduct.FItemList(i).Fsaleyn="Y" and oproduct.FItemList(i).Fcustomerprice>oproduct.FItemList(i).Fsaleprice then
												%>
													<td style='text-decoration:line-through;' style="vetical-align:bottom; line-height:15px;">
														<strong><span class="currencychardefault" style='font-size:8px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:8px;'><%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %></span></strong>
													</td>
													<td style="vetical-align:bottom; line-height:15px;">
														<strong><span class="currencychardefault" style='font-size:15px; color:red;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px; color:red;'><%= FormatNumber(oproduct.FItemList(i).Fsaleprice,0) %></span></strong>
													</td>
												<%
												'/쿠폰 처리
												elseif oproduct.FItemList(i).FitemCouponYn="Y" and oproduct.FItemList(i).Fcustomerprice>oproduct.FItemList(i).GetCouponAssignPrice then
												%>
													<td style='text-decoration:line-through;' style="vetical-align:bottom; line-height:15px;">
														<strong><span class="currencychardefault" style='font-size:8px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:8px;'><%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %></span></strong>
													</td>
													<td style="vetical-align:bottom; line-height:15px;">
														<strong><span class="currencychardefault" style='font-size:15px; color:red;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px; color:red;'><%= FormatNumber(oproduct.FItemList(i).GetCouponAssignPrice(),0) %></span></strong>
													</td>
												<% else %>
													<td style="vetical-align:bottom; line-height:15px;">
														<strong><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %></span></strong>
													</td>
												<% end if %>
											<%
											'/판매가 표시
											elseif printpriceyn="R" then
											%>
												<td style="vetical-align:bottom; line-height:15px;">
													<strong><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).Fsellprice,0) %></span></strong>
												</td>
											<%
											'/심플금액표시
											elseif printpriceyn="S" then
											%>
												<td style="vetical-align:bottom; line-height:15px;">
													<strong><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %></span></strong>
												</td>
											<%
											'//소비자가 표시
											else
											%>
												<td style="vetical-align:bottom; line-height:15px;">
													<strong><span class="currencychardefault" style='font-size:15px;'><%= currencychar %></span><span class="currencychar10X10" style='font-size:15px;'><%= FormatNumber(oproduct.FItemList(i).Fcustomerprice,0) %></span></strong>
												</td>
											<% end if %>
										<% end if %>
									<% end if %>
								</tr>
								</table>
							</td>
						</tr>
						</table>
					</td>			
				</tr>
				</table>
			</td>
			<%
			'/ 세로칸 중간에 공백 줌
			if tmptdcnt=1 or tmptdcnt=2 then
				response.write "<td style='width:19px; height:133px;'>&nbsp;</td>"
			end if
			%>
			<%
			'/ 3개 넘으면 줄내림
			if tmptdcnt >= 3 then
				tmptrcnt = tmptrcnt + 1

				if (oproduct.FresultCount/3) <> tmptrcnt then
					response.write "</tr>" & vbcrlf
					response.write "<tr bgcolor='#FFFFFF'><td colspan=5 style='height:15px;'>&nbsp;</td></tr>" & vbcrlf
					response.write "<tr bgcolor='#FFFFFF'>" & vbcrlf
				end if

				tmptdcnt = 0
			end if
			%>
			<% next %>
		</tr>

	<% end if %>

	<tr bgcolor="FFFFFF">
		<td colspan="10" align="center">
	       	<% if oproduct.HasPreScroll then %>
				<font size=1><a href="javascript:reg(<%=oproduct.StartScrollPage-1%>)">[pre]</a></font>
			<% else %>
				<font size=1>[pre]</font>
			<% end if %>
			<% for i = 0 + oproduct.StartScrollPage to oproduct.StartScrollPage + oproduct.FScrollCount - 1 %>
				<% if (i > oproduct.FTotalpage) then Exit for %>
				<% if CStr(i) = CStr(oproduct.FCurrPage) then %>
					<font color="red" size=1><b><%= i %></b></font>
				<% else %>
					<a href="javascript:reg(<%=i%>);" class="list_link"><font color="#000000" size=1><%= i %></font></a>
				<% end if %>
			<% next %>
			<% if oproduct.HasNextScroll then %>
				<font size=1><a href="javascript:reg(<%=i%>);">[next]</a></font>
			<% else %>
				<font size=1>[next]</font>
			<% end if %>
		</td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td align="center">
			<% if not(isdispsql) then %>
				<font color="red"><strong>검색 조건(브랜드,상품명,물류코드,상품코드,범용바코드)을 입력 하셔야 검색이 됩니다.</strong></font>
			<% else %>
				[검색결과가 없습니다.]
			<% end if %>
		</td>
	</tr>
<% end if %>

</table>

<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="onoffgubun" value="<%= onoffgubun %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="1">
<input type="hidden" name="itembarcodearr" value="<%= itembarcodearr %>">
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="itemname" value="<%= itemname %>">
<input type="hidden" name="prdcode" value="<%= prdcode %>">
<input type="hidden" name="itemid" value="<%=replace(itemid,",",chr(10))%>">
<input type="hidden" name="generalbarcode" value="<%= generalbarcode %>">
<input type="hidden" name="sellyn" value="<%= sellyn %>">
<input type="hidden" name="usingyn" value="<%= usingyn %>">
<input type="hidden" name="isforeignprint" value="<%= isforeignprint %>">
<input type="hidden" name="printpriceyn" value="<%= printpriceyn %>">
<input type="hidden" name="makeriddispyn" value="<%= makeriddispyn %>">
<input type="hidden" name="barcodetype" value="<%= barcodetype %>">
<input type="hidden" name="listgubun" value="<%= listgubun %>">
<input type="hidden" name="papername" value="<%= papername %>">
</form>

<%
set oproduct = nothing

Function AddSpace(byval str)
	if ((str = "") or (IsNull(str))) then
		AddSpace = "&nbsp;"
	else
		AddSpace = str
	end if
End Function
%>