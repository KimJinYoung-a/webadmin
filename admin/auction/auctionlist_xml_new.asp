<%@ CODEPAGE = 65001 %>
<% option explicit %>
<% Response.CharSet = "utf-8" %> 
<%	'xml 출력시작
Response.ContentType = "text/xml"
Response.CacheControl = "public"
Response.AddHeader "Content-Type", "text/xml; charset=utf-8"
Response.AddHeader "Content-Disposition", "attachment;filename="+"Excel_Form"+".xls"
%>
<%
'###########################################################
' Description :  옥션 솔루션대로 엑셀 파일로 출력 페이지
' History : 2007.09.12 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/classes/auction/auctionclass.asp"-->
<%	
dim idx,idxsum ,itemid, rectitemid , a
	idxsum = request("idx")
	idx = left(idxsum,len(idxsum)-1)
	
'// 인덱스를 넣고 상품코드를 반환 시킨다.
dim oitemid_output , i 
	set oitemid_output = new Cauctionlist
	oitemid_output.Frectidx = idx
	oitemid_output.fitemid_output()

dim oip_image , oip3 , ten_auction_option_rect , ten_auction_option_rect2 , ten_auction_cnt_rect , ten_auction_cntsum
dim ten_auction_option , ten_auction_cnt 
Function regTest(sText)
	dim oReg
	
	set oReg= New RegExp
	
	oReg.Pattern  = "<[^>]*>"
	oReg.IgnoreCase = false
	oReg.Global = True
	regTest = oReg.Replace(sText,"")
	Set oReg = Nothing

End Function 
	
'// 받아온 상품코드를 전부 배열에 넣는다.
for i = 0 to oitemid_output.ftotalcount -1
	rectitemid = rectitemid&oitemid_output.flist(i).ten_itemid&","
next
itemid = left(rectitemid, len(rectitemid)-1)

dim oip
	set oip = new Cauctionlist        			'클래스 지정 
	oip.frectitemid = itemid
	oip.fauction_excel()

dim Imginfo,t, oADD , idxvar
idxvar = 1
%><?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author>nyhong</Author>
  <LastAuthor>10x10</LastAuthor>
  <Created>2007-12-07T02:47:39Z</Created>
  <LastSaved>2008-03-04T04:37:26Z</LastSaved>
  <Company>auction</Company>
  <Version>12.00</Version>
 </DocumentProperties>
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>11640</WindowHeight>
  <WindowWidth>16545</WindowWidth>
  <WindowTopX>240</WindowTopX>
  <WindowTopY>90</WindowTopY>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
 </ExcelWorkbook>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Center"/>
   <Borders/>
   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern" ss:Size="11"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID="s63" ss:Name="하이퍼링크">
   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern" ss:Size="11"
    ss:Color="#0000FF" ss:Underline="Single"/>
  </Style>
  <Style ss:ID="s64">
   <Borders/>
   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s65">
   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s66">
   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="s67">
   <Borders/>
   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="s68">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s69">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Borders/>
   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="s70">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>
   <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="s71">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Borders/>
   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>
   <Interior ss:Color="#FFFF99" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="s72">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Borders/>
   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="s73">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Font ss:FontName="돋움" x:CharSet="129" x:Family="Modern"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
  </Style>
  <Style ss:ID="s74" ss:Parent="s63">
   <Alignment ss:Vertical="Center"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
   <NumberFormat ss:Format="@"/>
   <Protection/>
  </Style>
 </Styles>
 <Worksheet ss:Name="일괄등록 양식">
  <Table ss:ExpandedColumnCount="64" ss:ExpandedRowCount="163" x:FullColumns="1"
   x:FullRows="1" ss:StyleID="s65" ss:DefaultColumnWidth="72.75"
   ss:DefaultRowHeight="12">
   <Column ss:Index="43" ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="76.5"/>
   <Column ss:Index="45" ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="123.75"/>
   <Column ss:Index="59" ss:StyleID="s64" ss:AutoFitWidth="0" ss:Span="5"/>
   <Row ss:StyleID="s73">
    <Cell ss:StyleID="s70"><Data ss:Type="String">SequenceNo</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">CategoryCode</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">CatalogCode</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">ProductTypeCode</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">ItemName</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">AdMessage</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">SellerProductCode</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">BrandCode</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">BrandNameCustom</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">PlaceOfOriginCode</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">NationCode</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">Importer</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">ProductionDate</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">UseByDate</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">MarketPrice</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">UsedMonth</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">ListImage</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">Image1</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">Image2</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">Image3</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">Description1</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">Description2</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">Description3</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">Description</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">EnableAS</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">ASInfo</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">MedicalInstrumentRegistrationNo</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">MedicalRegistrationNo</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">MedicalRegistrationOfficeName</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">HealthFoodRegistrationNo</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">HealthFoodRegistrationOfficeName</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">FoodProcessingRegistrationNo</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">FoodProcessingRegistrationOfficeName</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">DeliberationNumber</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">LimitedAge</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">SellingPrice</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">Subsidy</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">MonthAmount</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">InstallmentCount</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">PrincipalAmount</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">BuyLimitQty</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">SellingRegionCode</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">ShippingTypeCode</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">ShippingCostChargeCode</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">ShippingCostDiscountCode</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">IsPrepayable</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">EnableBundleShipping</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">ConditionValue</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">ShippingCost</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">ReturnPostNo</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">ReturnAddressPostNo</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">ReturnAddressDetail</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">DeliveryAgencyCode</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">SupportNoInterest</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">SupportThreeMonth</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">SupportSixMonth</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">LumbSumDiscount</Data></Cell>
    <Cell ss:StyleID="s70"><Data ss:Type="String">IsPCSRegistration</Data></Cell>
    <Cell ss:StyleID="s71"><Data ss:Type="String">OrderTypeCode</Data></Cell>
    <Cell ss:StyleID="s71"><Data ss:Type="String">BuyerDescriptiveText</Data></Cell>
    <Cell ss:StyleID="s71"><Data ss:Type="String">Stock</Data></Cell>
    <Cell ss:StyleID="s71"><Data ss:Type="String">OptionTypeCode</Data></Cell>
    <Cell ss:StyleID="s71"><Data ss:Type="String">Option</Data></Cell>
    <Cell ss:StyleID="s72"/>
   </Row>
<%
for i = 0 to oip.ftotalcount -1

	set oip3 = new Cauctionlist        	'클래스 지정
		oip3.Frectitemid = oip.flist(i).ten_itemid
		oip3.fwritelist_auction()
	
	ten_auction_option_rect=""
	ten_auction_option_rect2 = ""	 
	ten_auction_cnt_rect = ""
	ten_auction_cntsum = 0
	
		for a=0 to oip3.FTotalCount - 1	
			ten_auction_cntsum = cint(oip3.flist(a).ten_jaego) + ten_auction_cntsum
				if a <> 0 then 
					ten_auction_option_rect = ",^"
				else 
					ten_auction_option_rect = ""	
				end if
				if a = oip3.FTotalCount-1 then 
					ten_auction_option_rect2 = ","
				else 
					ten_auction_option_rect2 = ""	
				end if
				
			ten_auction_option = ten_auction_option + ten_auction_option_rect + cstr(oip3.flist(a).ten_itemid) + "," + cstr(oip3.flist(a).ten_option) + "," + cstr(oip3.flist(a).ten_jaego) + ten_auction_option_rect2				
			ten_auction_cnt_rect = ten_auction_cnt_rect+cstr(oip3.flist(a).ten_jaego)+"," 			                  
			ten_auction_cnt = left(ten_auction_cnt_rect , len(ten_auction_cnt_rect)-1)
			ten_auction_option_rect=""
		next	
	
	'//상세설명이미지 출력
	set oip_image = new Cauctionlist
	oip_image.frectitemid = oip.flist(i).ten_itemid
	oip_image.fauction_excel_infoimage() 
%>
   <Row ss:Height="13.5" ss:StyleID="s66">
    <Cell><Data ss:Type="String"><%= idxvar %><% idxvar = idxvar + 1 %></Data></Cell>
    <Cell><Data ss:Type="String"><%= oip.flist(i).auction_cate_code %></Data></Cell>
    <Cell ss:Index="4"><Data ss:Type="String">0</Data></Cell>
	<Cell><Data ss:Type="String"><%= replace(replace(replace(oip.flist(i).ten_itemname,"'",""),"<","["),">","]") %></Data></Cell>
    <Cell><Data ss:Type="String"></Data></Cell>
    <Cell ss:Index="9"><Data ss:Type="String">텐바이텐</Data></Cell>
    <Cell><Data ss:Type="String">1</Data></Cell>
    <Cell ss:Index="17" ss:StyleID="s72" ss:HRef="<%= oip.flist(i).FImagebasic %>"><Data
      ss:Type="String"><%= oip.flist(i).FImagebasic %></Data></Cell>
    <Cell ss:StyleID="s72" ss:HRef="<%= oip.flist(i).FImagebasic %>"><Data
      ss:Type="String"><%= oip.flist(i).FImagebasic %></Data></Cell>
    <Cell ss:StyleID="s72" ss:HRef="<% set oADD = new CAutoCategory %><% oADD.getAddImage(oip.flist(i).ten_itemid) %><% if (oAdd.FResultCount>0) then %><% if oAdd.FADD(0).FAddimage <> "" then %><%= oAdd.FADD(0).FAddimage %><% end if %><% end if %>"><Data
      ss:Type="String"><% set oADD = new CAutoCategory %><% oADD.getAddImage(oip.flist(i).ten_itemid) %><% if (oAdd.FResultCount>0) then %><% if oAdd.FADD(0).FAddimage <> "" then %><%= oAdd.FADD(0).FAddimage %><% end if %><% end if %></Data></Cell>
    <Cell ss:StyleID="s72" ss:HRef="<% if (oAdd.FResultCount>2) then %><% if oAdd.FADD(1).FAddimage <> "" then %><%= oAdd.FADD(1).FAddimage %><% end if %><% end if %>"><Data
      ss:Type="String"><% if (oAdd.FResultCount>2) then %><% if oAdd.FADD(1).FAddimage <> "" then %><%= oAdd.FADD(1).FAddimage %><% end if %><% end if %></Data></Cell>    
    <Cell ss:Index="24" ss:StyleID="s65"><Data ss:Type="String"><![CDATA[<%= oip.flist(i).ten_itemcontent %><% response.write "<br>" %><% if oip_image.FTotalCount > 0 then %><% for t=0 to oip_image.FTotalCount -1 %><img src="<%= oip_image.flist(t).FImageInfoStr %>" board=0><% response.write "<br>" %><% next %><% end if %>]]></Data></Cell>
    <Cell><Data ss:Type="String">false</Data></Cell>
    <Cell ss:Index="35"><Data ss:Type="String"></Data></Cell>
    <Cell><Data ss:Type="String"><%= oip.flist(i).fsellcash %></Data></Cell>
    <Cell ss:Index="41"><Data ss:Type="String">10</Data></Cell>
    <Cell ss:StyleID="Default"><Data ss:Type="String">01Z0</Data></Cell>
    <Cell><Data ss:Type="String">1</Data></Cell>
    <Cell><Data ss:Type="String">5</Data></Cell>
    <Cell><Data ss:Type="String">2</Data></Cell>
    <Cell><Data ss:Type="String">true</Data></Cell>
    <Cell><Data ss:Type="String">true</Data></Cell>

		<Cell ss:Index="50"><Data ss:Type="String">11154</Data></Cell>
		<Cell><Data ss:Type="String">경기도 포천시 군내면 용정경제로2길 83 </Data></Cell>
		<Cell><Data ss:Type="String">텐바이텐 물류센터</Data></Cell>

    <Cell ss:Index="59" ss:StyleID="s67"><Data ss:Type="String"><% if oip.flist(i).foptioncnt = "0" then %><% response.write "0" %><% else %><% response.write "1" %><% end if %></Data></Cell>
    <Cell ss:StyleID="s67"/>
    <Cell ss:StyleID="s67"><Data ss:Type="String"><% if oip.flist(i).foptioncnt = "0" then %><%= ten_auction_cntsum %><% else %><%= ten_auction_option %><% end if %></Data></Cell>
    <Cell ss:StyleID="s67"><Data ss:Type="String">0</Data></Cell>
    <Cell ss:StyleID="s67"/>
   </Row>
<%   
	ten_auction_option_rect="" 
	ten_auction_cnt_rect = ""
	ten_auction_cntsum = 0
	ten_auction_option = ""
next 
%>
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <Print>
    <ValidPrinterInfo/>
    <PaperSizeIndex>9</PaperSizeIndex>
    <HorizontalResolution>600</HorizontalResolution>
    <VerticalResolution>600</VerticalResolution>
   </Print>
   <Selected/>
   <LeftColumnVisible>30</LeftColumnVisible>
   <FreezePanes/>
   <FrozenNoSplit/>
   <SplitHorizontal>1</SplitHorizontal>
   <TopRowBottomPane>1</TopRowBottomPane>
   <ActivePane>2</ActivePane>
   <Panes>
    <Pane>
     <Number>3</Number>
    </Pane>
    <Pane>
     <Number>2</Number>
     <ActiveRow>12</ActiveRow>
     <ActiveCol>35</ActiveCol>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
 <Worksheet ss:Name="필드설명">
  <Table ss:ExpandedColumnCount="4" ss:ExpandedRowCount="73" x:FullColumns="1"
   x:FullRows="1" ss:StyleID="s65" ss:DefaultColumnWidth="60"
   ss:DefaultRowHeight="12">
   <Column ss:StyleID="s68" ss:Width="23.25"/>
   <Column ss:StyleID="s68" ss:Width="54"/>
   <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="140.25"/>
   <Column ss:StyleID="s65" ss:AutoFitWidth="0" ss:Width="207.75"/>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="String">No.</Data></Cell>
    <Cell><Data ss:Type="String">구분</Data></Cell>
    <Cell><Data ss:Type="String">필드명</Data></Cell>
    <Cell><Data ss:Type="String">필드설명</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">1</Data></Cell>
    <Cell><Data ss:Type="String">필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">SequenceNo</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">일련번호</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">2</Data></Cell>
    <Cell><Data ss:Type="String">필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">CategoryCode</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">카테고리 코드</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">3</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">CatalogCode</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">카탈로그 코드</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">4</Data></Cell>
    <Cell><Data ss:Type="String">필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">ProductTypeCode</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">물품상태 코드 0:new(새물품), 1:used(중고)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">5</Data></Cell>
    <Cell><Data ss:Type="String">필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">ItemName</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">물품명</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">6</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">AdMessage</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">홍보문구</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">7</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">SellerProductCode</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">판매자 관리코드(varchar 30)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">8</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">BrandCode</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">브랜드 코드</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">9</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">BrandNameCustom</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">브랜드 직접 입력</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">10</Data></Cell>
    <Cell><Data ss:Type="String">필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">PlaceOfOriginCode</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">원산지 정보 (0:모름, 1:국내, 2:국외)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">11</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">NationCode</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">수입국가 코드</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">12</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">Importer</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">수입원 정보</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">13</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">ProductionDate</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">제조일자&amp;발행일자(예 : 2007-08-03)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">14</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">UseByDate</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">유효일자(예 : 2007-08-03)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">15</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">MarketPrice</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">시중가</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">16</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">UsedMonth</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">사용개월 수</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">17</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">ListImage</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">목록이미지 URL</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">18</Data></Cell>
    <Cell><Data ss:Type="String">필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">Image1</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">첫번째 이미지</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">19</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">Image2</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">두번째 이미지</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">20</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">Image3</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">세번째 이미지</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">21</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">Description1</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">첫번째 이미지 설명</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">22</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">Description2</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">두번째 이미지 설명</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">23</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">Description3</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">세번째 이미지 설명</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">24</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">Description</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">물품 상세 설명(Text 또는 Html로 작성)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">25</Data></Cell>
    <Cell><Data ss:Type="String">필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">EnableAS</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">AS가능 여부(true, false)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">26</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">ASInfo</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">AS정보</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">27</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">MedicalInstrumentRegistrationNo</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">의료기기 품목 허가번호</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">28</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">MedicalRegistrationNo</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">의료기기 판매업 신고번호</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">29</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">MedicalRegistrationOfficeName</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">의료기기 판매업 신고기관명</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">30</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">HealthFoodRegistrationNo</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">건강기능 식품 판매업 신고번호</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">31</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">HealthFoodRegistrationOfficeName</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">건강기능 식품 판매업 신고기관명</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">32</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">FoodProcessingRegistrationNo</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">식품제조 가공업 신고번호</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">33</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">FoodProcessingRegistrationOfficeName</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">식품제조 가공업 신고기관명</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">34</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">DeliberationNumber</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">심의번호</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">35</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">LimitedAge</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">이용등급 (18, 19)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">36</Data></Cell>
    <Cell><Data ss:Type="String">필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">SellingPrice</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">판매가</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">37</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">Subsidy</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">핸드폰 보조금</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">38</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">MonthAmount</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">핸드폰 할부 월납입액</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">39</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">InstallmentCount</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">핸드폰 할부 개월 수</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">40</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">PrincipalAmount</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">핸드폰 할부 원금</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">41</Data></Cell>
    <Cell><Data ss:Type="String">필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">BuyLimitQty</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">최대구매허용수량</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">42</Data></Cell>
    <Cell><Data ss:Type="String">필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">SellingRegionCode</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">판매가능지역 코드 좌측에서 코드만 입력  (01A0:서울, 01B0:인천, 01C0:광주, 01D0:대구, 01E0:대전, 01F0:부산, 01G0:울산, 01H0:경기, 01I0:강원, 01J0:충남, 01K0:충북, 01L0:경남, 01M0:경북, 01N0:전남, 01O0:전북, 01P0:제주, 01Q0:서울/경기, 01R0:서울/경기/대전, 01S0:충북/충남, 01T0:경북/경남, 01U0:전북/전남, 01V0:부산/울산, 01Z0:전국, 01Z1:전국(제주,도서지역 제외) )</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">43</Data></Cell>
    <Cell><Data ss:Type="String">필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">ShippingTypeCode</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">배송구분 코드(0:Unavailable(배송없음),1:Door2Door(택배),2:Parcel(우편,소포,등기),3:QuickService(퀵서비스),4:Direct(직접배송),7:Phone)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">44</Data></Cell>
    <Cell><Data ss:Type="String">필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">ShippingCostChargeCode</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">배송비 부담 코드(1:Free,2:PayOnArrival,3:SingleConditional,4:MultiConditional,5:SellerConditional)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">45</Data></Cell>
    <Cell><Data ss:Type="String">필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">ShippingCostDiscountCode</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">배송비 할인 구분(1:ProductAmount,2:ProductQuantity) * ShippingCostChargeCode가 4인 경우만 1,2선택이 가능하고, 나머지 경우에는 모두 2여야 입력가능</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">46</Data></Cell>
    <Cell><Data ss:Type="String">필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">IsPrepayable</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">선결제 여부(true, false)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">47</Data></Cell>
    <Cell><Data ss:Type="String">필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">EnableBundleShipping</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">묶음배송 여부(true, false)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">48</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">ConditionValue</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">배송조건 (금액/수량) *수량별 차등 일 경우 배송 조건을 &quot;,&quot;로 구분하여 최대 4개까지 입력(예: 금액인 경우 : 1,10000,20000 -&gt; 무조건 1부터 시작하며 구간 적용됨 &amp; 수량인 경우 : 1,2,3,4)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">49</Data></Cell>
    <Cell><Data ss:Type="String">필수</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">ShippingCost</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">배송비 *수량별 차등 일 경우 배송비를 수량별 차등 배송 조건수와 동일하게 최대 4개까지 입력(예: 5000,3000,2500,2000)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">50</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">ReturnPostNo</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">반품지 우편번호 (예:137070)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">51</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">ReturnAddressPostNo</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">반품지 주소 (예:서울시 서초구 서초동)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">52</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">ReturnAddressDetail</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">반품지 주소 상세 (예:교보타워 14층)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">53</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">DeliveryAgencyCode</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">반품 택배사 코드 (좌측에서 코드만 입력한다.  코드:택배사명 &gt; 1:대한통운택배 5:현대택배 6:우체국택배 A:동부익스프레스택배 B:삼성HTH택배 C:아주택배 D:CJ GLS택배 E:로젠택배 G:옐로우캡 H:한진택배 I:이젠택배 K:KGB택배 L:(주)벨익스프레스 N:사가와익스프레스 O:KT로지스 P:세덱스 Q:하나로로지스)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">54</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">SupportNoInterest</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">판매자 무이자 할부 여부(true, false)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">55</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">SupportThreeMonth</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">판매자 무이자 3개월 할부 여부(true, false)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">56</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">SupportSixMonth</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">판매자 무이자 6개월 할부 여부(true, false)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">57</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">LumbSumDiscount</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">판매자 무이자 일시불 할인 여부(true, false)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">58</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s66"><Data ss:Type="String">IsPCSRegistration</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">가격비교 사이트 등록 여부(true, false)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">59</Data></Cell>
    <Cell><Data ss:Type="String">필수</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">OrderTypeCode</Data></Cell>
    <Cell><Data ss:Type="String">주문선택사항(0:사용안함, 1:구매자선택형, 3:구매자작성형)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">60</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">BuyerDescriptiveText</Data></Cell>
    <Cell><Data ss:Type="String">구매자작성형 예시문구</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">61</Data></Cell>
    <Cell><Data ss:Type="String">선택&amp;필수</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">Stock</Data></Cell>
    <Cell><Data ss:Type="String">주문선택사항 미사용시 재고수량 입력(ex 500) &amp; 주문선택사항 사용시 주문선택사항 선택형 상세내용입력(구분1-1명칭,구분2-1명칭,재고수량,판매자상세관리코드^구분2-1명칭,구분2-2명칭,재고수량,판매자상세관리코드^구분3-1명칭,구분3-2명칭,재고수량,판매자상세관리코드^…)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">62</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">OptionTypeCode</Data></Cell>
    <Cell><Data ss:Type="String">옵션구매(0:사용안함, 1:사용함)</Data></Cell>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell><Data ss:Type="Number">63</Data></Cell>
    <Cell><Data ss:Type="String">선택</Data></Cell>
    <Cell ss:StyleID="s64"><Data ss:Type="String">Option</Data></Cell>
    <Cell><Data ss:Type="String">옵션구매 상세내용(옵션구분1, 상세물품명1,재고수량^옵션구분2, 상세물품명2,재고수량^옵션구분3, 상세물품명3,재고수량^…)</Data></Cell>
   </Row>
   <Row ss:Index="69" ss:AutoFitHeight="0">
    <Cell ss:Index="3" ss:StyleID="s64"/>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="3" ss:StyleID="s64"/>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="3" ss:StyleID="s64"/>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="3" ss:StyleID="s64"/>
   </Row>
   <Row ss:AutoFitHeight="0">
    <Cell ss:Index="3" ss:StyleID="s64"/>
   </Row>
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <Unsynced/>
   <TopRowVisible>24</TopRowVisible>
   <LeftColumnVisible>2</LeftColumnVisible>
   <Panes>
    <Pane>
     <Number>3</Number>
     <ActiveRow>36</ActiveRow>
     <ActiveCol>2</ActiveCol>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
</Workbook>
<!-- #include virtual="/lib/db/dbclose.asp" -->