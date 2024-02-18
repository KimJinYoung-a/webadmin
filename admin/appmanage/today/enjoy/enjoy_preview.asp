<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : enjoy_insert.asp
' Discription : 모바일 enjoybanner_new
' History : 2014.06.23 이종화
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/mobile/todayenjoyCls.asp" -->
<%
'###############################################
'이벤트 신규 등록시
'###############################################
%>
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<%
	Dim idx, sqlStr, linkurl, evttitle, evtimg, evtalt, evttitle2, enddate, startdate, evt_todaybanner, evt_mo_listbanner, linktype, evtstdate, evteddate, tag_gift, tag_plusone
	Dim tag_launching, tag_actively, sale_per, coupon_per, itemid1, itemid2, itemid3, addtype, iteminfoValPreview, vTrendImg, evttag
	Dim sellCash, orgPrice, sailYN, couponYN, couponvalue, coupontype, itemimg
	Dim itemimg1, price1, sale1, itemimg2, price2, sale2, itemimg3, price3, sale3, ii, coupon_flag
	Dim itemid1url, itemid2url, itemid3url

	idx = request("idx")

	sqlStr = " SELECT TOP 1 t.idx, t.linkurl, t.evttitle, t.evtimg, t.evtalt, t.evttitle2, t.enddate , t.startdate , d.evt_todaybanner "
	sqlStr = sqlStr & " , d.evt_mo_listbanner , t.linktype , t.evtstdate , t.evteddate , t.tag_gift , t.tag_plusone , t.tag_launching , t.tag_actively , t.sale_per "
	sqlStr = sqlStr & " ,  t.coupon_per , t.itemid1 , t.itemid2 , t.itemid3 , t.addtype  "
	sqlStr = sqlStr & " , STUFF((  "
	sqlStr = sqlStr & "  			SELECT ',' + cast(i.itemid as varchar(120)) +'|'+ cast(i.sellCash as varchar(120)) +'|'+ cast(i.orgPrice as varchar(50))  "
	sqlStr = sqlStr & "  			+'|'+ cast(i.sailyn as varchar(50))+'|'+ cast(i.itemcouponYn as varchar(50))+'|'+ cast(i.itemcouponvalue as varchar(50))+'|'+ cast(i.limitYN as varchar(50)) "
	sqlStr = sqlStr & "  			+'|'+ cast(i.itemcoupontype as varchar(50))+'|'+ cast(i.icon1image as varchar(50)) "
	sqlStr = sqlStr & "  			FROM db_item.dbo.tbl_item as i  "
	sqlStr = sqlStr & "  			WHERE i.itemid in (t.itemid1 , t.itemid2 , t.itemid3 ) and i.itemid<>0  "
	sqlStr = sqlStr & "  			FOR XML PATH('')  "
	sqlStr = sqlStr & "  			), 1, 1, '') AS iteminfo "
	sqlStr = sqlStr & " FROM db_sitemaster.dbo.tbl_mobile_main_enjoyevent_new as t  "
	sqlStr = sqlStr & " LEFT JOIN db_event.dbo.tbl_event_display as d on t.evt_code = d.evt_code  "
	sqlStr = sqlStr & " WHERE  "
	sqlStr = sqlStr & " t.idx='"&idx&"' "
    rsget.Open sqlStr, dbget, 1
	If Not(rsget.bof Or rsget.eof) Then
		linkurl = rsget("linkurl")
		evttitle = rsget("evttitle")
		evtimg = rsget("evtimg")
		evtalt = rsget("evtalt")
		evttitle2 = rsget("evttitle2")
		enddate = rsget("enddate")
		startdate = rsget("startdate")
		evt_todaybanner = rsget("evt_todaybanner")
		evt_mo_listbanner = rsget("evt_mo_listbanner")
		linktype = rsget("linktype")
		evtstdate = rsget("evtstdate")
		evteddate = rsget("evteddate")
		tag_gift = rsget("tag_gift")
		tag_plusone = rsget("tag_plusone")
		tag_launching = rsget("tag_launching")
		tag_actively = rsget("tag_actively")
		sale_per = rsget("sale_per")
		coupon_per = rsget("coupon_per")
		itemid1 = rsget("itemid1")
		itemid2 = rsget("itemid2")
		itemid3 = rsget("itemid3")
		addtype = rsget("addtype")
		iteminfoValPreview = rsget("iteminfo")
	Else
		response.write "<script>alert('정상적인 경로로 접근해주세요.');window.close();</script>"
		response.End
	End If
	rsget.close

	If linktype=1 Then
		If application("Svr_Info") = "Dev" Then
			vTrendImg = evt_mo_listbanner
		Else
			vTrendImg = getThumbImgFromURL(evt_mo_listbanner,750,"","","")
		End If 
	Else
		vTrendImg = staticImgUrl & "/mobile/enjoyevent" & evtimg
	End If

	coupon_flag		= chkiif(coupon_per<>"","1","0")
	itemid1url = "/category/category_itemPrd.asp?itemid="& itemid1
	itemid2url = "/category/category_itemPrd.asp?itemid="& itemid2
	itemid3url = "/category/category_itemPrd.asp?itemid="& itemid3

	If tag_actively = "Y" Then evttag = "참여"	'//actively
	If tag_launching = "Y" Then evttag = "런칭"	'//launching
	If tag_plusone = "Y" Then evttag = "1+1"	'//plusone
	If tag_gift = "Y" Then evttag = "GIFT"	'//gift
	If coupon_per <> "" Then evttag = "쿠폰"&coupon_per

	If addtype = 2 Then '//가격 정보 addtype 기본형 + 상품 3개 일때
		If iteminfoValPreview <> "" Or Not isnull(iteminfoValPreview) Then 
			If ubound(Split(iteminfoValPreview,",")) > 0 Then ' 이미지 3개 정보
				For ii = 0 To ubound(Split(iteminfoValPreview,","))

					sellCash	= Split(Split(iteminfoValPreview,",")(ii),"|")(1)
					orgPrice	= Split(Split(iteminfoValPreview,",")(ii),"|")(2)
					sailYN		= Split(Split(iteminfoValPreview,",")(ii),"|")(3)
					couponYn	= Split(Split(iteminfoValPreview,",")(ii),"|")(4)
					couponvalue = Split(Split(iteminfoValPreview,",")(ii),"|")(5)
					coupontype	= Split(Split(iteminfoValPreview,",")(ii),"|")(7)
					itemimg		= Split(Split(iteminfoValPreview,",")(ii),"|")(8)

					'//1번 상품	
					If CStr(itemid1) = CStr(Split(Split(iteminfoValPreview,",")(ii),"|")(0)) Then
						itemimg1 = webImgUrl & "/image/icon1/" & GetImageSubFolderByItemid(itemid1) & "/" & itemimg

						If sailYN = "N" and couponYn = "N" Then
							price1 = ""&formatNumber(orgPrice,0) &""
						End If
						If sailYN = "Y" and couponYn = "N" Then
							price1 = ""&formatNumber(sellCash,0) &""
						End If
						if couponYn = "Y" And couponvalue>0 Then
							If coupontype = "1" Then
								price1 = ""&formatNumber(sellCash - CLng(couponvalue*sellCash/100),0) &""
							ElseIf coupontype = "2" Then
								price1 = ""&formatNumber(sellCash - couponvalue,0) &""
							ElseIf coupontype = "3" Then
								price1 = ""&formatNumber(sellCash,0) &""
							Else
								price1 = ""&formatNumber(sellCash,0) &""
							End If
						End If
						If sailYN = "Y" And couponYn = "Y" Then
							If coupontype = "1" Then
								'//할인 + %쿠폰
								sale1 = ""& CLng((orgPrice-(sellCash - CLng(couponvalue*sellCash/100)))/orgPrice*100)&"%"
							ElseIf coupontype = "2" Then
								'//할인 + 원쿠폰
								sale1 = ""& CLng((orgPrice-(sellCash - couponvalue))/orgPrice*100)&"%"
							Else
								'//할인 + 무배쿠폰
								sale1 = ""& CLng((orgPrice-sellCash)/orgPrice*100)&"%"
							End If 
						ElseIf sailYN = "Y" and couponYn = "N" Then
							If CLng((orgPrice-sellCash)/orgPrice*100)> 0 Then
								sale1 = ""& CLng((orgPrice-sellCash)/orgPrice*100)&"%"
							End If
						elseif sailYN = "N" And couponYn = "Y" And couponvalue>0 Then
							If coupontype = "1" Then
								sale1 = ""&  CStr(couponvalue) & "%"
							ElseIf coupontype = "2" Then
								sale1 = ""
							ElseIf coupontype = "3" Then
								sale1 = ""
							Else
								sale1 = ""& couponvalue &"%"
							End If
						Else 
							sale1 = ""
						End If
					End If
					
					'//2번 상품
					If CStr(itemid2) = CStr(Split(Split(iteminfoValPreview,",")(ii),"|")(0)) Then
						itemimg2 =  webImgUrl & "/image/icon1/" & GetImageSubFolderByItemid(itemid2) & "/" & itemimg

						If sailYN = "N" and couponYn = "N" Then
							price2 = ""&formatNumber(orgPrice,0) &""
						End If
						If sailYN = "Y" and couponYn = "N" Then
							price2 = ""&formatNumber(sellCash,0) &""
						End If
						if couponYn = "Y" And couponvalue>0 Then
							If coupontype = "1" Then
							price2 = ""&formatNumber(sellCash - CLng(couponvalue*sellCash/100),0) &""
							ElseIf coupontype = "2" Then
							price2 = ""&formatNumber(sellCash - couponvalue,0) &""
							ElseIf coupontype = "3" Then
							price2 = ""&formatNumber(sellCash,0) &""
							Else
							price2 = ""&formatNumber(sellCash,0) &""
							End If
						End If
						If sailYN = "Y" And couponYn = "Y" Then
							If coupontype = "1" Then
								'//할인 + %쿠폰
								sale2 = ""& CLng((orgPrice-(sellCash - CLng(couponvalue*sellCash/100)))/orgPrice*100)&"%"
							ElseIf coupontype = "2" Then
								'//할인 + 원쿠폰
								sale2 = ""& CLng((orgPrice-(sellCash - couponvalue))/orgPrice*100)&"%"
							Else
								'//할인 + 무배쿠폰
								sale2 = ""& CLng((orgPrice-sellCash)/orgPrice*100)&"%"
							End If 
						ElseIf sailYN = "Y" and couponYn = "N" Then
							If CLng((orgPrice-sellCash)/orgPrice*100)> 0 Then
								sale2 = ""& CLng((orgPrice-sellCash)/orgPrice*100)&"%"
							End If
						elseif sailYN = "N" And couponYn = "Y" And couponvalue>0 Then
							If coupontype = "1" Then
								sale2 = ""&  CStr(couponvalue) & "%"
							ElseIf coupontype = "2" Then
								sale2 = ""
							ElseIf coupontype = "3" Then
								sale2 = ""
							Else
								sale2 = ""& couponvalue &"%"
							End If
						Else 
							sale2 = ""
						End If
					End If

					'//3번 상품
					If CStr(itemid3) = CStr(Split(Split(iteminfoValPreview,",")(ii),"|")(0)) Then
						itemimg3 =  webImgUrl & "/image/icon1/" & GetImageSubFolderByItemid(itemid3) & "/" & itemimg

						If sailYN = "N" and couponYn = "N" Then
							price3 = ""&formatNumber(orgPrice,0) &""
						End If
						If sailYN = "Y" and couponYn = "N" Then
							price3 = ""&formatNumber(sellCash,0) &""
						End If
						if couponYn = "Y" And couponvalue>0 Then
							If coupontype = "1" Then
							price3 = ""&formatNumber(sellCash - CLng(couponvalue*sellCash/100),0) &""
							ElseIf coupontype = "2" Then
							price3 = ""&formatNumber(sellCash - couponvalue,0) &""
							ElseIf coupontype = "3" Then
							price3 = ""&formatNumber(sellCash,0) &""
							Else
							price3 = ""&formatNumber(sellCash,0) &""
							End If
						End If
						If sailYN = "Y" And couponYn = "Y" Then
							If coupontype = "1" Then
								'//할인 + %쿠폰
								sale3 = ""& CLng((orgPrice-(sellCash - CLng(couponvalue*sellCash/100)))/orgPrice*100)&"%"
							ElseIf coupontype = "2" Then
								'//할인 + 원쿠폰
								sale3 = ""& CLng((orgPrice-(sellCash - couponvalue))/orgPrice*100)&"%"
							Else
								'//할인 + 무배쿠폰
								sale3 = ""& CLng((orgPrice-sellCash)/orgPrice*100)&"%"
							End If 
						ElseIf sailYN = "Y" and couponYn = "N" Then
							If CLng((orgPrice-sellCash)/orgPrice*100)> 0 Then
								sale3 = ""& CLng((orgPrice-sellCash)/orgPrice*100)&"%"
							End If
						elseif sailYN = "N" And couponYn = "Y" And couponvalue>0 Then
							If coupontype = "1" Then
								sale3 = ""&  CStr(couponvalue) & "%"
							ElseIf coupontype = "2" Then
								sale3 = ""
							ElseIf coupontype = "3" Then
								sale3 = ""
							Else
								sale3 = ""& couponvalue &"%"
							End If
						Else 
							sale3 = ""
						End If
					End If
				Next 
			End If
		End If 
	End If 
%>
<style>
/************************* DEFAULT *************************/
html, body, div, span, applet, object, iframe,
h1, h2, h3, h4, h5, h6, p, blockquote, pre,
a, abbr, acronym, address, big, cite, code,
del, dfn, em, img, ins, kbd, q, s, samp,
small, strike, strong, sub, sup, tt, var,
b, u, i, center,
dl, dt, dd, ol, ul, li,
fieldset, form, label, legend,
table, caption, tbody, tfoot, thead, tr, th, td,
article, aside, canvas, details, embed,
figure, figcaption, footer, header, hgroup,
menu, nav, output, ruby, section, summary,
time, mark, audio, video {margin:0; padding:0; border:0; font:inherit; vertical-align:baseline;}

/* HTML5 display-role reset for older browsers */
article, aside, details, figcaption, figure,
footer, header, hgroup, menu, nav, section {display:block;}

body {line-height:1;}

ol, ul {list-style:none;}
blockquote, q {quotes:none;}

blockquote:before, blockquote:after,
q:before, q:after {content:''; content:none;}

table {border-collapse:collapse; border-spacing:0;}
strong {font-weight:bold;}

/* 디바이스 해상도 아니고 웹브라우저 해상도(뷰포트) 기준임 */
html {font-size:10px;} /* iphone5 */
@media (max-width:320px) {html{font-size:10px;}} /* iphone5 */
@media (min-width:360px) and (orientation:portrait) {html{font-size:11.25px;}} /* galaxy3, galaxy note3, galaxy alpha, nexus5, sony xperia, LG G2 */
@media (min-width:375px) and (orientation:portrait) {html{font-size:11.71875px;}} /* iphone6 */
@media (min-width:384px) and (orientation:portrait) {html{font-size:12px;}} /* LG Optimus G, nexus4 */
@media (min-width:412px) and (orientation:portrait) {html{font-size:12.875px;}} /* galaxy note7 */
@media (min-width:414px) and (orientation:portrait) {html{font-size:12.93px;}} /* iphone6+ */
@media (min-width:768px) {html{font-size:14px;}}

a {text-decoration:none; color:inherit;}
img {width:100%; vertical-align:top;}
* {box-sizing:border-box; -webkit-box-sizing:border-box; -moz-box-sizing:border-box; box-sizing:border-box; -webkit-text-size-adjust:none;}

/************************* SWIPER *************************/
.swiper-container {overflow:hidden; position:relative; width:100%; margin:0 auto; z-index:1;}
.swiper-container-no-flexbox .swiper-slide {float:left;}
.swiper-wrapper {position:relative; width:100%; z-index:1; display:-webkit-box; display:-moz-box; display:-ms-flexbox; display:-webkit-flex; display:flex;}
.swiper-container-android .swiper-slide,
.swiper-wrapper {-webkit-transform:translate3d(0px, 0, 0); -moz-transform:translate3d(0px, 0, 0); -o-transform:translate(0px, 0px); -ms-transform:translate3d(0px, 0, 0); transform:translate3d(0px, 0, 0);}
.swiper-container-free-mode > .swiper-wrapper {margin:0 auto; -webkit-transition-timing-function:ease-out; -moz-transition-timing-function:ease-out; -ms-transition-timing-function:ease-out; -o-transition-timing-function:ease-out; transition-timing-function:ease-out;}
.swiper-container-vertical > .swiper-wrapper {-webkit-box-orient:vertical; -moz-box-orient:vertical; -ms-flex-direction:column; -webkit-flex-direction:column; flex-direction:column;}
.swiper-slide {position:relative; -webkit-flex-shrink:0; -ms-flex:0 0 auto; flex-shrink:0;}
/* IE10 Windows Phone 8 Fixes */
.swiper-wp8-horizontal {-ms-touch-action:pan-y; touch-action:pan-y;}
.swiper-wp8-vertical {-ms-touch-action:pan-x; touch-action:pan-x;}
/* Arrows */
.swiper-button-prev,
.swiper-button-next {position:absolute; top:50%; z-index:10; cursor:pointer;}
.swiper-button-prev.swiper-button-disabled,
.swiper-button-next.swiper-button-disabled {opacity:0.35; cursor:auto; pointer-events:none;}
/* Fade */
.swiper-container-fade.swiper-container-free-mode .swiper-slide {-webkit-transition-timing-function:ease-out; -moz-transition-timing-function:ease-out; -ms-transition-timing-function:ease-out; -o-transition-timing-function:ease-out; transition-timing-function:ease-out;}
.swiper-container-fade .swiper-slide {pointer-events:none;}
.swiper-container-fade .swiper-slide .swiper-slide {pointer-events:none;}
.swiper-container-fade .swiper-slide-active,
.swiper-container-fade .swiper-slide-active .swiper-slide-active {pointer-events:auto;}
/* Scrollbar */
.swiper-scrollbar {border-radius:10px; position:relative; -ms-touch-action:none; background:rgba(0, 0, 0, 0.1);}
.swiper-container-horizontal > .swiper-scrollbar {position:absolute; left:1%; bottom:3px; z-index:50; height:5px; width:98%;}
.swiper-container-vertical > .swiper-scrollbar {position:absolute; right:3px; top:1%; z-index:50; width:5px; height:98%;}
.swiper-scrollbar-drag {height:100%; width:100%; position:relative; background:rgba(0, 0, 0, 0.5); border-radius:10px; left:0; top:0;}
.swiper-scrollbar-cursor-drag {cursor:move;}
/* pagination */
.paginationDot {width:100%; height:auto; z-index:5; text-align:center;}
.paginationDot .swiper-pagination-switch {display:inline-block; position:relative; width:1.5rem; height:1.5rem; margin:0 0.15rem; cursor:pointer;}
.paginationDot .swiper-pagination-switch:after {content:' '; display:block; position:absolute; top:50%; left:50%; width:0.6rem; height:0.6rem; margin-top:-0.3rem; margin-left:-0.3rem; border-radius:50%; background-color:rgba(255,255,255,0.5);}
.paginationDot .swiper-active-switch:after {background-color:#fff;}

.pagination {text-align:center; padding-top:9px; height:16px; font-size:3px; line-height:3px;}
.pagination .swiper-pagination-switch {width:7px; height:7px; display:inline-block; background-color:#cbcbcb; border-radius:50%; -webkit-border-radius:50%; margin:0 3px;}
.pagination .swiper-active-switch {background-color:#000;}



/************************* LAYOUT *************************/
html, body {height:100%;}
body > .heightGrid {min-height:100%; height:auto;}
body {position:relative;}

.heightGrid {position:relative;}
.mainSection {position:absolute; left:0; top:0; bottom:0; width:100%; height:100%; background-color:#fff; z-index:500;}
.mainBlankCover {width:100%; height:100%; position:absolute; left:0; top:0; z-index:1000; display:none; background-color:transparent;}
#mainBlankCover {background-color:rgba(0,0,0,.7);}
#contBlankCover {display:none; position:absolute; left:0; top:0; bottom:0; width:100%; height:100%; z-index:100; background-color:rgba(0,0,0,.5);}
.container {position:relative; width:100%;}
.content {overflow:hidden; min-height:300px; padding-bottom:50px;}

.goTop {display:block; overflow:hidden; position:fixed; right:1.25rem; bottom:7rem; width:3.7rem; height:3.7rem; background:transparent url(http://fiximage.10x10.co.kr/m/2016/common/btn_top.png) 50% 50% no-repeat; background-size:100%; text-indent:-999em; z-index:10000; cursor:pointer;}

span.topHigh {bottom:60px;}
span.topHigh2 {bottom:110px;}

/************************* FORM *************************/
input {-webkit-appearance:none; -webkit-border-radius:0; outline-style:none; border:0;}
button {padding:0; margin:0; border:0; cursor:pointer;}
select {height:36px; padding:0 23px 0 10px; margin:0; background:#fff url(http://fiximage.10x10.co.kr/m/2017/common/element_select.png) no-repeat right 50%; vertical-align:middle; background-size:24px 6px; color:#888; font-size:13px; -webkit-border-radius:4px; border:1px solid #cfcfcf; -webkit-appearance:none;}
input[type=radio],
input[type=checkbox] {-webkit-border-radius:4px; -webkit-appearance:none; border:1px solid #cfcfcf; width:20px; height:20px; vertical-align:middle; background-color:#fff; margin:0;}
input[type=radio]:checked {background:#fff url(http://fiximage.10x10.co.kr/m/2017/common/element_radio.png) no-repeat 50% 50%; background-size:10px 10px;}
input[type=radio]:checked:disabled {background:#fff url(http://fiximage.10x10.co.kr/m/2016/common/element_radio_disabled.png) no-repeat 50% 50%; background-size:12px 12px;}
input[type=radio]:disabled {background:#efefef;}
input[type=checkbox]:checked {background:#fff url(http://fiximage.10x10.co.kr/m/2017/common/element_checkbox.png) no-repeat 50% 50%; background-size:12px 12px;}
input[type=checkbox]:checked:disabled {background:#fff url(http://fiximage.10x10.co.kr/m/2016/common/element_checkbox_disabled.png) no-repeat 50% 50%; background-size:12px 12px;}
input[type=checkbox]:disabled {background:#efefef;}
input[type=text],
input[type=password],
input[type=tel],
input[type=number],
input[type=email] {height:36px; padding:0 10px; margin:0; -webkit-border-radius:4px; border:1px solid #cfcfcf; font-size:13px; color:#888; vertical-align:middle;}
input[type=text]:disabled {background:#f4f7f7; color:#999;}
input[type=number]::-webkit-inner-spin-button, 
input[type=number]::-webkit-outer-spin-button {-webkit-appearance:none;}
input[type=number] {-moz-appearance:textfield;}
textarea {-webkit-appearance:none; -webkit-border-radius:0; outline-style:none; padding:7px; border-radius:4px; -webkit-border-radius:4px; border:1px solid #cfcfcf; font-size:13px; color:#888;}

select.frmSelectV16 {width:5rem; height:2.7rem; padding:0.4rem 2rem 0.4rem 0.6rem; font-size:1.2rem; color:#000; border-radius:0.2rem; border:1px solid #cbcbcb; background:#fff url(http://fiximage.10x10.co.kr/m/2016/common/select_arrow_gry.png) 100% 50% no-repeat; background-size:auto 0.6rem;}
input.frmCheckV16[type=checkbox] {border-radius:0.2rem; width:1.8rem; height:1.8rem; background-size:1.2rem auto;}
input.frmCheckV16[type=checkbox]:checked {background:#fff url(http://fiximage.10x10.co.kr/m/2017/common/element_checkbox.png) no-repeat 50% 50%; background-size:1.2rem auto;}
input.frmRadioV16[type=radio] {border-radius:0.9rem 0.9rem; width:1.8rem; height:1.8rem;}
input.frmRadioV16[type=radio]:checked {background:#fff url(http://fiximage.10x10.co.kr/m/2017/common/element_radio.png) no-repeat 50% 50%; background-size:1rem;}
input.frmInputV16[type=text],
input.frmInputV16[type=password],
input.frmInputV16[type=tel],
input.frmInputV16[type=number],
input.frmInputV16[type=email],
input.frmInputV16[type=search] {min-width:5rem; height:2.7rem; margin:0; padding:0.4rem 0.6rem; font-size:1.2rem; font-family:'helveticaNeue Roman', helveticaNeueRoman, helveticaNeue, helvetica, sans-serif !important; color:#000; border-radius:0.2rem; border:1px solid #cbcbcb;}
input.frmInputV16[type=number] {padding:0.4rem 0.6rem;}
textarea.frmTxtareaV16 {min-width:5rem; padding:0.6rem 0.6rem 0.4rem 0.6rem; margin:0; font-family:'helveticaNeue Roman', helveticaNeueRoman, helveticaNeue, helvetica, sans-serif !important; font-size:1.1rem; color:#999; border-radius:0.2rem; border:1px solid #cbcbcb; vertical-align:top;}



/************************* POPUUP, MODAL POPUP *************************/
div.popWin {padding-bottom:0;}
.popWin .header {position:fixed; left:0; top:0; width:100%; height:4.86rem; padding:1.62rem 1.19rem 0 1.54rem; background-color:#f6f6f6; background-image:none; z-index:900; font-family:'helveticaNeue Roman', helveticaNeueRoman, helveticaNeue, helvetica, sans-serif !important; text-align:center;}
.popWin .header:after {content:' '; position:absolute; bottom:0; left:0; width:100%; height:1px; background-color:rgba(0, 0, 0, 0.1);}
.popWin .header h1 {width:75%; padding:0; margin:0 auto;  color:#000; font:1.45rem 'AvenirNext-Regular', 'AppleSDGothicNeo-Regular';  vertical-align:middle; letter-spacing:0;}
.popWin .header .btnPopClose {overflow:hidden; display:block; position:absolute; top:50%; right:1.19rem; width:2.39rem; height:2.22rem; margin-top:-1.11rem;}
.popWin .header .btnPopClose .pButton {width:100%; height:2.22rem; background-color:transparent; background-position:-13.01rem -1.83rem; text-indent:-999em;}
.popWin .content {padding:4.86rem 0;}

/* layer */
#layerScroll {position:absolute; left:0; bottom:0; right:0; top:4.78rem; width:100%;}
#scrollarea {overflow:hidden; position:absolute; left:0; bottom:0; right:0; top:0; z-index:1; -webkit-tap-highlight-color:rgba(0,0,0,0); width:100%; height:100%;}
#scrollarea2 {overflow:hidden; position:absolute; left:0; bottom:0; right:0; top:0; z-index:1; -webkit-tap-highlight-color:rgba(0,0,0,0); width:100%; height:100%;}
.layerPopup {overflow:hidden; position:fixed; left:0; top:0; right:0; bottom:0; width:100%; height:100%; z-index:100000; background:#fff;}
.layerPopup .popWin .content {padding:0;}

/************************* CONTENT COMMON *************************/
/* paging */
.pagingV15a {margin-top:1.96rem; text-align:center;}
.pagingV15a span {display:inline-block; min-width:2.39rem; height:2.39rem; margin:0 0.34rem; color:#838383; font:1.37rem/2.3rem 'AvenirNext-Regular', 'RobotoRegular', sans-serif;}
.pagingV15a .arrow {position:relative; text-indent:-9999em;}
.pagingV15a a {display:block; width:100%; height:100%; padding-top:0.09rem;}
.pagingV15a a.arrow {background:none;}
.pagingV15a .prevBtn {margin-right:0.09rem;}
.pagingV15a .nextBtn {margin-left:0.26rem;}
.pagingV15a .current {color:#0d0d0d; font-family:'AvenirNext-DemiBold'; font-weight:bold;}
.pagingV15a .arrow a:after {content:' '; position:absolute; top:50%; left:50%; width:1.54rem; height:1.54rem; margin:-0.77rem 0 0 -0.77rem; background-position:-4.05rem -9.56rem;}
.pagingV15a .prevBtn a:after {transform:rotateY(180deg); -webkit-transform:rotateY(180deg);}

/* breadcrumb */
.breadcrumbV15a {overflow:hidden; height:3.7rem; padding-top:0.7rem; padding-right:0.8rem; border-bottom:1px solid #c9cbcb; background-color:#dfe3e3; color:#676767; font-size:1.1rem; }
.breadcrumbV15a em {height:0.9rem; padding:0.6rem 1.2rem 0.7rem 0.6rem; background:url(http://fiximage.10x10.co.kr/m/2014/common/blt_arrow_rt6.png) no-repeat 100% 50%; background-size:0.5rem auto; box-sizing:content-box; -webkit-box-sizing:content-box; -moz-box-sizing:content-box;}
.breadcrumbV15a em:last-child {background:none;}
.breadcrumbV15a p {display:inline-block; margin:0; padding:0 0.2rem 0 0.6rem;}
.breadcrumbV15a .button {height:2.2rem; padding:0; border:none;}
.breadcrumbV15a .button a {padding:0.6rem 2rem 0.4rem 1rem; background:#a8a8a8 url(http://fiximage.10x10.co.kr/m/2014/common/blt_arrow_btm2.png) 94% 50% no-repeat; background-size:0.9rem auto; line-height:1.2;}
.breadcrumbV15a .swiper-container {margin:0 0.5rem;}
.breadcrumbV15a * {white-space:nowrap;}

.locationV16a {height:3.8rem; background-color:#dfe3e3; border-bottom:1px solid #c9cbcb; font-family:'helvetica Neue', helveticaNeue, helvetica, sans-serif !important;}
.locationV16a .swiper-container {padding:0 0.5rem;}
.locationV16a .swiper-slide {float:left; height:3.7rem; padding:0 1rem 0 0.5rem; font-size:1.1rem; line-height:1.2; vertical-align:middle; color:#676767;}
.locationV16a .swiper-slide:after {display:block; position:absolute; right:0; top:50%; width:0.45rem; height:0.9rem; margin-top:-0.45rem; background:url(http://fiximage.10x10.co.kr/m/2016/common/blt_arrow_gry3.png) no-repeat 50% 50%; background-size:100% auto; content:'';}
.locationV16a .swiper-slide:last-child {padding:0 0.5rem; color:#000; font-weight:500;}
.locationV16a .swiper-slide:last-child:after {display:none;}
.locationV16a .swiper-slide a {display:block; width:100%; height:100%; padding:1.3rem 0;}

/* sorting option 20160509 */
.viewSortV16a {position:relative;}
.sortV16a {position:relative; display:table; width:100%; border-collapse:collapse; border-spacing:0; border:0; z-index:110; vertical-align:top;}
.sortV16a .sortGrp {display:table-cell; vertical-align:top; height:4rem; border-left:1px solid #e4e4e4; border-right:1px solid #e4e4e4; z-index:10005;}
.sortV16a .sortGrp:first-child {border-left:0;}
.sortV16a .sortGrp:last-child {border-right:0;}
.sortV16a .sortGrp button, .viewSortV16a .sortGrp p a {overflow:hidden; position:relative; display:block; width:100%; height:4rem; background-color:#f4f7f7; font-size:1.2rem; color:#676767; border-bottom:1px solid #e4e4e4; text-align:left; -webkit-appearance:none; border-radius:0; -webkit-border-radius:0; outline:none;}
.sortV16a .sortGrp button {padding-left:1.25rem; padding-right:2.7rem; text-overflow:ellipsis; white-space:nowrap;}
.sortV16a .sortGrp button:after {display:block; position:absolute; width:1rem; right:1.25rem; top:50%; background-repeat:no-repeat; background-position:100% 0; content:'';}
.sortV16a .sortGrp p a {position:relative; text-align:center; line-height:4rem;}
.sortV16a .sortGrp p a:before {display:block; position:absolute; left:50%; top:50%; content:'';}
.sortV16a .sortGrp .sortNaviV16a {display:none; position:absolute; top:4rem; left:0; right:0; background-color:#fff; border-bottom:1px solid #e4e4e4;}
.sortV16a .sortGrp:last-child .sortNaviV16a {margin-right:0; border-right:0;}
.sortV16a div.current .sortNaviV16a {display:block;}
.sortV16a div.current button {background-color:#fff; border-bottom:1px solid #fff;}
.sortV16a div.current button:after {background-position:50% 100%;}

.sortV16a .category button:after {height:0.55rem; margin-top:-0.275rem; background-image:url(http://fiximage.10x10.co.kr/m/2016/common/blt_updown_gryred.png); background-size:0.9rem auto;}
.sortV16a .category .sortNaviV16a li {width:33.333%;}
.sortV16a .category .sortNaviV16a.depth2 li {width:50%;}
.sortV16a .array {position:relative;}
.sortV16a .array button:after {height:0.6rem; margin-top:-0.3rem; background-image:url(http://fiximage.10x10.co.kr/m/2017/common/blt_updown_gryred2.png); background-size:0.9rem auto;}
.sortV16a .array .sortNaviV16a {margin-left:-1px; margin-right:-1px; border-right:1px solid #e4e4e4; border-left:1px solid #e4e4e4;}
.sortV16a .array .sortNaviV16a li {width:100%;}

.sortNaviV16a ul {overflow:hidden; padding:1.2rem 0 1.35rem;}
.sortNaviV16a li {float:left;}
.sortNaviV16a li a {overflow:hidden; display:block; width:100%; padding:0.9rem 0.8rem 0.9rem 1rem; font-size:1.2rem; color:#888; letter-spacing:-0.03rem; text-overflow:ellipsis; white-space:nowrap;}
.sortNaviV16a li.selected a {font-weight:bold; color:#ff3131;}

/* 기획전,이벤트(20160509 적용) */
.evtIndexV16a .listCardV16 {margin-top:0;}
.evtIndexV16a .listCardV16 ul li {border-bottom:1px solid #e4e4e4;}
.evtIndexV16a .listCardV16 ul li .desc {position:relative; padding:1.7rem 1.25rem 1.5rem 1.25rem; line-height:1.7rem;}
.evtIndexV16a .listCardV16 ul li p strong {overflow:hidden; display:-webkit-box; width:75%; max-height:3.6rem; -webkit-line-clamp:2; -webkit-box-orient:vertical; word-wrap:break-word; text-overflow:ellipsis;}
.evtIndexV16a .listCardV16 ul li p span {margin-top:0.55rem; font-size:1.2rem; color:#888; line-height:1.4rem;}

.evtIndexV16a .emptyMsgV16a {height:100%; padding-top:10rem; padding-bottom:17.5rem; background-color:#fff;}
.emptyExhtV16a div {padding-top:10rem; background-image:url(http://fiximage.10x10.co.kr/m/2016/common/img_no_exhibit.png); background-size:6.875rem auto;}
.emptyEvtV16a div {padding-top:10rem; background-image:url(http://fiximage.10x10.co.kr/m/2016/common/img_no_event.png); background-size:7.725rem auto;}

.stickySort .viewSortV16a {position:fixed; top:0;}
div.evtSortV16a .category {width:70%;}
div.evtSortV16a .array {width:30%;}

/* BEST(sorting option, 201605 적용) */
div.bestSortV16a .category {width:34%;}
div.bestSortV16a .array {width:33%;}
div.bestSortV16a .array .sortNaviV16a {margin-right:-1px; border-right:1px solid #e4e4e4;}
div.bestSortV16a .linkBtn {width:33%;}
div.bestSortV16a .prmBtn {padding-left:1.5rem;}
div.bestSortV16a .prmBtn:before {margin:-0.7rem 0 0 -3.3rem; width:1.05rem; height:1.4rem; background:url(http://fiximage.10x10.co.kr/m/2016/common/ico_premium_best.png) 50% 50% no-repeat; background-size:1.05rem auto;}

/* NEW(sorting option, 201605 적용) */
div.newSortV16a .category {width:70%;}
div.newSortV16a .array {width:30%;}
.newIdxV15a .pdtListWrapV15a .pdtListV15a {background:none;} /* 추후 css옮김 수정 */

/* SALE(sorting option, 201605 적용) */
div.saleSortV16a .category {width:34%;}
div.saleSortV16a .array {width:33%;}
div.saleSortV16a .array .sortNaviV16a {margin-right:-1px; border-right:1px solid #e4e4e4;}
div.saleSortV16a .linkBtn {width:33%;}
div.saleSortV16a .clrncBtn {padding-left:1.75rem;}
div.saleSortV16a .clrncBtn:before {margin:-0.75rem 0 0 -4.2rem; width:1.5rem; height:1.5rem; background:url(http://fiximage.10x10.co.kr/m/2016/common/ico_clearance.png) 50% 50% no-repeat; background-size:1.5rem auto;}
.saleIdxV15a .pdtListWrapV15a .pdtListV15a {background:none;} /* 추후 css옮김 수정 */

/* GIFT(sorting option, 201605 적용) */
.giftV15a .btnAreaV16a {padding:1.5rem 1.25rem 0 1.25rem;}
.giftV15a .btnAreaV16a .btnV16a {height:2.7rem; font-size:1.2rem; font-weight:500;}
.giftV15a .btnAreaV16a .btnTalkWt {width:69.5%;}
.giftV15a .btnAreaV16a .btnTalkWt button img {width:1.175rem; margin:0.2rem 0.5rem 0 0; vertical-align:top;}
.giftV15a .giftArticle {padding-top:1.5rem;}
.giftV15a .giftArticle .desc:before, .giftV15a .giftArticle .desc:after {background:#e7eaea;}
.giftV15a .giftHint {padding-top:1.5rem;}
.giftV15a .giftHint .hint .topic {background-color:#e7eaea;}

/* WRAPPPING(sorting option, 201605 적용) */
div.wrappSortV16a .category {width:34%;}
div.wrappSortV16a .array {width:33%;}
.wrapListV16a .pdtListWrapV15a .pdtListV15a {background:none;} /* 추후 css옮김 수정 */

/* CATEGORY LIST(sorting option, 201605 적용) */
div.breadcrumbV15a {height:3rem; padding-top:0.45rem; background-color:#e7eaea; border-bottom:1px solid #e4e4e4;}
div.ctgySortV16a .category {width:43.4375%;}
div.ctgySortV16a .array {width:43.4375%;}
div.ctgySortV16a .linkBtn {width:13.125%;}
div.ctgySortV16a .fltrBtn {overflow:hidden; text-indent:-999rem; background:url(http://fiximage.10x10.co.kr/m/2016/common/ico_filter.png) 50% 50% no-repeat; background-size:1.7rem auto;}
.ctgyListV15a .pdtListWrapV15a .pdtListV15a {background:none;} /* 추후 css옮김 수정 */

/* WISH(sorting option, 201605 적용) */
div.wishMainV15a .content {padding:0;}
div.wishListV15a {padding:1rem 0.5rem;}
div.wishSortV16a .category {width:70%;}
div.wishSortV16a .linkBtn {width:30%;}
div.wishSortV16a .mywishBtn {padding-left:1.5rem;}
div.wishSortV16a .mywishBtn:before {margin:-0.5rem 0 0 -2.9rem; width:1rem; height:0.9rem; background:url(http://fiximage.10x10.co.kr/m/2016/common/ico_mywish.png) 50% 50% no-repeat; background-size:1rem auto;}

/* BRAND(sorting option, 201605 적용) */
.brViewV16a .pdtListWrapV15a .pdtListV15a {background:none;} /* 추후 css옮김 수정 */
.brandSortV16a .category {width:70%;}
.brandSortV16a .array {width:30%;}
div.zzimListWrap {padding-top:1.5rem;}


/************************* ZIPCODE *************************/
/* zipcode 2017 */
.zipcodeV17 legend {overflow:hidden; visibility:hidden; position:absolute; top:-1000%; width:0; height:0; line-height:0;}
.zipcodeV17 .tabs li a {display:block; text-align:center;}
.zipcodeV17 .commonTabV16a li a {position:relative; padding:1.4rem 0; color:#676767; font-size:1.3rem; font-weight:bold;}
.zipcodeV17 .commonTabV16a .on {color:#ff3131;}
.zipcodeV17 .commonTabV16a .on:after {content:' '; display:block; position:absolute; bottom:-1px; left:0; width:100%; height:3px; background-color:#ff3131;}
.zipcodeV17 .tabcontainer .tabcont {padding-bottom:2rem;}

.zipcodeV17 .searchForm {padding:1.5rem 1.25rem; background-color:#fff;}
.zipcodeV17 .searchForm input,
.zipcodeV17 .searchForm select {width:100%;}
.zipcodeV17 .searchForm input::-webkit-input-placeholder {color:#888;}
.zipcodeV17 .searchForm input::-moz-placeholder {color:#888;} /* firefox 19+ */
.zipcodeV17 .searchForm input:-ms-input-placeholder {color:#888;} /* ie */
.zipcodeV17 .searchForm input:-moz-placeholder {color:#888;}

.zipcodeV17 .searchForm .finder {position:relative; padding-right:8.6rem;}
.zipcodeV17 .searchForm .finder .inner {position:relative;}
.zipcodeV17 .searchForm .finder input[type=search]::-webkit-search-cancel-button {-webkit-appearance:none;}
.zipcodeV17 .searchForm .finder input[type=reset] {position:absolute; top:50%; right:1.1rem; width:1.1rem; height:1.1rem; margin-top:-0.55rem; background:transparent url(http://fiximage.10x10.co.kr/m/2017/common/btn_delete_black.png) 50% 50% no-repeat; background-size:1.1rem auto; color:transparent;}
.zipcodeV17 .searchForm .finder input[type=submit] {position:absolute; top:0; right:0; width:8.1rem; height:2.7rem; font-size:1.3rem;}

.zipcodeV17 .searchForm ul li {position:relative; margin-top:0.9rem; padding-left:8.75rem;}
.zipcodeV17 .searchForm ul li:first-child {margin-top:0;}
.zipcodeV17 .searchForm ul li label {width:8.75rem; position:absolute; top:0; left:0; height:2.7rem; font-size:1.3rem; line-height:2.7rem; color:#000;}

.zipcodeV17 .btnAreaV16a {margin-top:1.75rem; padding:0;}
.zipcodeV17 .btnAreaV16a .half {float:left; width:50%; padding:0 0.25rem;}
.zipcodeV17 .btnAreaV16a .half:first-child {padding-left:0;}
.zipcodeV17 .btnAreaV16a .half:last-child {padding-right:0;}
.zipcodeV17 .btnAreaV16a a {display:block; line-height:4rem;}

.zipcodeV17 .guide {padding:3.2rem 0 4.5rem; text-align:center;}
.zipcodeV17 .guide p {position:relative; padding-top:5.95rem; color:#000; font-size:1.3rem; line-height:1.5em;}
.zipcodeV17 .guide p:after {content:' '; display:block; position:absolute; top:0; left:50%; width:5rem; height:5rem; margin-left:-2.5rem; background:url(http://fiximage.10x10.co.kr/m/2017/common/ico_search.png) 50% 0 no-repeat; background-size:100% auto;}
.zipcodeV17 .noData p:after {background-position:50% 100%;}

.zipcodeV17 .tip {margin:0 1.25rem 1.25rem; background-color:#fff; text-align:center;}
.zipcodeV17 .tip h3 {height:4.1rem; padding-top:1.45rem; color:#ff3131; font-size:1.2rem;}
.zipcodeV17 .tip h3 span {display:inline-block; width:2.1rem; height:1.3rem; padding-top:0.1rem; border:1px solid #ff3131; border-radius:0.8rem; font-size:0.9rem; line-height:1.3rem; text-transform:uppercase; vertical-align:0.15rem;}
.zipcodeV17 .tip ul {padding-bottom:2rem; border-top:0.2rem solid #f4f7f7;}
.zipcodeV17 .tip ul li {margin-top:2.2rem; color:#000; font-size:1.3rem;}
.zipcodeV17 .tip ul li span {display:block; margin-top:0.4rem; color:#888; font-size:1.1rem;}

.zipcodeV17 .result .total {margin:1.5rem 1.25rem 0; padding-bottom:0.9rem; color:#888; font-size:1.1rem;}
.zipcodeV17 .result .total em {color:#000;}
.zipcodeV17 .result ul {margin:0 1.25rem;}
.zipcodeV17 .result ul li {margin-top:1.5rem; background-color:#fff;}
.zipcodeV17 .result ul li:first-child {margin-top:0;}
.zipcodeV17 .result ul li span,
.zipcodeV17 .result ul li a {overflow:hidden; display:block; color:#000; font-size:1.3rem; line-height:1.375em;}
.zipcodeV17 .result ul li span {padding:1.1rem 1.25rem 0.8rem;}
.zipcodeV17 .result ul li a {overflow:hidden; display:table; width:100%; padding:0.9rem 1.25rem 0.9rem; border-top:1px solid #f4f7f7;}
.zipcodeV17 .result ul li div {display:table-cell; width:auto; padding-right:1.5rem; background:url(http://fiximage.10x10.co.kr/m/2017/common/blt_arrow_black_22x38.png) 100% 50% no-repeat; background-size:0.55rem auto;}
.zipcodeV17 .result ul li em {display:table-cell; width:3.4rem; color:#ff3131;}
.zipcodeV17 .result .pagingV15a {margin-top:2rem;}

.zipcodeV17 .addressForm {margin:0 1.25rem 0;}


/************************* ELEMENT *************************/
/* Tab */
.tabNav:after {content:" "; display:block; clear:both;}
.tabNav li {position:relative; float:left; text-align:center;}
.tabNav li a {display:block; height:100%;}
.tabNav li.current a {color:#ff3131; font-weight:bold;}
.tNum2 li {width:50%;}
.tNum3 li {width:33.33333%;}
.tNum4 li {width:25%;}

.tab01, .tab02 {width:100%;}
.tab01 .tabNav li {height:35px;}
.tab01 .tabNav li a {padding-top:11px; font-size:12px; line-height:13px; color:#676767; text-shadow:1px 1px 0 #fff;}
.tab01 .tabNav li.current {box-shadow:0px 1px 1px 0px rgba(0, 0, 0, 0.15); background-color:#fff;}
.tab01 .tabNav li.current:before {content:''; position:absolute; width:100%;  height:3px; left:0; top:0; background-color:#ff3131; border-radius:3px 3px 0 0; }
.tab01 .tabNav li.current:after {content:''; position:absolute; left:50%; bottom:4px; margin-left:-3px; width:0; height:0; border-style:solid; border-width:3px 3px 0 3px; border-color:#ff3131 transparent transparent transparent;}
.tab01 .tabNav li.current a {color:#ff3131; font-weight:bold;}
.tab01 .tabNav li.current span {position:absolute; width:100%; height:10px; left:0; bottom:-4px; background-color:#fff;}

.tab02 .tabNav {border-bottom:1px solid #cbcbcb;}
.tab02 .tabNav li {font-size:12px;}
.tab02 .tabNav li a {padding:12px 0 15px 0;}
.tab02 .tabNav li.current:after {content:' '; display:inline-block; position:absolute; left:0; bottom:-1px; width:100%; height:3px; background:#ff3131;}

ul.commonTabV16a {display:table; width:100% !important; height:3.9rem; border-bottom:1px solid #e4e4e4; background-color:#fff;}
.commonTabV16a li {display:table-cell; position:relative; height:100%; vertical-align:middle; text-align:center; font-size:1.3rem; color:#676767; letter-spacing:-0.025rem; white-space:nowrap; font-weight:600;}
.commonTabV16a li span {font-size:1rem;}
.commonTabV16a li.current {color:#ff3131;}
.commonTabV16a li.current:after {position:absolute; left:0; bottom:-1px; content:''; width:100%; height:3px; background-color:#ff3131;}

.floatingBar {position:fixed; left:0; right:0; bottom:0; z-index:100; width:100%; border-top:1px solid #e8e8e8; box-shadow:0 0 3px #cfcfcf; background:#fff; height:52px;}
.floatingBar .btnWrap {padding:5px; height:52px;}
.floatingBar .bNum2 {padding:2.5px !important;}
.floatingBar .bNum2 .ftBtn {float:left; width:50%; padding:2.5px;}

.btnBarV16a {overflow:hidden; width:100%; margin-left:1px;}
.btnBarV16a li {float:left; height:2.7rem; margin-left:-1px; text-align:center;}
.btnBarV16a li div {width:100%; height:2.7rem; border:1px solid #e5e5e5; border-radius:0; -webkit-appearance:none; -webkit-border-radius:0; background-color:#f4f7f7; color:#999; font-size:1.2rem; line-height:2.5rem; vertical-align:middle; z-index:1;}
.btnBarV16a li:first-child {margin-left:0;}
.btnBarV16a li:first-child div {border-top-left-radius:0.2rem; border-bottom-left-radius:0.2rem;}
.btnBarV16a li:last-child div {border-top-right-radius:0.2rem; border-bottom-right-radius:0.2rem;}
.btnBarV16a li.current {position:relative;}
.btnBarV16a li.current div {background-color:#fff; color:#000; border-color:#cbcbcb; font-weight:600; line-height:2.6rem;}

/* button */
.btnWrap {overflow:hidden;}
/*.btnWrap > div {display:table-cell;}*/
.button {display:inline-block; border-radius:3px; vertical-align:middle;}
.button a,.button input,.button button {display:block; width:100%; border:0; margin:0; line-height:1; text-align:center; cursor:pointer; color:inherit; background-color:inherit; border-radius:3px; white-space:nowrap;}

.btB1 a,.btB1 input,.btB1 button {padding:13px 35px 12px; font-size:14px; font-weight:bold;}
.btB2 a,.btB2 input,.btB2 button {padding:11px 16px 10px; font-size:12px; font-weight:bold;}

.btM1 a,.btM1 input,.btM1 button {padding:9px 16px 8px; font-size:14px; font-weight:bold;}
.btM2 a,.btM2 input,.btM2 button {padding:7px 16px 6px; font-size:13px;}

.btS1 a,.btS1 input,.btS1 button {padding:5px 13px 4px; font-size:11px; font-weight:bold;}
.btS2 a,.btS2 input,.btS2 button {padding:3px 6px; font-size:11px;}

.btRed {border:1px solid #ff3131; background-color:#ff3131;}
.btWht {border:1px solid #cbcbcb; background-color:#fff;}
.btBck {border:1px solid #000; background-color:#000;}
.btGry {border:1px solid #cfcfcf; background-color:#dadddd;}
.btGry2 {border:1px solid #a8a8a8; background-color:#a8a8a8;}
.btGry3 {border:1px solid #373a3a; background-color:#373a3a;}
.btGrn {border:1px solid #1b8e09; background-color:#1b8e09;}
.btRedBdr {border:1px solid #ff3131;}
.btRedBdr.cWh1 a {color:#ff3131 !important;}
.btGryBdr {border:1px solid #cbcbcb;}
.btBckBdr {border:1px solid #000;}

.rdArr {display:inline-block; width:21px; height:6px; margin-top:-4px; vertical-align:middle; background:url(http://fiximage.10x10.co.kr/m/2014/common/blt_arrow_rt7.png) right top no-repeat; background-size:15px 6px;}
.rdArr2 {display:inline-block; width:12px; height:9px; margin-top:2px; vertical-align:top; background:url(http://fiximage.10x10.co.kr/m/2014/common/blt_arrow_rt.png) right top no-repeat; background-size:7px 9px;}
.rdArr3 {display:inline-block; width:7px; height:10px; margin:1px 0 0 5px; vertical-align:top; background:url(http://fiximage.10x10.co.kr/m/2014/common/blt_arrow_rt5.png) right top no-repeat; background-size:7px 10px;}
.plusArr {display:inline-block; width:9px; height:9px; margin:2px 5px 0 0; vertical-align:top; background:url(http://fiximage.10x10.co.kr/m/2014/common/ico_plus.png) left top no-repeat; background-size:100% 100%;}
.checkArr {display:inline-block; width:9px; height:8px; margin:2px 5px 0 0; vertical-align:top; background:url(http://fiximage.10x10.co.kr/m/2014/common/ico_check3.png) left top no-repeat; background-size:100% 100%;}
.rdBtn2 {display:block; padding-right:7px; background:url(http://fiximage.10x10.co.kr/m/2014/common/blt_arrow_rt.png) right 50% no-repeat; background-size:4px 6px;}
.downArr {display:inline-block; width:14px; height:12px; margin:-3px 0 0 5px; vertical-align:middle; background:url(http://fiximage.10x10.co.kr/m/2014/common/ico_down.png) left bottom no-repeat; background-size:100% 100%;}
.moreArr {display:inline-block; width:10px; height:10px; margin:-4px 0 0 5px; vertical-align:middle; background:url(http://fiximage.10x10.co.kr/m/2017/common/ico_detail2.png) left bottom no-repeat; background-size:100% 100%;}
.rdArrLt {display:inline-block; width:9px; height:12px; margin:-2px 5px 0 0; vertical-align:middle; background:url(http://fiximage.10x10.co.kr/m/2014/common/blt_arrow_lt3.png) left top no-repeat; background-size:9px 12px;}

.btnAreaV16a {display:table; width:100%; padding:1.5rem 1.25rem 1.75rem 1.25rem;}
.btnAreaV16a p {display:table-cell;}
.btnAreaV16a .btnV16a {width:100%; height:4rem; font-size:1.5rem; font-weight:600;}

.btnV16a {margin:0; border-radius:0.2rem; background-color:#fff; font-size:1rem; letter-spacing:-0.05rem; text-align:center;}
.btnLGryV16a {background-color:#fff; border:1px solid #e4e4e4; color:#888;}
.btnGrnV16a {background-color:#1b8e09; border:1px solid #1b8e09; color:#fff;}
.btnRed1V16a {background-color:#fff; border:1px solid #ff3131; color:#ff3131;}
.btnRed2V16a {background-color:#ff3131; border:1px solid #ff3131; color:#fff;}
.btnWht1V16a {background-color:#fff; border:1px solid #fff; color:#ff3131;}
.btnWht2V16a {background-color:#fff; border:1px solid #e4e4e4; color:#888;}
.btnBlk1V16a {background-color:#000; border:1px solid #000; color:#fff;}
.btnDGryV16a {background-color:#676767; border:1px solid #676767; color:#fff;}
.btnMGryV16a {background-color:#a4aaaa; border:1px solid #a4aaaa; color:#fff;}
.btnBlu1V16a {background-color:#419ed1; border:1px solid #419ed1; color:#fff;}
.btnLinkBl {padding-right:0.75rem; font-size:1.1rem; color:#0154d0; background:url(http://fiximage.10x10.co.kr/m/2016/common/blt_arrow_blue.png) 100% 50% no-repeat; background-size:0.375rem auto;}



/************************* ETC *************************/
/* 20160509 */
.hide {display:none;}
.hidden {visibility:hidden; width:0; height:0; overflow:hidden; position:absolute; top:-1000%; line-height:0;}

.fs1r {font-size:1rem;}
.fs1-1r {font-size:1.1rem; line-height:1.3;}
.fs1-2r {font-size:1.2rem !important; line-height:1.4;}
.fs1-3r {font-size:1.3rem !important; line-height:1.4;}
.fs1-5r {font-size:1.5rem !important; line-height:1.5;}
.fs1-6r {font-size:1.6rem;}
.fs1-9r {font-size:1.9rem;}

.cLGy1V16a {color:#999 !important;} /* light grey */
.cMGy1V16a {color:#888 !important;} /* mid grey */
.cDGy1V16a {color:#676767 !important;} /* dark grey */
.cBk1V16a {color:#000 !important;}
.cRd1V16a {color:#ff3131 !important;}
.cBl1V16a {color:#0154d0 !important;}
.cABl1V16a {color:#8df7f9 !important;} /* aqua blue */
.cTqi1V16a {color:#1fbcb6 !important;} /* turquoise */
.cGr1V16a {color:#00a061 !important;} /* green */

.tMar0-2r {margin-top:0.2rem;}
.tMar0-3r {margin-top:0.3rem;}
.tMar0-4r {margin-top:0.4rem !important;}
.tMar0-5r {margin-top:0.5rem !important;}
.tMar0-6r {margin-top:0.6rem !important;}
.tMar0-7r {margin-top:0.7rem !important;}
.tMar0-8r {margin-top:0.8rem !important;}
.tMar0-9r {margin-top:0.9rem !important;}
.tMar1r {margin-top:1rem !important;}
.tMar1-1r {margin-top:1.1rem !important;}
.tMar1-3r {margin-top:1.3rem;}
.tMar1-8r {margin-top:1.8rem;}
.tMar3r {margin-top:3rem;}
.lMar0-4r {margin-left:0.4rem;}
.lMar0-5r {margin-left:0.5rem;}
.lMar2-5r {margin-left:2.5rem;}
.lMar3r {margin-left:3rem;}
.bMar0-5r {margin-bottom:0.5rem;}

.tPad3r {padding-top:3rem;}
.tPad0-1r {padding-top:0.1rem;}
.tPad0-9r {padding-top:0.9rem !important;}
.tPad1r {padding-top:1rem !important;}
.tPad1-5r {padding-top:1.5rem !important;}
.lPad0-5r {padding-left:0.5rem;}
.lPad0-6r {padding-left:0.6rem !important;}
.lPad1r {padding-left:1rem !important;}
.bPad0-3r {padding-bottom:0.3rem !important;}
.bPad0-4r {padding-bottom:0.4rem !important;}

.grid1 {width:100%;}
.grid2, .grid2 li, .grid2 .swiper-pagination-switch {width:50%;}
.grid3, .grid3 li, .grid3 .swiper-pagination-switch {width:33.33%;}
.grid4, .grid4 li, .grid4 .swiper-pagination-switch {width:25%;}
.grid5, .grid5 li, .grid5 .swiper-pagination-switch {width:20%;}

.inner1r {padding:1rem;}

/* old version */
.cRd1 {color:#ff3131 !important;}
.cBl1 {color:#0060ff !important;}
.cBl2 {color:#18b1d7 !important;}
.cBl3 {color:#287fbf !important;}
.cBl4 {color:#0154d0 !important;}
.cGy1 {color:#999 !important;}
.cGy2 {color:#676767 !important;}
.cGy3 {color:#555 !important;}
.cWh1, .cWh1 a {color:#fff !important;}
.cBk1 {color:#000 !important;}
.cGr1 {color:#00a061 !important;}
.cGr2 {color:#1fbcb6 !important;}
.cOr1 {color:#ff6000 !important;}
.cPk1 {color:#ff0cf1 !important;}

.bgGry {background-color:#f4f7f7;}
.bgGry2 {background-color:#a8a8a8;}
.bgWht {background-color:#fff !important;} /* 계속 유지 */
.bgRed {background-color:#d60000;}
.bgGrn {background-color:#1b8e09;}

.inner5 {padding-left:5px; padding-top:5px; padding-right:5px; padding-bottom:5px; overflow:hidden;}
.inner10 {padding-left:10px; padding-top:10px; padding-right:10px; padding-bottom:10px; overflow:hidden;}

.col02 {width:49%;}
.col01 {width:55%;}

.w10p {width:10% !important;}
.w20p {width:20% !important;}
.w25p {width:25% !important;}
.w30p {width:30% !important;}
.w35p {width:35% !important;}
.w40p {width:40% !important;}
.w49p {width:49% !important;}
.w50p {width:50% !important;}
.w60p {width:60% !important;}
.w70p {width:70% !important;}
.w80p {width:80% !important;}
.w90p {width:90% !important;}
.w100p {width:100% !important;}

.txtLine {text-decoration:underline;}

.ct {text-align:center !important;}
.lt {text-align:left !important;}
.rt {text-align:right !important;}

.overHidden {overflow:hidden;}
.posRel {position:relative;}
.ftLt {float:left;}
.ftRt {float:right;}
.vTop {vertical-align:top !important;}

.fs10 {font-size:10px;}
.fs11 {font-size:11px !important; line-height:1.2 !important;}
.fs12 {font-size:12px;}
.fs13 {font-size:13px !important;}
.fs15 {font-size:15px;}

.lh1 {line-height:1;}
.lh12 {line-height:1.2;}
.lh14 {line-height:1.4;}

.pad0 {padding:0 !important;}
.tPad0 {padding-top:0 !important;}
.tPad05 {padding-top:5px !important;}
.tPad10 {padding-top:10px !important;}
.tPad15 {padding-top:15px !important;}
.tPad20 {padding-top:20px !important;}
.bPad10 {padding-bottom:10px !important;}
.bPad15 {padding-bottom:15px !important;}
.bPad35 {padding-bottom:35px !important;}
.bPad45 {padding-bottom:45px !important;}
.lPad05 {padding-left:5px !important;}
.lPad10 {padding-left:10px !important;}

.mar0 {margin:0 !important;}
.tMar05 {margin-top:5px !important;}
.tMar10 {margin-top:10px !important;}
.tMar15 {margin-top:15px;}
.tMar20 {margin-top:20px;}
.tMar25 {margin-top:25px !important;}
.tMar30 {margin-top:30px;}
.rMar05 {margin-right:5px;}
.bMar05 {margin-bottom:5px;}
.bMar20 {margin-bottom:20px;}
.lMar05 {margin-left:5px;}
.lMar10 {margin-left:10px;}
.lMar28 {margin-left:28px;}
.lMar30 {margin-left:30px;}

@media all and (min-width:480px){
	input[type=radio],
	input[type=checkbox] {width:30px; height:30px;}
	input[type=radio]:checked {background-size:15px 15px;}
	input[type=checkbox]:checked {background-size:18px 18px;}
	input[type=text],
	input[type=password],
	input[type=tel],
	input[type=number],
	input[type=email] {height:54px; font-size:20px;}

	select {height:54px; padding:0 35px 0 19px; background-size:24px 5px; font-size:20px;}
	textarea {padding:11px; font-size:20px;}

	.tab01 .tabNav li {height:53px;}
	.tab01 .tabNav li a {padding-top:16px; font-size:18px; line-height:19px;}
	.tab01 .tabNav li.current:before {height:5px;}
	.tab01 .tabNav li.current:after {bottom:6px; margin-left:-4px; border-width:4px 4px 0 4px;}
	.tab01 .tabNav li.current span {height:15px; bottom:-6px;}

	.tab02 .tabNav li {font-size:18px;}
	.tab02 .tabNav li a {padding:18px 0 22px 0;}
	.tab02 .tabNav li.current:after {height:4px;}

	.floatingBar .button {border-radius:3px;}
	.floatingBar .button a, .floatingBar .button input, .floatingBar .button button {border-radius:3px;}

	.floatingBar .btB1 a, .floatingBar .btB1 input, .floatingBar .btB1 button {padding:13px 35px 12px; font-size:14px;}
	.floatingBar .btB2 a, .floatingBar .btB2 input, .floatingBar .btB2 button {padding:11px 16px 10px; font-size:12px;}

	.floatingBar .btM1 a, .floatingBar .btM1 input, .floatingBar .btM1 button {padding:9px 16px 8px; font-size:14px;}
	.floatingBar .btM2 a, .floatingBar .btM2 input, .floatingBar .btM2 button {padding:7px 16px 6px; font-size:13px;}

	.floatingBar .btS1 a, .floatingBar .btS1 input, .floatingBar .btS1 button {padding:5px 13px 4px; font-size:11px;}
	.floatingBar .btS2 a, .floatingBar .btS2 input, .floatingBar .btS2 button {padding:3px 6px; font-size:11px;}

	.button {border-radius:4px;}
	.button a,.button input,.button button {border-radius:4px;}

	.btB1 a,.btB1 input,.btB1 button {padding:20px 53px 18px; font-size:21px;}
	.btB2 a,.btB2 input,.btB2 button {padding:17px 24px 16px; font-size:18px;}

	.btM1 a,.btM1 input,.btM1 button {padding:13px 24px 12px; font-size:21px;}
	.btM2 a,.btM2 input,.btM2 button {padding:10px 24px 9px; font-size:20px;}

	.btS1 a,.btS1 input,.btS1 button {padding:7px 20px 6px; font-size:17px;}
	.btS2 a,.btS2 input,.btS2 button {padding:5px 9px; font-size:17px;}

	.rdArr {width:32px; height:9px; margin-top:-5px; background-size:23px 9px;}
	.rdArr2 {width:18px; height:13px; margin-top:3px; background-size:11px 13px;}
	.rdArr3 {width:11px; height:15px; margin:1px 0 0 7px; background-size:11px 15px;}
	.plusArr { width:13px; height:13px; margin:3px 7px 0 0;}
	.checkArr {width:13px; height:12px; margin:3px 7px 0 0;}
	.rdBtn2 {padding-right:10px; background-size:6px 8px;}
	.downArr {width:21px; height:18px; margin:-4px 0 0 8px;}
	.moreArr {width:15px; height:15px; margin:-6px 0 0 8px;}
	.rdArrLt {width:13px; height:18px; margin:-3px 8px 0 0; background-size:13px 18px;}

	.inner5 {padding-left:7px; padding-top:7px; padding-right:7px; padding-bottom:7px;}
	.inner10 {padding-left:15px; padding-top:15px; padding-right:15px; padding-bottom:15px;}

	.fs10 {font-size:15px;}
	.fs11 {font-size:17px !important;}
	.fs12 {font-size:18px;}
	.fs13 {font-size:19px !important;}
	.fs15 {font-size:22px;}

	.tPad05 {padding-top:7px !important;}
	.tPad10 {padding-top:15px !important;}
	.tPad15 {padding-top:22px !important;}
	.tPad20 {padding-top:30px !important;}
	.bPad10 {padding-bottom:15px !important;}
	.bPad15 {padding-bottom:22px !important;}
	.bPad35 {padding-bottom:53px !important;}
	.bPad45 {padding-bottom:68px !important;}
	.lPad05 {padding-left:7px !important;}
	.lPad10 {padding-left:15px !important;}

	.tMar05 {margin-top:7px !important;}
	.tMar10 {margin-top:15px !important;}
	.tMar15 {margin-top:22px;}
	.tMar20 {margin-top:30px;}
	.tMar25 {margin-top:38px !important;}
	.tMar30 {margin-top:45px;}
	.rMar05 {margin-right:7px;}
	.bMar05 {margin-bottom:7px;}
	.bMar20 {margin-bottom:30px;}
	.lMar05 {margin-left:7px;}
	.lMar10 {margin-left:15px;}
	.lMar28 {margin-left:42px;}
	.lMar30 {margin-left:45px;}
}

/* ------------------ 2017 renewal --------------------- */
.default-font, .default-font button, .default-font input,.default-font textarea, .default-font select {font-family:'AvenirNext-Regular', 'AppleSDGothicNeo-Regular', 'RobotoRegular', 'Noto Sans', sans-serif;}
.default-font,
.default-font .headline {color:#0d0d0d;}
.default-font a {-webkit-tap-highlight-color:transparent;}
.default-font button {outline:none;}

/* layout */
.default-font .content {min-height:auto; padding-bottom:0;}
.body-main {padding-top:9.22rem;}
.body-sub .content,
.body-popup .content {padding-bottom:4.27rem;}
.body-1depth .content {padding-bottom:8.87rem;}
.body-1depth .tab-bar {display:block;}
.body-1depth .btn-top {bottom:6.48rem;}
.body-popup {padding-top:4.78rem;}
.body-popup.piece {padding-top:0;}
.body-popup .header-popup {top:0; left:0; z-index:50;}
.body-sub #scrollarea > div {padding-bottom:4.27rem;}

/* header */
.tenten-header {width:100%;}
.tenten-header .title-wrap {position:relative; height:4.78rem; padding:1.62rem 1.19rem 0 1.54rem; background-color:rgba(255, 255, 255, 0.93);}
.tenten-header .tenten {width:8.36rem; height:1.62rem; background-position:0 0; color:#000; font:bold 1.28rem/1.62rem 'AppleSDGothicNeo-SemiBold';}
.tenten-header .tenten a {color:transparent;}
.tenten-header .toolbar {position:absolute; top:50%; right:1.19rem; margin-top:-1.11rem;}
.toolbar:after {content:' '; display:block; clear:both;}
.toolbar a {position:relative; float:left; margin-left:1.19rem; line-height:2.22rem; text-indent:-999em;}
.toolbar a,
.tenten-header .btn-home,
.tenten-header .btn-close,
.tenten-header .btn-back {width:2.39rem; height:2.22rem;}
.btn-back,
.tenten-header .btn-home {position:absolute; top:50%; left:1.19rem; margin-top:-1.11rem; text-indent:-999em;}
.btn-back {background-position:0 -1.83rem;}
.tenten-header .btn-home {left:4.27rem; background-position:-2.6rem -1.83rem;}
.toolbar .btn-search {background-position:-5.2rem -1.83rem;}
.tenten-header .btn-share {background-position:-10.41rem -1.83rem;}
.tenten-header .btn-shoppingbag {background-position:-7.81rem -1.83rem;}
.tenten-header .badge {position:absolute; background-color:#ff3131; text-indent:0;}
.tenten-header .toolbar .badge {bottom:-0.43rem; right:-0.51rem; min-width:1.71rem; height:1.71rem; border-radius:1.71rem; padding:0.09rem 0.34rem 0; color:#fff; font:bold 0.94rem/1.71rem 'AvenirNext-DemiBold'; letter-spacing:-0.076rem; text-align:center;}
.tenten-header .new {overflow:hidden; top:50%; right:-1.28rem; width:1.19rem; height:1.19rem; margin-top:-0.68rem; border-radius:50%; background-color:#ff3131;}
.tenten-header .new:after {content:' '; position:absolute; top:0; left:0; width:100%; height:100%; background-position:-8.57rem 0;}
@media all and (min-width:360px) and (max-width:360px){
	.tenten-header .new {width:12px; height:12px; margin-top:-5px;}
	.tenten-header .new:after {background-position:-97px -1px; background-size:200px auto;}
}
.tenten-header .btn-close {position:absolute; top:50%; right:1.19rem; margin-top:-1.11rem; background-color:transparent; background-position:-13.01rem -1.83rem; text-indent:-999em;}

.header-main {position:fixed; top:0; left:0; z-index:100; transition:top 0.2s ease-in-out; border-bottom:1px solid rgba(0, 0, 0, 0.1);}
.header-main .title-wrap {transition:margin-top 1s cubic-bezier(0.86, 0, 0.07, 1);}
.nav-up .header-main {top:-4.78rem;}
/*.nav-up .tab-bar {bottom:-4.6rem;}*/

.header-sub {position:relative; background-color:#f7f7f7;}
.header-popup {position:fixed;}
.header-sub:after, .header-popup:after {content:' '; position:absolute; bottom:0; left:0; width:100%; height:1px; background-color:rgba(0, 0, 0, 0.1);}
.header-sub .title-wrap,
.header-popup .title-wrap {background-color:#f7f7f7;}
.header-sub h1,
.header-popup h1,
.header-popup h2 {overflow:hidden; width:15.01rem; margin:-0.17rem auto 0; padding-top:0.17rem; color:#000; font-size:1.45rem; line-height:1.88rem; text-overflow:ellipsis; white-space:nowrap; text-align:center;}
.header-white,
.header-white .title-wrap {background-color:rgba(255, 255, 255, 0.9);}
.header-black .title-wrap {background-color:#262626;}
.header-black:after {display:none;}
.header-black h1 {color:#fff; font-family:'AvenirNext-Regular', 'AppleSDGothicNeo-Light';}

.header-popup-transparent {position:absolute; top:0; left:0; z-index:10; border-bottom:0;}
.header-popup-transparent:after {display:none;}
.header-popup-transparent,
.header-popup-transparent .title-wrap {background-color:transparent;}

.header-black .btn-close {background-position:-7.81rem -4.26rem;}
.header-popup-transparent .btn-close {background-position:-10.41rem -4.26rem;}
.header-black .btn-back,
.header-popup-transparent .btn-back {background-position:0 -4.26rem;}
.header-black .btn-home {background-position:-2.6rem -4.26rem;}

.nav-gnb {width:100%; height:4.44rem; background-color:rgba(255, 255, 255, 0.93);}
.nav-gnb .swiper-container {padding:0 2.56rem 0 1.71rem;}
.nav-gnb .swiper-slide {margin-left:3.75rem;}
.nav-gnb .swiper-slide:first-child {margin-left:0;}
.nav-gnb a {display:block; position:relative; height:100%; font:1.28rem/4.44rem 'AvenirNext-Regular', 'AppleSDGothicNeo-Light';}
.nav-gnb .on {color:#ff3131; font-family:'AppleSDGothicNeo-SemiBold'; font-weight:bold;}
.nav-gnb .on:after {bottom:0.51rem; left:50%; margin-left:-0.215rem;}

/* tab bar */
.tab-bar {display:none; left:0; z-index:20; width:100%; height:4.6rem; border-top:1px solid #dcdcdc; background-color:#fff; transition:bottom 0.2s ease-in-out;}
.body-main .tab-bar {display:block;}
.tab-bar ul {overflow:hidden;}
.tab-bar li {float:left; width:20%; text-align:center;}
.tab-bar a {display:block; position:relative; height:100%; padding-top:2.9rem; color:#2d2d2d; font:0.85rem 'AppleSDGothicNeo-SemiBold'; letter-spacing:-0.017rem;}
.tab-bar a:after {content:' '; position:absolute; top:0.51rem; left:50%; width:2.22rem; height:2.22rem; margin-left:-1.11rem;}
.tab-bar .on {color:#ff3131;}
.tab-bar .category a:after {background-position:0 -6.69rem;}
.tab-bar .category .on:after {background-position:0 -9.13rem;}
.tab-bar .my a:after {background-position:-2.43rem -6.69rem;}
.tab-bar .my .on:after {background-position:-2.43rem -9.13rem;}
.tab-bar .home a:after {background-position:-4.86rem -6.69rem;}
.tab-bar .home .on:after {background-position:-4.86rem -9.13rem;}
.tab-bar .order a:after {background-position:-7.3rem -6.69rem;}
.tab-bar .order .on:after {background-position:-7.3rem -9.13rem;}
.tab-bar .history a:after {background-position:-9.79rem -6.69rem;}
.tab-bar .history .on:after {background-position:-9.79rem -9.13rem;}

/* footer */
.tenten-footer {display:none; border-top:1px solid #eaeaea; text-align:center;}
.body-main .tenten-footer {display:block; padding-bottom:4.6rem;}
.footer-nav, .footer-content .cs, .footer-link, .tenten-sns {display:flex; display:-webkit-flex; justify-content:center;}
.footer-nav {height:4.01rem; padding-top:1.54rem; border-bottom:1px solid #eaeaea;}
.footer-nav a {margin:0 1.02rem; color:#6e6e6e; font-size:1.11rem; line-height:1.11rem;}
.footer-nav a:first-child {margin-left:0;}
.footer-nav a:last-child {margin-right:0;}
.footer-content .copyright {margin-top:0.85rem; padding-top:0; color:#adadad; font-size:0.94rem;}
.footer-content {padding-bottom:2.47rem; background-color:#f4f4f4;}
.footer-content address {padding-top:2.56rem;}
.footer-content .tenten a {position:relative; color:#000; font-size:1.28rem; padding-right:1.71rem;}
.footer-content .tenten a:after {content:' '; position:absolute; top:50%; right:0; width:1.02rem; height:0.6rem; margin-top:-0.43rem; background-position:-4.44rem 0; transition:-webkit-transform 0.3s; transition:transform 0.3s;}
.footer-content .tenten .on:after {-webkit-transform:rotate(180deg); transform:rotate(180deg);}
.footer-content .desc {display:none;}
.footer-content .info, .footer-content .cs {font-size:1.02rem;}
.footer-content .info {margin-top:1.02rem; color:#6e6e6e; line-height:1.54rem;}
.footer-content .cs {margin-top:0.85rem; margin-bottom:1.54rem; color:#3f75ff;}
.footer-content .cs a, .footer-link a {position:relative; padding:0 0.77rem;}
.footer-content .cs a {color:#3f75ff;}
.footer-content .cs a:last-child:after,
.footer-link a:last-child:after {content:' '; position:absolute; top:50%; left:0; width:1px; height:0.85rem; margin-top:-0.51rem; background-color:#e1e1e1;}
.footer-link {margin-top:0.85rem;}
.footer-link a {color:#838383; font-size:0.94rem;}
.tenten-sns { margin-top:1.62rem;}
.tenten-sns li {display:inline-block; width:2.13rem; height:2.13rem; margin:0 1.28rem;}
.tenten-sns a {overflow:hidden; display:block; width:100%; height:100%; color:transparent;}
.tenten-sns .facebook {background-position:0 -11.6rem;}
.tenten-sns .instagram {background-position:-2.3rem -11.6rem;}
.tenten-sns .china {background-position:-4.6rem -11.6rem;}
/*.tenten-sns .thefingers {background-position:-14.34rem -11.6rem;}*/
.platform-nav {overflow:hidden; position:relative; text-align:left;}
.platform-nav:after {content:' '; position:absolute; top:50%; left:50%; width:1px; height:0.85rem; margin:-0.51rem 0 0 -0.5px; background-color:#e1e1e1;}
.platform-nav a {float:left; width:50%; height:3.93rem; padding-top:0.085rem; background-color:#fff; color:#838383; font:1.11rem/3.93rem 'AvenirNext-Medium', 'AppleSDGothicNeo-Medium';}
.platform-nav a:first-child {padding-right:2.22rem; text-align:right;}
.platform-nav a:last-child {padding-left:2.22rem;}
.btn-top,
.btn-zoom {position:fixed; right:1.02rem; z-index:20; width:4.1rem; height:4.1rem; border:1px solid #e0e0e0; background-color:rgba(255, 255, 255, 0.9); line-height:4.44rem; text-indent:-9999em; cursor:pointer;}
.btn-top {display:none; bottom:1.45rem;}
.btn-top a, .btn-zoom a {display:block; width:100%; height:100%;}
.btn-top:after, .btn-zoom:after {content:' '; position:absolute; top:50%; left:50%; width:2.22rem; height:2.22rem; margin:-1.11rem 0 0 -1.11rem; background-position:-12.16rem -6.69rem;}
.body-main .btn-top {bottom:6.05rem;}
.category-item .btn-top {bottom:6.82rem;}
.piece .btn-top {border-color:#363636; background-color:rgba(29, 29, 29, 0.9);}
.piece .btn-top:after {background-position:-12.16rem -9.13rem;}
.btn-zoom {bottom:11.95rem;}
.btn-zoom:after {background-position:-14.59rem -6.69rem;}
/*.body-main .btn-top.nav-up {bottom:1.88rem;}
.body-main .btn-top.nav-down {bottom:6.48rem;}*/
.btn-next-tab {background-color:#6794ef;}
.btn-next-tab a {display:block; padding:1.96rem 0 2.3rem; color:#ecf3ff; font-size:1.19rem; line-height:1.54rem;}
.btn-next-tab b {color:#fff; font-family:'AvenirNext-DemiBold', 'AppleSDGothicNeo-SemiBold'; font-weight:bold;}

/* on dot */
.nav-gnb .on:after,
.category-menu .on .name:after,
.my-profile .toolbar .on:after,
.btn-floating .on:before {content:' '; position:absolute; width:0.43rem; height:0.43rem; border-radius:50%; background-color:#ff3131;}
@media all and (max-width:360px){
	.nav-gnb .on:after,
	.category-menu .on .name:after,
	.my-profile .toolbar .on:after,
	.btn-floating .on:before {width:4px; height:4px;}
}

/* fixed css */
.fixed-top {position:fixed; top:0; left:0; z-index:10;}
.fixed-bottom {position:fixed; bottom:0; left:0;}

/* button */
/*.btn {text-align:center;}
.btn-red {background-color:#ff3131; color:#fff;}
.btn-block {display:block;}
.btn-large {height:3.41rem; padding-top:0.09rem; font:1.37rem/3.41rem 'AvenirNext-Medium', 'ppleSDGothicNeo-Medium';}*/
.btn-plus {display:block; width:100%; height:4.27rem; font:1.19rem/4.1rem 'AppleSDGothicNeo-Medium'; text-align:center;}

.btn {display:inline-block; min-width:4.95rem; padding:0 1.02rem; border:1px solid rgba(0, 0, 0, 0.15); color:#4a4a4a; font-family:'AvenirNext-Regular', 'AppleSDGothicNeo-Regular', 'RobotoRegular', 'Noto Sans', sans-serif; text-align:center;}
button.btn {background-color:transparent;}

/* button size */
.btn-xsmall {height:2.3rem; font:1.02rem/2.22rem 'AvenirNext-Regular', 'AppleSDGothicNeo-Light';}
.btn-small {height:2.82rem; font:1.02rem/2.82rem 'AvenirNext-Medium', 'AppleSDGothicNeo-Medium';}
.btn-default {height:3.07rem; font-size:1.11rem; line-height:3.07rem;}
a.btn.btn-default.btn-red,
a.btn.btn-default.btn-green,
a.btn.btn-default.btn-blue,
a.btn.btn-default.btn-grey,
a.btn.btn-default.btn-black {padding-top:0.17rem; line-height:2.99rem;}
.btn-large {height:3.41rem; padding-top:0.09rem; font-size:1.19rem; line-height:3.33rem;}
.btn-xlarge {height:3.75rem; padding-top:0.09rem; font:bold 1.37rem/3.67rem 'AvenirNext-DemiBold', 'AppleSDGothicNeo-SemiBold';}

/* button outline style */
.btn-line-grey, .btn-line-red, .btn-line-green, .btn-line-blue {border-style:solid; border-width:1px;}
.btn-line-grey {border-color:#a9a9a9; color:#a9a9a9;}
.btn-line-red {border-color:#ff3131; color:#ff3131;}
.btn-line-green {border-color:#00a061; color:#00a061;}
.btn-line-blue {border-color:#3f75ff; color:#3f75ff;}

/* button color */
.btn-grey, .btn-red, .btn-green, .btn-blue, .btn-black {border:0; color:#fff;}
.btn-red {background-color:#ff3131;}
.btn-green {background-color:#00a061;}
.btn-blue {background-color:#3f75ff;}
.btn-black {background-color:#262626;}
.btn-grey {background-color:#a9a9a9;}

/* button radius */
.btn-radius {border-radius:0.17rem;}

/* button align */
.btn-align-right {text-align:right;}

/* block level button */
.btn-block {display:block; width:100%;}

/* color */
.color-red {color:#ff3131;}
.color-green {color:#00a061;}
.color-blue {color:#3f75ff;}
.color-black {color:#0d0d0d;}
.color-grey {color:#6e6e6e;}

.bg-grey {background-color:#f4f4f4;}
.bg-black {background-color:#262626;}
.bg-red {background-color:#ff3131;}
.bg-green {background-color:#00a061;}

/* icon */
.icon {display:inline-block;}
.icon-download {width:1.71rem; height:1.71rem; margin-left:0.43rem; background-position:-5.5rem -29.38rem; vertical-align:top;}
.icon-plus {display:inline-block; position:relative; width:0.85rem; height:0.85rem; margin-right:0.43rem; margin-top:-0.045rem;}
.icon-plus:after, .icon-plus:before {content:' '; position:absolute; top:50%; left:50%; width:100%; height:1px; margin:-0.5px 0 0 -50%;}
.icon-plus:before {-webkit-transform:rotate(-90deg); transform:rotate(-90deg);}
.icon-plus-blue:after, .icon-plus-blue:before {background-color:#3f75ff;}
.icon-plus-white:after, .icon-plus-white:before {background-color:#fff;}
.icon-facebook:after {background-position:0 -13.35rem;}
.icon-twitter:after {background-position:-4.82rem -13.35rem;}
.icon-kakao:after {background-position:-9.64rem -13.35rem;}
.icon-pinterest:after {background-position:0 -18.18rem;}
.icon-url:after {background-position:-4.82rem -18.18rem;}
.icon-line:after {background-position:0 -22.99rem;}

/* bg image sprite */
.sprite,
.list-card .icon-culture {background-image:url(http://fiximage.10x10.co.kr/m/2017/common/bg_sp.png?v=1.60); background-repeat:no-repeat; background-size:17.07rem auto;}

/* bg image sprite for layout */
.tenten-header .tenten,
.tenten-header .btn-close,
.btn-search,
.tenten-header .btn-shoppingbag,
.tenten-header .new:after,
.btn-back,
.btn-home,
.tenten-header .btn-share,
.tenten-sns a,
.btn-top:after,
.tab-bar a:after,
.btn-zoom:after,
.popWin .header .btnPopClose .pButton,
.sns-list .icon:after {background-image:url(http://fiximage.10x10.co.kr/m/2017/common/bg_sp_layout.png?v=1.67); background-repeat:no-repeat; background-size:17.07rem auto;}

/* bg image sprite for arrow */
.footer-content .tenten a:after,
.list-card .culture-bnr .subcopy:after,
.pagingV15a .arrow a:after {background-image:url(http://fiximage.10x10.co.kr/m/2017/common/bg_sp_arrow.png?v=1.79); background-repeat:no-repeat; background-size:17.07rem auto;}

/* label */
.label.label-line {border-bottom:1px solid rgba(54, 147, 219, 0.7); color:#3693db; font:0.94rem/0.94rem 'AvenirNext-Medium', 'AppleSDGothicNeo-Medium'; vertical-align:top;}
.label-color {position:relative; margin-right:0.51rem; font-family:'AvenirNext-Medium', 'AppleSDGothicNeo-Medium';}
.label-color:after {content:' '; display:inline-block; width:1px; height:0.94rem; margin-left:0.51rem; background-color:#e1e1e1; vertical-align:-0.045rem;}
.label-speech {position:relative; display:inline-block; height:2.05rem; padding:0.17rem 0.51rem 0; border-radius:2.05rem; background-color:#ff3131; color:#fff; font:italic bold 1.11rem/1.96rem 'AvenirNext-DemiBold', 'AppleSDGothicNeo-SemiBold'; letter-spacing:-0.0426rem;}
.label-speech:after {content:' '; position:absolute; bottom:-0.26rem; left:0; width:0; height:0; border-style:solid; border-width:1.37rem 1.37rem 0 0; border-color:#ff3131 transparent transparent transparent;}
.label-speech b {position:relative; z-index:5;}
.label-speech b:nth-child(2):before {content:' '; display:inline-block; width:1px; height:0.68rem; margin:0 0.26rem 0 0.17rem; background-color:rgba(255, 255, 255, 0.55); vertical-align:0.09rem;}
.label-speech i {font-style:normal;}
.label-circle {display:inline-block; width:3.93rem; height:3.93rem; padding-top:0.09rem; border-radius:50%; background-color:#000; color:#fff; font:1.02rem/3.93rem 'AppleSDGothicNeo-Medium'; text-align:center;}
.label-star {width:4.05rem; height:4.05rem; color:#fff; font:italic 0.94rem/4.05rem 'AvenirNext-Bold'; font-weight:bold; text-align:center;}
.label-star em {position:absolute; top:0; left:0; width:100%; height:100%; transform:rotate(-25deg);}

/* pagination */
.pagination-dot {position:absolute; bottom:0.85rem; left:0; z-index:5; width:100%;}
.pagination-dot .swiper-pagination-switch {display:inline-block; width:0.6rem; height:0.6rem; margin:0 0.26rem; border-radius:50%;}
@media all and (max-width:360px){
	.pagination-dot .swiper-pagination-switch {width:6px; height:6px;}
}
.block-dot {text-align:right;}
.block-dot .swiper-pagination-switch {border:1px solid #000; background-color:transparent;}
.block-dot .swiper-active-switch {background-color:#000;}

/* loading image */
.default-font .thumbnail {position:relative; background-color:#f4f4f4;}
.default-font .thumbnail img {position:relative; z-index:2;}
.default-font .thumbnail:before {content:' '; position:absolute; top:50%; left:50%; width:4.27rem; height:4.27rem; margin:-2.22rem 0 0 -2.22rem; background:url(http://fiximage.10x10.co.kr/m/2017/common/bg_img_loading.png) 50% 0 no-repeat; background-size:100% auto;}

/* sns */
.sns-list .icon {display:block; position:relative; border-radius:50%;}
.sns-list .icon:after {content:' '; position:absolute; top:0; left:0; width:100%; height:100%;}

/* item list */
.items {overflow:hidden; background-color:#fff;}
.items:after {content:' '; display:block; clear:both;}
.items li {position:relative;}
.items li > a {display:block;}
.items .desc {position:relative;}
.items .thumbnail {overflow:hidden; position:relative;}
.items .soldout,
.items .myview {position:absolute; top:0; left:0; z-index:5; width:100%; height:100%; padding-top:43.5%; background-color:rgba(0, 0, 0, 0.4); color:#fff; font:1.54rem 'AppleSDGothicNeo-Medium'; text-align:center;}
.items .myview {font-size:1.28rem;}
.items .no {display:block; margin:-0.26rem 0 0.94rem; color:#434343; font-size:2.05rem;}
.items .brand {color:#b0b0b0; font:1.02rem/1.02rem 'AvenirNext-Medium', 'AppleSDGothicNeo-Medium';}
.items .name {overflow:hidden; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; height:2.82rem; margin-top:0.77rem; font:1.11rem/1.45rem 'AvenirNext-Regular', 'AppleSDGothicNeo-Light'; text-overflow:ellipsis;}
.items .price {margin-top:1.19rem;}
.items .price b {font-family:'AvenirNext-DemiBold'; font-weight:bold;}
.items .won {font-family:'AvenirNext-Regular', 'AppleSDGothicNeo-Light'; font-size:1.02rem; font-weight:normal;}
.items .unit {margin-top:0.34rem;}
.items .unit:first-child {margin-top:0;}
.items .sum {color:#2c2c2c; font-size:1.28rem;}
.items .discount {font-size:1.11rem; line-height:1.11rem;}
.items .discount small {font:0.94rem 'AppleSDGothicNeo-Medium';}
.items .red {color:#ff3131;}
.items .green {color:#00a061;}
.items .etc {position:absolute;}
.items .etc:after {content:' '; display:block; clear:both;}
.items .etc,
.items .etc button {color:#8c8c8c; font-size:0.85rem;}
.items .btn-view {overflow:hidden; position:absolute; width:2.39rem; height:2.65rem; color:transparent;}
.items .btn-view:before {content:' '; position:absolute; top:50%; left:50%; width:1.37rem; height:1.37rem; margin:-0.68rem 0 0 -0.68rem; background-position:-11.05rem -10.7rem;}
.items .btn-view:after {display:none;}
.items .tag {float:left; height:1.02rem; margin-right:0.68rem; line-height:1.28rem;}
.items .icon {position:relative; height:1.02rem; vertical-align:top;}
.items .icon:before {content:' '; position:absolute; top:0; left:0; width:1.02rem; height:1.02rem;}
.items .icon-wish:before {margin-top:0.17rem; background-position:0 -10.7rem;}
.items .btn-wish {font-size:0.85rem; background-color:transparent;}
.items .btn-wish .on:before {background-position:-1.24rem -10.7rem;}
.items .icon-wish {width:1.02rem;}
.items .icon-wish:before {top:50%; margin-top:-0.43rem;}
.items .icon-wish i {padding-left:1.19rem; font-size:1.02rem; line-height:1.19rem;}
.items .counting {padding-left:0.51rem;}
.items .btn-wish .counting {padding-left:0.17rem;}
.items .icon-rating {display:inline-block; width:4.1rem;}
.items .icon-rating:before {width:4.35rem; background-position:-12.63rem -10.7rem;}
.items .icon-rating i {position:absolute; top:0; left:0; height:1.2rem; background-position:-12.63rem -11.94rem; text-indent:-999em;}
.items .icon-shipping {width:1.37rem;}
.items .icon-shipping:before {top:50%; width:1.37rem; height:1.37rem; margin-top:-0.68rem; background-position:-6.31rem -10.7rem;}
.items .icon-shipping i {display:none;}
.items .btn-compare-add {display:none; position:absolute; top:0; left:0; z-index:5; text-indent:-999em; background-color:transparent;}
.items .btn-compare-add:after {content:' '; position:absolute; width:3.41rem; height:3.41rem; border-radius:50%; background-color:#8d8d8d; background-position:-3.67rem -5.93rem;}
.items .btn-compare-add.on:after {background-color:#ff4646;}
.type-list .label, .type-grid .label, .type-card .label {display:inline-block; padding:0.09rem 0.51rem 0; border:1px solid rgba(0, 0, 0, 0.7); font-size:0.94rem;}
.type-list .label, .type-grid .label, .type-card .label {height:1.54rem; padding-top:0.09rem; border-radius:1.54rem; line-height:1.45rem;}

.type-column {overflow:hidden; width:29.7rem; margin:2.73rem auto 0;}
.type-column ul {margin-top:-2.05rem;}
.type-column-2 li {width:14.59rem;}
.type-column-3 li {width:9.64rem;}
.type-column li {float:left; margin:2.05rem 0.13rem 0; text-align:center;}
.type-column-2 .thumbnail,
.type-column-2 .thumbnail img {width:14.59rem; height:14.59rem;}
.type-column-3 .thumbnail,
.type-column-3 .thumbnail img {width:9.64rem; height:9.64rem;}
.type-column .name {display:block; height:1.37rem; margin-top:1.11rem; padding:0 0.43rem; white-space:nowrap;}
.type-column .price {margin-top:0.26rem;}
.type-column .sum {font-size:1.11rem;}
.type-column-3 .name {margin-top:0.85rem;}
.type-photowall ul {overflow:hidden; width:29.7rem; margin:0 auto;}
.type-photowall li {float:left; margin:0.26rem 0.13rem 0;}
.type-photowall li,
.type-photowall .thumbnail {position:relative;}
.type-photowall .thumbnail,
.type-photowall .thumbnail img {width:9.64rem; height:9.64rem;}
.type-photowall .thumbnail:after {content:' '; position:absolute; top:0; left:0; z-index:10; width:100%; height:100%; background-color:rgba(0, 0, 0, 0.03);}
.type-photowall li:first-child,
.type-photowall li:first-child .thumbnail,
.type-photowall li:first-child .thumbnail img {width:19.54rem; height:19.54rem;}
.type-photowall li:nth-child(4) {clear:left;}

.type-full {width:29.44rem; margin:2.56rem auto 0; text-align:center;}
.type-full li {margin-top:3.75rem;}
.type-full li:first-child {margin-top:0;}
.type-full .thumbnail,
.type-full .thumbnail img {width:29.44rem; height:29.44rem;}
.type-full .thumbnail:after {content:' '; position:absolute; top:0; left:0; z-index:15; width:100%; height:100%; background-color:rgba(0, 0, 0, 0.03);}
.type-full .brand {display:block; margin-top:2.3rem; color:rgba(0, 0, 0, 0.4); font-size:1.19rem;}
.type-full .name {margin-top:1.11rem; height:1.79rem; padding-top:0.17rem; font-size:1.37rem; line-height:1.71rem;}
.type-full .price {margin-top:0.77rem;}
.type-full .sum {font-size:1.37rem;}
.type-full .discount {font-size:1.28rem;}

.type-list li > a:after, .type-list li:last-child:after {content:' '; position:absolute; width:100%; height:1px; background-color:rgba(0, 0, 0, 0.07);}
.type-list li > a:after {top:0; left:0; z-index:5; width:100%; height:1px; background-color:rgba(0, 0, 0, 0.07);}
.type-list li:last-child:after {bottom:0; left:0; z-index:5;}
.type-list li > a {overflow:hidden;}
.type-list .thumbnail {float:left; width:13.99rem; height:13.99rem;}
.type-list .thumbnail img {width:100%; height:100%;}
.type-list .desc {margin-left:13.99em; padding:1.62rem 2.99rem 1.17rem 1.19rem;}
.type-list .etc {bottom:1.45rem; left:15.19rem;}
.type-list .btn-view {top:0; right:0;}
.type-list .btn-compare-add {width:13.99rem; height:13.99rem;}
.type-list .btn-compare-add:after {bottom:0.85rem; left:9.73rem;}

.type-grid {position:relative; border-top:1px solid rgba(0, 0, 0, 0.07);}
.type-grid:after {display:none;}
.type-grid ul {width:32rem; margin:0 auto; padding:0 0.43rem;}
.type-grid ul:after {content:' '; display:block; clear:both;}
.type-grid li {float:left; width:50%; padding:0.85rem 0.43rem 0;}
.type-grid .thumbnail {width:14.68rem; height:14.68rem;}
.type-grid .desc {height:11.52rem; padding:1.11rem 0 0;}
.type-grid .name {margin-top:0.68rem; padding-right:1.96rem;}
.type-grid .price {margin-top:1.11rem;}
.type-grid .unit {display:inline;}
.type-grid .price .unit:nth-child(2) .sum {display:none;}
.type-grid .etc {bottom:1.37rem; left:0.43rem;}
.type-grid .tag.shipping {display:none;}
.type-grid .btn-view {right:0; bottom:8.62rem;}
.type-grid .btn-compare-add {width:100%; height:58%;}
.type-grid .btn-compare-add:after {right:1.02rem; bottom:0.51rem;}

.type-big li {padding-bottom:3.93rem;}
.type-big .thumbnail {margin:0 auto;}
.type-big .thumbnail,
.type-big .thumbnail img {width:28.07rem;}
.type-big .desc {width:28.07rem; margin:0 auto; padding:1.96rem 0 2.3rem; text-align:center;}
.type-big .name {margin-top:1.11rem;}
.type-big .price {margin-top:1.11rem;}
.type-big .price .unit:nth-child(2) .sum {display:none;}
.type-big .unit {display:inline;}
.type-big .etc {width:28.07rem; bottom:3.93rem; left:50%; margin-left:-14.035rem; text-align:center;}
.type-big .tag {float:none; display:inline-block;}

.type-card li {overflow:hidden; border-radius:0.26rem;}
.type-card .desc {height:6.83rem; padding:1.02rem 0.94rem 1.11rem; border-bottom-right-radius:0.26rem; border-bottom-left-radius:0.26rem; background-color:#fff;}
.type-card .price {margin-top:0.17rem;}
.type-card .sum {font-size:1.11rem;}
.type-card .price .discount {font-size:1.02rem;}

.type-box-grey .swiper-slide {width:11.61rem; background-color:rgba(0, 0, 0, 0.03);}
.type-box-grey .thumbnail {width:11.61rem; height:11.61rem;}
.type-box-grey .thumbnail img {width:100%; height:100%;}
.type-box-grey .thumbnail:after {content:' '; position:absolute; top:0; left:0; z-index:15; width:100%; height:100%; background-color:rgba(0, 0, 0, 0.03);}
.type-box-grey .desc {height:8.7rem; padding:1.19rem 1.19rem 0;}
.type-box-grey .desc .name {margin-top:0; color:#000;}
.type-box-grey .price {margin-top:0.09rem;}
.type-box-grey.items .price .sum,
.type-box-grey.items .price .won {color:#000;}
.type-box-grey .name + .price {margin-top:0.68rem; color:#838383;}
.type-box-grey.items .name + .price .sum,
.type-box-grey.items .name + .price .won {color:#838383;}
.type-box-grey .price + .name {margin-top:0.85rem; color:#838383;}

.category-item-list {margin-top:1.71rem; padding:2.56rem 0 0.77rem; background-color:#fff;}
.category-item-list > .headline {width:28.41rem; height:2.99rem; margin:0 auto; font:1.37rem/2.99rem 'AppleSDGothicNeo-Medium';}
.category-item-list .headline .icon {width:2.99rem; height:2.99rem; margin-left:-0.85rem; vertical-align:bottom;}
.category-item-list .items ul {overflow:hidden; width:29.7rem; margin:-1.11rem auto 0;}
.category-item-list .items li {float:left; width:8.7rem; margin:2.56rem 0.6rem 0;}
.category-item-list .items a {height:15.02rem;}
.category-item-list .items .thumbnail {width:8.7rem; height:8.7rem;}
.category-item-list .items .thumbnail img {width:100%; height:100%;}
.category-item-list .exhibition-list {margin-top:3.33rem;}

/* event list card type */
.list-card .desc {padding:0 2.05rem;}
.list-card .thumbnail img {min-height:17.32rem;}
.list-card .ad-bnr {margin-bottom:0.85rem;}
.list-card .ad-bnr .thumbnail img {min-height:6.14rem;}
.list-card .headline, .list-card .subcopy {display:block;}
.list-card .ellipsis {overflow:hidden; display:inline-block; text-overflow:ellipsis; white-space:nowrap; vertical-align:middle;}
.list-card .subcopy {font-size:1.19rem}
.list-card .culture-bnr {background-color:#095f8e; text-align:center;}
.list-card .culture-bnr a {display:block; height:5.8rem; padding-top:0.26rem; line-height:5.8rem;}
.list-card .culture-bnr b {color:#fff; font:1.37rem/1.37rem 'AppleSDGothicNeo-SemiBold';}
.list-card .culture-bnr .subcopy {display:inline-block; margin:-0.51rem 0 0 0.51rem; color:#c2e7ff; font-size:1.11rem;}
.list-card .culture-bnr .subcopy:after {content:' '; display:inline-block; width:0.85rem; height:1.11rem; margin-left:0.17rem; background-position:-0.73rem -1.54rem; vertical-align:-0.17rem;}
.list-card .icon-culture {width:2.56rem; height:2.56rem; margin-right:0.35rem; background-position:-7.42rem -29.38rem; vertical-align:-0.6rem;}

.type-align-center li {margin-top:2.99rem;}
.type-align-center li:frist-child {margin-top:0;}
.type-align-center .desc {padding-top:1.54rem; text-align:center;}
.type-align-center .headline {color:#0d0d0d; font:1.54rem/1.88rem 'AvenirNext-Medium', 'AppleSDGothicNeo-Medium';}
.type-align-center .subcopy {margin-top:0.51rem; color:#2c2c2c; font-family:'AvenirNext-Regular', 'AppleSDGothicNeo-Light';}
.type-align-center .ellipsis {max-width:76.15%;}
.type-align-center .discount {font:bold 1.19rem/1.19rem 'AvenirNext-DemiBold', 'AppleSDGothicNeo-SemiBold'; vertical-align:middle;}
.type-align-center .discount small {font:normal 1.02rem 'AppleSDGothicNeo-Medium';}

.type-align-left .desc {padding-top:1.37rem; padding-bottom:2.56rem;}
.type-align-left .headline {min-height:2.22rem; font:bold 1.62rem/1.96rem 'AvenirNext-DemiBold', 'AppleSDGothicNeo-Bold';}
.type-align-left .ellipsis {width:80%;}
.type-align-left .ellipsis.full {width:100%;}
.type-align-left .subcopy {margin-top:0.34rem; color:#838383; line-height:1.62rem;}
.type-align-left .discount {display:inline-block; width:4.95rem; font:bold 1.71rem/1.71rem 'AvenirNext-DemiBold'; vertical-align:middle; text-align:right;}


.body-main {padding-top:0 !important;}

/* button */
.btn-arrow {overflow:hidden; position:relative; width:2.39rem; height:2.39rem; text-indent:-999em;}
.btn-arrow:after {content:' '; position:absolute; top:50%; left:50%; width:1.28rem; height:1.54rem; margin:-0.77rem 0 0 -0.64rem; background-position:-15.45rem 0;}

/* label */
.label-triangle {width:4.01rem; height:4.01rem; color:#000;}
.label-triangle:after {content:' '; position:absolute; bottom:0; right:0; width:0; height:0; border-style:solid; border-width:0 0 4.01rem 4.01rem; border-color:transparent transparent #fff transparent;}
.label-triangle em {display:block; position:relative; z-index:5; width:100%; margin-left:-0.17rem; padding-top:2.65rem; font:0.85rem/0.85rem 'AvenirNext-Medium', sans-serif-medium; letter-spacing:-0.017rem; text-align:center; transform:rotate(-45deg); -webkit-transform:rotate(-45deg);}

/* heading */
.hgroup {position:relative;}
.hgroup .btn-more {position:absolute; top:0; right:1.28rem;}
.headline-speech {position:relative; min-height:2.39rem; padding:0 2.05rem; font-size:1.28rem; line-height:2.39rem;}
.headline-speech span, .headline-speech small {vertical-align:middle;}
.headline-speech span {display:inline-block; position:relative; height:2.39rem; margin-top:-0.17rem; margin-right:0.43rem; padding:0.17rem 0.94rem 0 1.11rem; border:1px solid #000; border-radius:2.22rem; color:#000; line-height:2.13rem; text-align:center;}
.headline-speech span:lang(ko) {font:bold 1.02rem/2.13rem 'AppleSDGothicNeo-Bold';}
.headline-speech span:lang(en) {padding-top:0.09rem; font:italic bold 0.94rem/2.22rem 'AvenirNext-DemiBold';}
.headline-speech span:after {content:' '; position:absolute; left:-1px; bottom:-0.32rem; width:0.47rem; height:1.51rem; background:url(http://fiximage.10x10.co.kr/m/2017/common/bg_tail.png) 50% 100% no-repeat; background-size:100% auto; }
.headline-speech small {color:#838383; font-size:1.28rem;}

/* bg image sprite */
.menu-etc .icon,
.playing-bnr > a:after,
.hitchhiker-bnr > a:after,
.shipping-bnr .icon-car {background-image:url(http://fiximage.10x10.co.kr/m/2017/common/bg_sp_today_categorymain.png?v=1.60); background-repeat:no-repeat; background-size:17.07rem auto;}

/* bg image sprite - icon categroy */
.menu-category li .icon,
.category-item-list .headline .icon {background-image:url(http://fiximage.10x10.co.kr/m/2017/common/bg_sp_icon_category.png?v=1.54); background-repeat:no-repeat; background-size:17.07rem auto;}

/* bg image sprite - arrow */
.menu-category .icon-arrow:after,
.icon-up:after,
.btn-arrow:after,
.brand-items .btn-more span,
.shipping-bnr em:after,
.time-sale .icon-arrow {background-image:url(http://fiximage.10x10.co.kr/m/2017/common/bg_sp_arrow.png?v=1.56); background-repeat:no-repeat; background-size:17.07rem auto;}

/* items list */
.items .name {height:3.07rem; color:#4a4a4a; font-size:1.19rem; line-height:1.54rem;}
.items .name, .items .price {font-family:'AvenirNext-Regular', 'AppleSDGothicNeo-Regular';}
.items .price b {font-family:'AppleSDGothicNeo-Regular'; font-weight:normal;}
.items .price .discount {font-family:'AvenirNext-Medium', sans-serif-medium;}
.items .price .sum,
.items .price .won {color:#838383; font:1.11rem/1.11rem 'AvenirNext-Regular', 'AppleSDGothicNeo-Regular';}

/* banner type text */
.text-bnr {position:relative; padding-bottom:5.8rem;}
.text-bnr .desc {position:absolute; bottom:0; left:0; z-index:10; width:100%; padding:0 2.05rem;}
.text-bnr .headline,
.text-bnr .subcopy {display:block;}
.text-bnr .headline {margin-top:1.11rem; padding-right:4.27rem; font:bold 2.39rem/3.07rem 'AvenirNext-Bold', 'AppleSDGothicNeo-Bold';}
.text-bnr .subcopy {overflow:hidden; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; height:3.24rem; margin-top:1.19rem; color:#4a4a4a; font-size:1.19rem; line-height:1.71rem; word-break:break-all;}
.text-bnr .thumbnail img {min-height:32rem;}

/* main banner */
.main-bnr {padding-bottom:0;}
.main-bnr a {display:block; position:relative; padding-bottom:5.8rem;}
.main-bnr .pagination-dot {bottom:7.42rem; left:auto; right:1.79rem;}
.main-bnr .thumbnail img {min-height:41.3rem;}

/* juest one day */
.time-sale .headline {font:bold 1.54rem/1.54rem 'AvenirNext-DemiBold'; letter-spacing:0.256rem;}
.time-sale .weekday,
.time-sale .bnr {margin-top:5.21rem;}
.time-sale .weekday a {display:block; position:relative; height:11.95rem; padding:0 2.05rem;}
.time-sale .weekday .thumbnail {position:absolute; top:0; left:2.05rem; width:11.95rem; height:11.95rem;}
.time-sale .weekday .thumbnail img {width:100%; height:100%;}
.time-sale .weekday .thumbnail:after,
.time-sale .weekend .thumbnail:after {content:' '; position:absolute; top:0; left:0; z-index:15; width:100%; height:100%; background-color:rgba(0, 0, 0, 0.03);}
.time-sale .desc {height:10.75rem; padding:0.85rem 0 0 14.42rem;}
.time-sale .name {display:block; margin-top:1.28rem; height:1.54rem; color:#000; white-space:nowrap;}
.time-sale .price {margin-top:0; color:#838383; font-size:1.37rem; letter-spacing:-0.034rem;}
.time-sale .price .sum,
.time-sale .price .won {color:#000; font:bold 1.37rem 'AvenirNext-DemiBold', 'AppleSDGothicNeo-SemiBold';}
.time-sale .discount {display:block; margin-top:0.34rem; font:bold 4.1rem/4.1rem 'AvenirNext-DemiBold';}
.time-sale .time-line {position:relative; margin:0.77rem 2.05rem 0; padding-bottom:0.68rem; border-bottom:2px solid #ececec;}
.time-sale .time {display:inline-block; position:relative; width:8.53rem; height:2.05rem; padding-top:0.0426rem; border-radius:2.05rem; background-color:#000; color:#fff; font:bold 1.02rem/2.13rem 'AvenirNext-DemiBold', 'AppleSDGothicNeo-SemiBold'; text-align:center;}
.time-sale .time-line:after, .time-sale .time-line:before {content:' '; position:absolute;}
.time-sale .time-line:after {bottom:-0.43rem; width:0.34rem; height:0.34rem; border:2px solid #000; border-radius:50%; background-color:#fff; transition:left 0.3s;}
@media all and (min-width:360px) and (max-width:360px){
	.time-sale .time-line:after {width:4px; height:4px; border:2px solid #000;}
}
.time-sale .time-line:before {bottom:-2px; height:0.17rem; background-color:#000; transition:width 0.3s;}
.time-sale .interval-1:after {left:3.07%;}
.time-sale .interval-1:before {width:3.07%;}
.time-sale .interval-2:after {left:12.27%;}
.time-sale .interval-2:before {width:12.27%;}
.time-sale .interval-3:after {left:21.47%;}
.time-sale .interval-3:before {width:21.47%;}
.time-sale .interval-3 .time {margin-left:8%;}
.time-sale .interval-4 .time {margin-left:16.34%;}
.time-sale .interval-4:after {left:30.67%;}
.time-sale .interval-4:before {width:30.67%;}
.time-sale .interval-5 .time {margin-left:26.2%;}
.time-sale .interval-5:after {left:40.49%;}
.time-sale .interval-5:before {width:40.49%;}
.time-sale .interval-6 .time {margin-left:36%;}
.time-sale .interval-6:after {left:50%;}
.time-sale .interval-6:before {width:50%;}
.time-sale .interval-7 .time {margin-left:45.5%;}
.time-sale .interval-7:after {left:59.51%;}
.time-sale .interval-7:before {width:59.51%;}
.time-sale .interval-8 .time {margin-left:55%;}
.time-sale .interval-8:after {left:69.02%;}
.time-sale .interval-8:before {width:69.02%;}
.time-sale .interval-9 .time {margin-left:64.5%;}
.time-sale .interval-9:after {left:78.52%;}
.time-sale .interval-9:before {width:78.52%;}
.time-sale .interval-10 .time, .time-sale .interval-11 .time {margin-left:70%;}
.time-sale .interval-10:after {left:88.03%;}
.time-sale .interval-10:before {width:88.03%;}
.time-sale .interval-11:after {left:97.55%;}
.time-sale .interval-11:before {width:97.55%;}
.time-sale .weekend {margin-top:5.97rem; margin-bottom:-0.77rem;}
.time-sale .weekend > a {display:block; position:relative; padding-bottom:7.85rem; text-align:center;}
.time-sale .weekend .headline {position:absolute; bottom:0; left:0; width:100%; font-family:'AvenirNext-Bold', 'AppleSDGothicNeo-Bold'; font-weight:bold;}
.time-sale .weekend .copy {display:block; height:1.96rem; line-height:1.96rem; letter-spacing:0.26rem;}
.time-sale .weekend .ellipsis {overflow:hidden; display:inline-block; max-width:22.99rem; text-overflow:ellipsis; white-space:nowrap;}
.time-sale .icon-arrow {width:0.51rem; height:1.02rem; margin-top:0.26rem; margin-left:0.34rem; border:0; background-position:0 -1.54rem; vertical-align:top;}
@media all and (min-width:360px) and (max-width:360px){
	.time-sale .icon-arrow {margin-top:0.43rem;}
}
.time-sale .weekend .discount {margin-top:-0.09rem; font-size:3.58rem; letter-spacing:0.26rem;}
.time-sale .weekend .thumbnail {position:relative; width:8.7rem; height:8.7rem;}
.time-sale .weekend .thumbnail img {width:100%; height:100%;}
.time-sale .weekend .items {overflow:hidden; width:28.92rem; margin:0 auto;}
.time-sale .weekend .items li {float:left; margin:0 0.47rem;}

.marketing-bnr {margin-top:4.01rem;}
.marketing-bnr .thumbnail img {min-height:6.1rem;}

.keyword-raking {margin-top:3.75rem;}
.keyword-raking + .exhibition-list {margin-top:2.3rem;}
.keyword-raking ul {column-count:2; -webkit-column-count:2; column-gap:0; -webkit-column-gap:0; margin-top:1.45rem; padding-top:1.28rem; border-top:1px solid #efefef;}
.keyword-raking li {padding:0.26rem 0;}
.keyword-raking li a {overflow:hidden; display:block; position:relative; height:2.56rem; padding:0 1.02rem 0 2.13rem; color:#000; font:1.28rem/2.56rem 'AvenirNext-Regular ', 'AppleSDGothicNeo-Light';}
.keyword-raking li:nth-child(n+6) a {padding:0 2.22rem 0 1.19rem;}
.keyword-raking .no {display:inline-block; width:1.79rem; height:1.96rem; font:italic bold 1.37rem/1.96rem 'AvenirNext-DemiBold'; vertical-align:middle;}
.keyword-raking li a .icon {position:absolute; top:50%; right:0; z-index:5; width:2.22rem; height:1.19rem; margin-top:-0.68rem; color:#ff3030; font:italic 0.85rem/1.19rem 'AvenirNext-Medium';}
@media all and (min-width:360px) and (max-width:360px){
	.keyword-raking li a .icon {margin-top:-0.6rem;}
}
.keyword-raking li:nth-child(n+6) .icon {right:2.05rem;}
.keyword-raking .icon-up {text-indent:-999em;}
.keyword-raking .icon-up:after {content:' '; position:absolute; top:0; left:50%; width:1.19rem; height:1.19rem; margin-left:-0.6rem; background-position:-14.03rem 0;}
.keyword-raking .keyword {overflow:hidden; display:inline-block; width:8.53rem; height:1.96rem; line-height:1.96rem; text-overflow:ellipsis; white-space:nowrap; vertical-align:middle;}
.keyword-raking .btn-plus {color:#3f75ff;}
.keyword-raking .btn-group {margin-top:0.68rem;}

.gif-bnr {margin-top:2.21rem;}
.gif-bnr .thumbnail {width:27.82rem; margin:0 auto;}

.hot-keyword {position:relative; padding:2.99rem 0 7.77rem; background-color:#f4f4f4;}
.hot-keyword .headline {position:absolute; bottom:2.65rem; left:0; width:100%; padding:0 2.65rem;}
.hot-keyword .headline span {font:1.37rem/1.79rem 'AvenirNext-Medium'; letter-spacing:0.09rem;}
.hot-keyword .headline em, .hot-keyword .headline .vol {font-family:'AvenirNext-Bold'; font-weight:bold;}
.hot-keyword .headline em {display:block;}
.hot-keyword .headline .vol {font-size:1.19rem;}
.hot-keyword .headline small {position:absolute; bottom:0; right:2.65rem; font:bold 2.39rem/3.33rem 'AvenirNext-Bold', 'AppleSDGothicNeo-Bold'; text-decoration:underline;}
.hot-keyword ul {overflow:hidden; width:28rem; margin:-1.02rem auto 0;}
.hot-keyword li {float:left; position:relative; width:11.78rem; margin:1.02rem 1.11rem 0;}
.hot-keyword .thumbnail {width:11.78rem; height:11.78rem; border-radius:50%;}
.hot-keyword .thumbnail img {width:100%; height:100%; border-radius:50%;}
.hot-keyword .label {position:absolute; right:0; bottom:0; z-index:10; width:3.93rem; height:3.93rem; text-transform:capitalize;}

.md-pick {position:relative; margin-top:3.67rem;}
.md-pick .swiper-container {margin-top:1.96rem; padding:0 1.02rem 0 2.05rem;}
.md-pick .desc {height:6.6rem;}
.md-pick .items .name {height:auto;}
.md-pick .discount {display:block; margin-top:0.34rem; font-family:'AvenirNext-Medium', sans-serif-medium;}

.type-multi-row .swiper-slide {overflow:hidden; width:8.53rem; margin-right:1.02rem;}
.type-multi-row .swiper-slide a {height:15.18rem;}
.type-multi-row .swiper-slide .btn-more {display:block; width:8.53rem; height:8.53rem; padding:4.78rem 0.6rem 0 0; color:#000; font:2.3rem 'AvenirNext-Medium'; letter-spacing:0.0853rem; text-align:right;}
.type-multi-row .thumbnail {width:8.53rem; height:8.53rem;}
.type-multi-row .thumbnail img {width:100%; height:100%;}

.md-pick + .exhibition-list {margin-top:2.56rem;}

.items-single-bnr {margin-top:5.38rem;}
.items-single-bnr .items {overflow:visible;}
.items-single-bnr li {margin-top:0.85rem; padding:0 2.05rem;}
.items-single-bnr a {display:block; position:relative; height:12.63rem; padding:2.99rem 2.39rem 0 12.46rem; background-color:rgba(0, 0, 0, 0.03);}
.items-single-bnr .headline {font-size:1.45rem; line-height:2.13rem;}
.items-single-bnr .headline u {font-size:1.71rem;}
.items-single-bnr .price {display:block; margin-top:1.11rem;}
.items-single-bnr .thumbnail {position:absolute; background-color:transparent; width:15.36rem; height:13.65rem;}
.items-single-bnr .thumbnail:before {display:none;}
.items-single-bnr .thumbnail img {width:100%; height:100%;}
.items-single-bnr li:nth-child(1) a {text-align:right;}
.items-single-bnr li:nth-child(1) .thumbnail {top:-1.02rem; left:-1.02rem;}
.items-single-bnr li:nth-child(2) .thumbnail {right:-1.02rem; bottom:-1.02rem;}
.items-single-bnr li:nth-child(2) a {padding:3.41rem 12.46rem 0 2.39rem;}
.items-single-bnr u {display:block; font-family:'AvenirNext-DemiBold', 'AppleSDGothicNeo-SemiBold'; font-weight:bold;}
.items-single-bnr .price .sum,
.items-single-bnr .price .won {color:#000;}
.items-single-bnr .label {position:absolute; right:0; bottom:0;}

.new-items {margin-top:5.21rem;}
.on-sale-items {margin-top:3.75rem;}
.enjoy-items {margin-top:3.75rem;}
.new-items .swiper-container,
.on-sale-items .swiper-container,
.enjoy-items .swiper-container {margin-top:1.71rem; padding:0 2.05rem;}
.on-sale-items .items .name {font-family:'AvenirNext-Regular', 'AppleSDGothicNeo-Light';}

.headline-bubble + .type-box-grey {margin-top:1.71rem;}
.type-box-grey .swiper-slide {margin-right:0.68rem;}
.type-box-grey .swiper-slide:last-child {margin-right:0;}
.type-box-grey .label {position:absolute; bottom:0; right:0;}

.exhibition-plus-item {margin-top:3.75rem}
.exhibition-plus-item + .exhibition-list {margin-top:4.18rem;}
.exhibition-plus-item .list-card .desc {position:relative;}
.exhibition-plus-item .list-card .desc:after {content:' '; position:absolute; top:-0.85rem; left:2.05rem; z-index:15; width:0; height:0; border-style:solid; border-width:0 1.02rem 1.02rem 1.02rem; border-color:transparent transparent #fff transparent;}
.exhibition-plus-item .items {margin-top:-0.85rem;}
.exhibition-plus-item .items ul {overflow:hidden; width:28.68rem; margin:0 auto;}
.exhibition-plus-item .items li {float:left; margin:0 0.43rem;}
.exhibition-plus-item .items .thumbnail:after {content:' '; position:absolute; top:0; left:0; z-index:10; width:100%; height:100%; background-color:rgba(0, 0, 0, 0.03);}
.exhibition-plus-item .items .thumbnail {width:8.7rem; height:8.7rem;}
.exhibition-plus-item .items .thumbnail img {width:100%; height:100%;}
.exhibition-plus-item .items .price {margin-top:0.77rem; font-size:1.11rem;}

/* brand */
.brand-bnr {margin-top:4.18rem;}
.brand-bnr .thumbnail {width:23.21rem; height:28.33rem; margin:0 auto;}
.brand-bnr .thumbnail img {width:100%; height:100%;}
.brand-bnr .desc {position:relative; z-index:10; margin-top:-1.11rem; padding:0 2.22rem;}
.brand-bnr h2 {font:bold 2.3rem/2.3rem 'AvenirNext-Bold'; letter-spacing:0.085rem;}
.brand-bnr h3 {margin-top:1.19rem; font:bold 1.28rem/1.71rem 'AvenirNext-Bold', 'AppleSDGothicNeo-Bold';}
.brand-bnr .desc p {margin-top:0.085rem; color:#838383; font-size:1.19rem; line-height:1.71rem;}
.brand-items {position:relative; width:27.8rem; margin:1.62rem auto 0;}
.brand-items ul {overflow:hidden;}
.brand-items li {float:left; margin-right:0.85rem;}
.brand-items .thumbnail {display:block; width:8.7rem; height:8.7rem;}
.brand-items .btn-group {position:absolute; top:0; right:0;}
.brand-items .btn-more {position:absolute; top:0; left:0; z-index:15; width:100%; height:8.7rem; padding:0 0 0 0.26rem; color:#000; font:1.11rem 'AppleSDGothicNeo-Medium'; text-align:center; line-height:8.7rem;}
.brand-items .btn-more span {width:0.6rem; height:1.02rem; margin-left:0.34rem; border:0; background-position:0 -1.54rem; vertical-align:-0.001rem;}
.brand-items .icon-arrow:after {display:none;}

.brand-bnr + .text-bnr {margin-top:4.18rem;}

.text-bnr + .exhibition-list {margin-top:5.38rem;}

.menu-category {margin-top:1.19rem; background-color:#f7f7f7; padding-bottom:0.43rem;}
.menu-category ul {overflow:hidden;}
.menu-category li {float:left; width:25%; margin-top:2.22rem; height:5.03rem; padding-top:0.6rem; text-align:center;}
.menu-category li:nth-child(1),
.menu-category li:nth-child(2),
.menu-category li:nth-child(3),
.menu-category li:nth-child(4) {margin-top:1.62rem;}
.menu-category li .icon {display:block; width:2.9rem; height:2.9rem; margin:0 auto;}
.menu-category li a {display:block;}
.menu-category .name {display:inline-block; position:relative; margin-top:0.26rem; color:#6e6e6e; font:1.11rem 'AppleSDGothicNeo-Light';}
.menu-category .on .name:before {content:' '; position:absolute; top:-0.26rem; right:-0.51rem; width:0.43rem; height:0.43rem; border-radius:50%; background-color:#ff3131;}
@media all and (max-width:360px){
	.menu-category .on .name:before {width:4px; height:4px;}
}
.menu-category .btn-group {margin-top:1.45rem;}
.menu-category button {width:100%; height:4.27rem; background-color:transparent; color:#3f75ff; font:1.28rem/4.1rem 'AppleSDGothicNeo-Medium'; text-align:center;}
.menu-category .icon-arrow {position:relative; width:1.19rem; height:100%; margin-left:0.77rem; text-indent:-999em;}
.menu-category .icon-arrow:after {content:' '; position:absolute; top:50%; left:0; width:1.19rem; height:1.19rem; margin-top:-0.85rem; background-position:-12.63rem 0; vertical-align:middle; -webkit-transform:rotate(180deg); transform:rotate(180deg); transition:transform 0.3s; transition:-webkit-transform 0.3s; vertical-align:middle;}
.menu-category .btn-close {display:none;}
.menu-category .btn-close .icon-arrow:after {-webkit-transform:rotate(0deg); transform:rotate(0deg);}

/* gift guide banner */
.gift-guide-bnr {height:6.61rem; margin-top:0.17rem; padding-top:0.51rem;}
.gift-guide-bnr a {display:block; position:relative; height:6.1rem; padding:1.4rem 0 0 55.2%; background-color:#cf2e28;}
.gift-guide-bnr span {position:relative; z-index:5; color:#fff; font-size:1.11rem; line-height:1.79rem;}
.gift-guide-bnr a:after {content:' '; position:absolute; bottom:0; left:50%; width:32rem; height:6.61rem; margin-left:-16rem; background:url(http://fiximage.10x10.co.kr/m/2017/today/bg_gift_guide.jpg) no-repeat 50% 0; background-size:32rem auto;}
.gift-guide-bnr + .menu-etc {margin-top:4.69rem;}

.menu-etc {margin-top:1.96rem;}
.menu-etc ul {width:27.32rem; margin:0 auto;}
.menu-etc ul:after {content:' '; display:block; clear:both;}
.menu-etc li {float:left; position:relative; width:6.83rem; height:4.27rem; border-left:1px solid #ededed; text-align:center;}
.menu-etc li:first-child {border:0;}
.menu-etc .icon {position:absolute; top:-0.34rem; left:50%; width:2.56rem; height:3.07rem; margin-left:-1.28rem;}
.menu-etc .icon-coupon {top:-0.17rem; background-position:-13.86rem -2.77rem;}
.menu-etc .icon-gift {background-position:-8.32rem 0;}
.menu-etc .icon-exhibition {background-position:0 0;}
.menu-etc .icon-event {top:-0.43rem; background-position:-11.09rem -2.77rem;}
.menu-etc .name {display:block; padding-top:2.81rem; color:#000; font:1.11rem 'AppleSDGothicNeo-Medium';}
.menu-etc .badge {position:absolute; top:-0.6rem; right:0.94rem; min-width:1.88rem; height:1.88rem; border-radius:1.88rem; padding:0.09rem 0.43rem 0; background-color:#00b25f; color:#fff; font:bold 1.02rem/1.88rem 'AvenirNext-DemiBold'; letter-spacing:-0.0853rem; text-align:center;}

/* category icon */
.menu-category .icon-category101 {background-position:0 0;}
.menu-category .icon-category102 {background-position:-3.11rem 0;}
.menu-category .icon-category124 {background-position:-6.23rem 0;}
.menu-category .icon-category121 {background-position:-9.34rem 0;}
.menu-category .icon-category122 {background-position:-12.46rem 0;}
.menu-category .icon-category120 {background-position:0 -3.11rem;}
.menu-category .icon-category112 {background-position:-3.11rem -3.11rem;}
.menu-category .icon-category119 {background-position:-6.23rem -3.11rem;}
.menu-category .icon-category117 {background-position:-9.34rem -3.11rem;}
.menu-category .icon-category116 {background-position:-12.46rem -3.11rem;}
.menu-category .icon-category125 {background-position:0 -6.23rem;}
.menu-category .icon-category118 {background-position:-3.11rem -6.23rem;}
.menu-category .icon-category103 {background-position:-6.23rem -6.23rem;}
.menu-category .icon-category104 {background-position:-9.34rem -6.23rem;}
.menu-category .icon-category115 {background-position:-12.46rem -6.23rem;}
.menu-category .icon-category110 {background-position:0 -9.34rem;}

.category-item-content {padding-top:1.71rem;}
.category-item-content section:first-child {margin-top:0;}

/* contents banner */
.contents-bnr {margin-top:4.69rem;}
.contents-bnr .playing-bnr,
.contents-bnr .hitchhiker-bnr {margin-top:3.07rem;}
.contents-bnr a {display:block; position:relative;}
.contents-bnr .headline {font:bold 1.54rem 'AvenirNext-Bold'; letter-spacing:0.17rem;}
.contents-bnr .subcopy {margin-top:0.26rem; color:#4a4a4a; font-size:1.19rem; line-height:1.71rem;}
.contents-bnr .thumbnail {border-radius:1.71rem;}
.playing-bnr .thumbnail, .hitchhiker-bnr .thumbnail {width:29.01rem; height:19.28rem; margin:0 auto;}
.playing-bnr .thumbnail img, .hitchhiker-bnr .thumbnail img {width:100%; height:100%; border-radius:1.71rem;}
.playing-bnr .desc, .hitchhiker-bnr .desc {position:relative; z-index:10; width:29.01rem; padding:0 1.54rem; margin:-1.37rem auto 0;}
.playing-bnr > a:after,
.hitchhiker-bnr > a:after {content:' '; position:absolute; top:14.68rem; left:50%; z-index:5; width:29.01rem; height:5.21rem; margin-left:-14.505rem; background-position:0 -22.19rem; background-size:34.2rem auto;}
.hitchhiker-bnr > a:after {transform:rotateY(180deg); -webkit-transform:rotateY(180deg);}
.contents-bnr .culture-bnr {margin-top:3.75rem; background-color:#f7f7f7;}
.culture-bnr > a {display:table; padding:3.41rem 3.07rem 3.24rem; align-items:center;}
.culture-bnr .thumbnail {display:table-cell; position:relative; width:12.46rem; height:12.97rem; background-color:transparent; text-align:center;}
.culture-bnr .thumbnail:after {content:' '; position:absolute; bottom:-0.85rem; left:0.17rem; width:11.09rem; height:0; border:0.51rem solid rgba(0, 0, 0, 0); border-top:0 solid; border-bottom:1.71rem solid #e9e9e9;}
.culture-bnr .thumbnail:before {display:none;}
.culture-bnr .thumbnail img {position:relative; z-index:5; width:8.79rem; height:100%; box-shadow:0 0 0.17rem 0 rgba(0, 0, 0, 0.12);}
.culture-bnr .desc {display:table-cell; text-align:center; vertical-align:middle;}
.culture-bnr .headline {font-size:1.71rem; line-height:2.39rem; letter-spacing:0.187rem;}
.culture-bnr .subcopy {margin-top:0.85rem;}
.culture-bnr .label {position:absolute; top:-1.96rem; left:0; z-index:10;}

.shipping-bnr {margin-top:5.55rem; padding-bottom:5.63rem;}
.shipping-bnr a {display:block; padding:0 4.01rem; font-size:1.62rem; line-height:2.47rem; letter-spacing:0.23rem;}
.shipping-bnr p {position:relative; padding-bottom:1.02rem; border-bottom:2px solid #000;}
.shipping-bnr em {display:inline-block; font:bold 1.79rem 'AppleSDGothicNeo-Bold';}
.shipping-bnr em:after {content:' '; display:inline-block; width:1.28rem; height:1.54rem; margin-left:0.34rem; background-position:-15.45rem 0; vertical-align:-0.17rem;}
.shipping-bnr .icon-car {position:absolute; bottom:0; right:0; width:3.58rem; height:2.39rem; background-position:-10.62rem -8.32rem; opacity:1;}

.tentencar {animation:tentencar 2s cubic-bezier(0.19, 1, 0.22, 1) forwards; -webkit-animation:tentencar 2s cubic-bezier(0.19, 1, 0.22, 1) forwards;}
@keyframes tentencar {
	0% {transform:translateX(200px);}
	30% {transform:translateX(100px);}
	80% {transform:translateX(50px);}
	100% {transform:translateX(0); opacity:1;}
}
@-webkit-keyframes tentencar {
	0% {transform:translateX(200px);}
	30% {transform:translateX(100px);}
	80% {transform:translateX(50px);}
	100% {transform:translateX(0); opacity:1;}
}

.category-item-content .category-item-list:last-child {padding-bottom:0.68rem;}
.category-item-list {padding:2.73rem 0 0.26rem}
.category-item-list .items ul {margin-top:-0.09rem;}
.category-item-list .items li {margin-top:1.28rem;}
.category-item-list .items .price {margin-top:0.76rem;}
.category-item-list .exhibition-list {margin-top:1.96rem;}
.category-item-list .btn-group {margin-top:-0.34rem;}

/* iphoneX */
@media only screen and (device-width : 375px) and (device-height : 812px) and (-webkit-device-pixel-ratio : 3) {
	.body-main.ios .btn-top {bottom:75px !important;}
}
</style>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/jqueryui/jquery-ui-1.10.2.custom.min.js"></script>
</script>
<body>

<% If addtype="1" Then %>
	<section id="enjoyevent1" class="exhibition-list" v-cloak>
		<h2 class="hidden">기획전</h2> 
		<div class="list-card type-align-left">
			<ul>
				<li>
					<a href="http://m.10x10.co.kr<%=linkurl%>" target="_blank">
						<div class="thumbnail">
							<img src="<%=vTrendImg%>" alt="<%=evtalt%>" style="display: block; width: 100%;">
						</div> 
						<div class="desc">
							<p>
								<b class="headline">
									<% If sale_per<>"" Then %>
										<span class="ellipsis"><%=evttitle%></span> 
										<b class="discount color-red"><%=sale_per%></b>
									<% Else %>
										<span class="ellipsis full"><%=evttitle%></span>
									<% End If %>
								</b> 
								<span class="subcopy">
									<% If evttag <> "" Then %>
										<span class="label label-color">
											<% If coupon_flag = "1" Then %>
												<em class="color-green"><%=evttag%></em>
											<% ElseIf coupon_flag = "0" Then %>
												<em class="color-blue"><%=evttag%></em>
											<% End If %>
										</span>
									<% End If %>
									<%=evttitle2%>
								</span>
							</p>
						</div>
					</a>
				</li>
			</ul>
		</div>
	</section>
<% End If %>

<% If addtype="2" Then %>
	<div id="enjoyeventitem1" class="exhibition-plus-item">
		<div>
			<div class="list-card type-align-left">
				<a href="http://m.10x10.co.kr/<%=linkurl%>" target="_blank">
					<div class="thumbnail"><img src="<%=vTrendImg%>" alt=""></div>
					<p class="desc">
						<b class="headline">
							<% If sale_per<>"" Then %>
								<span class="ellipsis"><%=evttitle%></span> 
								<b class="discount color-red"><%=sale_per%></b>
							<% Else %>
								<span class="ellipsis full"><%=evttitle%></span>
							<% End If %>
						</b> 
						<span class="subcopy">
							<% If evttag <> "" Then %>
								<span class="label label-color">
									<% If coupon_flag = "1" Then %>
										<em class="color-green"><%=evttag%></em>
									<% ElseIf coupon_flag = "0" Then %>
										<em class="color-blue"><%=evttag%></em>
									<% End If %>
								</span>
							<% End If %>
							<%=evttitle2%>
						</span>
					</p>
				</a>
			</div>
			<div class="items">
				<ul>
					<li>
						<a href="http://m.10x10.co.kr<%=itemid1url%>" target="_blank">
							<div class="thumbnail"><img src="<%=itemimg1%>" alt=""></div>
							<div class="desc">
								<div class="price">
									<% If sale1 <> "" Then %>
										<b class="discount color-red"><%=sale1%></b>
									<% End If %>
									<b class="sum"><%=price1%><span class="won">원</span></b>
								</div>
							</div>
						</a>
					</li>
					<li>
						<a href="http://m.10x10.co.kr<%=itemid2url%>" target="_blank">
							<div class="thumbnail"><img src="<%=itemimg2%>" alt=""></div>
							<div class="desc">
								<div class="price">
									<% If sale2 <> "" Then %>
										<b class="discount color-red"><%=sale2%></b>
									<% End If %>
									<b class="sum"><%=price2%><span class="won">원</span></b>
								</div>
							</div>
						</a>
					</li>
					<li>
						<a href="http://m.10x10.co.kr<%=itemid3url%>" target="_blank">
							<div class="thumbnail"><img src="<%=itemimg3%>" alt=""></div>
							<div class="desc">
								<div class="price">
									<% If sale3 <> "" Then %>
										<b class="discount color-red"><%=sale3%></b>
									<% End If %>
									<b class="sum"><%=price3%><span class="won">원</span></b>
								</div>
							</div>
						</a>
					</li>
				</ul>
			</div>
		</div>
	</div>
<% End If %>
</body>
</html>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->