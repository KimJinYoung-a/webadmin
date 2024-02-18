<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<%

dim CS_COMPANYID

CS_COMPANYID = "thefingers"



dim DATABASE_APPLICATION
dim CS_DATABASE_APPLICATION

dim TABLE_ORDERMASTER, TABLE_ORDERDETAIL, TABLE_ITEM, TABLE_SONGJANG_DIV, TABLE_ORDER_GIFT
dim TABLE_CSMASTER, TABLE_CSDETAIL
dim TABLE_CS_REFUND, TABLE_CS_COMMON_CODE, TABLE_CS_MEMO, TABLE_CS_DELIVERY
dim TABLE_MILEAGELOG, TABLE_USER_CURRENT_MILEAGE
dim TABLE_CS_CONFIRM
dim TABLE_USER_C
dim TABLE_CS_BRAND_MEMO

dim TABLE_UPCHE_ADD_JUNGSAN
dim TABLE_PARTNER
dim TABLE_PARTNER_GROUP
dim TABLE_MIBEASONG_LIST
dim TABLE_CATEGORY_LARGE


dim FIELD_DETAILIDX, FIELD_ITEMCOUPONIDX, FIELD_CURRENT_MILEAGE
dim FIELD_ITEMVAT

dim PROC_MINUS_ORDER_INVALID_CNT

dim MAIN_SITENAME1
dim MAIN_SITENAME2

dim EXCLUDE_SITENAME

dim DIRECTORY_IMAGE_SMALL

dim CS_MAIN_PHONENO, CS_MAIL_SITENAME, CS_RECEIVE_BANK_INFO
Dim CS_MAIL_ADDR


'==============================================================================
if application("Svr_Info")="Dev" then

	if (CS_COMPANYID = "10x10") then

		wwwUrl		    = "http://2010www.10x10.co.kr"
		webImgUrl		= "http://testwebimage.10x10.co.kr"

	elseif (CS_COMPANYID = "thefingers") then

		wwwUrl			= "http://test.thefingers.co.kr"
		webImgUrl		= "http://testimage.thefingers.co.kr"

	else

		'에러

	end if

else

	if (CS_COMPANYID = "10x10") then

		wwwUrl		    = "http://www.10x10.co.kr"
		webImgUrl		= "http://webimage.10x10.co.kr"

	elseif (CS_COMPANYID = "thefingers") then

		wwwUrl			= "http://www.thefingers.co.kr"
		webImgUrl		= "http://image.thefingers.co.kr"

	else

		'에러

	end if

end if



'==============================================================================
if (CS_COMPANYID = "10x10") then

	DATABASE_APPLICATION = "db_main11"			'일부러 에러 발생시킴(보안문제 확인 필요)
	CS_DATABASE_APPLICATION = "db_main11"



	TABLE_ORDERMASTER 		= "[db_order].[dbo].tbl_order_master"
	TABLE_ORDERDETAIL 		= "[db_order].[dbo].tbl_order_detail"
	TABLE_ITEM 				= "db_item.dbo.tbl_item"
	TABLE_SONGJANG_DIV 		= "db_order.dbo.tbl_songjang_div"
	'TABLE_ORDER_GIFT 		= "[db_order].[dbo].tbl_order_gift"

	TABLE_CSMASTER 			= "[db_cs].[dbo].tbl_new_as_list"
	TABLE_CSDETAIL 			= "[db_cs].[dbo].tbl_new_as_detail"

	TABLE_CS_REFUND 		= "[db_cs].[dbo].tbl_as_refund_info"
	TABLE_CS_COMMON_CODE 	= "[db_cs].[dbo].tbl_cs_comm_code"
	TABLE_CS_MEMO 			= "[db_cs].[dbo].tbl_cs_memo"
	TABLE_CS_DELIVERY 		= "[db_cs].[dbo].tbl_new_as_delivery"
	TABLE_MILEAGELOG		= "[db_user].[dbo].tbl_mileagelog"
	TABLE_USER_CURRENT_MILEAGE = "[db_user].[dbo].tbl_user_current_mileage"

	TABLE_CS_CONFIRM		= "[db_cs].[dbo].tbl_new_as_confirm"
	TABLE_UPCHE_ADD_JUNGSAN = "[db_cs].[dbo].tbl_as_upcheAddjungsan"
    TABLE_PARTNER           = "[db_partner].[dbo].tbl_partner"
    TABLE_PARTNER_GROUP		= "[db_partner].[dbo].tbl_partner_group"
    TABLE_CS_BRAND_MEMO		= "[db_cs].[dbo].tbl_cs_brand_memo"

    TABLE_MIBEASONG_LIST	= "[db_temp].dbo.tbl_mibeasong_list"
    TABLE_USER_C			= "[db_user].[dbo].tbl_user_c"
    TABLE_CATEGORY_LARGE	= "db_item.dbo.tbl_cate_large"



	FIELD_DETAILIDX 		= "idx"
	FIELD_ITEMCOUPONIDX 	= "itemcouponidx"
	FIELD_CURRENT_MILEAGE 	= "jumunmileage"
	FIELD_ITEMVAT			= "itemvat"



	PROC_MINUS_ORDER_INVALID_CNT = "db_order.dbo.sp_Ten_MinusOrderInValidCnt"



	'자체 사이트(마일리지 적립 등)
	MAIN_SITENAME1 = "10x10"
	MAIN_SITENAME2 = "xxxxxxxxxxx"

	EXCLUDE_SITENAME = "yyyyyyyyyyyy"

	DIRECTORY_IMAGE_SMALL = "/image/small/"



	CS_MAIN_PHONENO = "1644-6030"
	CS_MAIL_SITENAME = "텐바이텐"
	CS_MAIL_ADDR     = "mailzine@10x10.co.kr"
	CS_RECEIVE_BANK_INFO = "조흥은행534-01-016039"

elseif (CS_COMPANYID = "thefingers") then

	DATABASE_APPLICATION = "db_academy"
	CS_DATABASE_APPLICATION = "db_main"



	TABLE_ORDERMASTER 		= "[db_academy].[dbo].tbl_academy_order_master"
	TABLE_ORDERDETAIL 		= "[db_academy].[dbo].tbl_academy_order_detail"
	TABLE_ITEM 				= "[db_academy].[dbo].tbl_diy_item"
	TABLE_SONGJANG_DIV 		= "[db_academy].[dbo].tbl_songjang_div"
	'TABLE_ORDER_GIFT 		= "[db_academy].[dbo].tbl_order_gift"

	TABLE_CSMASTER 			= "[db_academy].[dbo].tbl_academy_as_list"
	TABLE_CSDETAIL 			= "[db_academy].[dbo].tbl_academy_as_detail"

	TABLE_CS_REFUND 		= "[db_academy].[dbo].tbl_academy_as_refund_info"
	TABLE_CS_COMMON_CODE 	= "[db_academy].[dbo].tbl_academy_cs_comm_code"
	TABLE_CS_MEMO 			= "[db_academy].[dbo].tbl_academy_cs_memo"
	TABLE_CS_DELIVERY 		= "[db_academy].[dbo].tbl_academy_as_delivery"

	TABLE_MILEAGELOG		= "[db_user].[dbo].tbl_mileagelog"
	TABLE_USER_CURRENT_MILEAGE = "[db_user].[dbo].tbl_user_current_mileage"

	TABLE_CS_CONFIRM		= "[db_academy].[dbo].tbl_academy_as_confirm"
	TABLE_UPCHE_ADD_JUNGSAN = "[db_academy].[dbo].tbl_academy_as_upcheAddjungsan"
    TABLE_PARTNER           = "[TENDB].[db_partner].[dbo].tbl_partner"
    'TABLE_PARTNER           = "[db_academy].[dbo].tbl_lec_user"


    TABLE_PARTNER_GROUP		= "[TENDB].[db_partner].[dbo].tbl_partner_group"
    TABLE_CS_BRAND_MEMO		= "[db_academy].[dbo].tbl_academy_cs_brand_memo"

    TABLE_MIBEASONG_LIST	= "[db_academy].dbo.tbl_academy_mibeasong_list"
    TABLE_USER_C			= "[TENDB].[db_user].[dbo].tbl_user_c"
    TABLE_CATEGORY_LARGE	= "db_academy.dbo.tbl_diy_item_Cate_large"



	FIELD_DETAILIDX 		= "detailidx"
	FIELD_ITEMCOUPONIDX 	= "leccouponidx"
	FIELD_CURRENT_MILEAGE 	= "academymileage"
	FIELD_ITEMVAT			= "couponNotAsigncost"



	PROC_MINUS_ORDER_INVALID_CNT = "db_academy.dbo.sp_Academy_MinusOrderInValidCnt"



	'자체 사이트(마일리지 적립 등)
	MAIN_SITENAME1 = "academy"
	MAIN_SITENAME2 = "diyitem"

	EXCLUDE_SITENAME = "academy"

	DIRECTORY_IMAGE_SMALL = "/diyitem/webimage/small/"



	CS_MAIN_PHONENO = "02-741-9070"
	CS_MAIL_SITENAME = "핑거스아카데미"
	CS_MAIL_ADDR     = "customer@thefingers.co.kr"
	CS_RECEIVE_BANK_INFO = "없음"

else

	'에러

end if

%>