<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
dim itemid : itemid = requestCheckvar(request("itemid"),10)
dim mallid  : mallid = requestCheckvar(request("mallid"),32)
dim mode    : mode = requestCheckvar(request("mode"),32)
dim mngOptAdd : mngOptAdd = requestCheckvar(request("mngOptAdd"),10)
dim optAddPrcRegType: optAddPrcRegType = requestCheckvar(request("optAddPrcRegType"),10)

dim sqlStr, arrRows
dim regitemname,outmallGoodNo,optaddPrcCnt, lastStatCheckDate
dim TEN_URI
dim i

'rw itemid
'rw mallid

if (mode="EDTRegType") then
    if (mallid="lotteCom") then
        sqlStr = ""
        sqlStr = sqlStr & " update R"&VbCRLF
        sqlStr = sqlStr & " set optAddPrcRegType="&optAddPrcRegType&VbCRLF
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_lotte_regItem R"&VbCRLF
        sqlStr = sqlStr & " where R.itemid="&itemid&VbCRLF

        dbget.Execute sqlStr
    elseif (mallid="lotteimall") then
        sqlStr = ""
        sqlStr = sqlStr & " update R"&VbCRLF
        sqlStr = sqlStr & " set optAddPrcRegType="&optAddPrcRegType&VbCRLF
        sqlStr = sqlStr & " from [db_item].[dbo].tbl_Ltimall_regItem R"&VbCRLF
        sqlStr = sqlStr & " where R.itemid="&itemid&VbCRLF

        dbget.Execute sqlStr
    else
        rw "미지정["&mallid&"]"
        dbget.Close() : response.end
    end if
end if



if (mallid="lotteCom") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, isNULL(lotteTmpGoodno,lotteGoodno) as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_item].[dbo].tbl_lotte_regItem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="lotteimall") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, isNULL(LtimallTmpGoodno,LtimallGoodno) as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_item].[dbo].tbl_Ltimall_regItem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="cjmall") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, '' as outmallregName, cjmallprdno as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_item].[dbo].tbl_cjmall_regItem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="gsshop") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, '' as outmallregName, GSShopGoodno as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_item].[dbo].tbl_gsshop_regItem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="interpark") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, '' as outmallregName, interParkPrdNo as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType "&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_item].[dbo].tbl_interpark_reg_item"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="auction1010") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, auctionGoodno as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].tbl_auction_regItem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="gmarket1010") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, GmarketGoodno as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].tbl_gmarket_regItem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="homeplus") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, HomeplusGoodno as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].tbl_homeplus_regItem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="nvstorefarm") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, nvstorefarmGoodno as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].tbl_nvstorefarm_regItem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="nvstorefarmclass") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, nvClassGoodno as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].tbl_nvstorefarmclass_regItem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="Mylittlewhoopee") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, MylittlewhoopeeGoodno as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].tbl_Mylittlewhoopee_regItem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="nvstoremoonbangu") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, nvstoremoonbanguGoodno as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].tbl_nvstoremoonbangu_regItem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="nvstoregift") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, nvstoregiftGoodno as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].tbl_nvstoregift_regItem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="11stmy") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, my11stGoodno as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].tbl_my11st_regItem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="gmarket") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, gmarketGoodno as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].tbl_gmarket_regitem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="11st1010") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, st11Goodno as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].tbl_11st_regitem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="sabangnet") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, sabangnetGoodno as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].tbl_sabangnet_regitem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="ssg") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, ssgGoodno as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].tbl_ssg_regitem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="zilingo") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, zilingoGoodno as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].tbl_zilingo_regitem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="halfclub") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, HalfClubGoodno as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].tbl_halfclub_regitem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="lfmall") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, lfmallGoodNo as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].[tbl_lfmall_regItem]"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="coupang") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, coupangGoodno as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].tbl_coupang_regitem"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="kakaogift") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, kakaoGiftGoodNo as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].[tbl_kakaoGift_regItem]"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="hmall") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, hmallGoodNo as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].[tbl_hmall_regItem]"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="WMP") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, wemakeGoodNo as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].[tbl_wemake_regItem]"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="wmpfashion") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, wfwemakeGoodNo as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].[tbl_wfwemake_regItem]"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="ezwel") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, ezwelGoodNo as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].[tbl_ezwel_regItem]"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="lotteon") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, lotteonGoodNo as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].[tbl_lotteon_regItem]"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="market_for") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, marketforGoodNo as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].[tbl_marketfor_regItem]"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="shintvshopping") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, shintvshoppingGoodNo as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].[tbl_shintvshopping_regItem]"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="wetoo1300k") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, wetoo1300kGoodNo as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].[tbl_wetoo1300k_regItem]"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="skstoa") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, skstoaGoodNo as outmallGoodNo , optaddPrcCnt, isNULL(optAddPrcRegType,0) as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].[tbl_skstoa_regItem]"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="shopify") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, shopifyGoodNo as outmallGoodNo , '' as optaddPrcCnt, '' as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].[tbl_shopify_regItem]"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="kakaostore") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, kakaostoreGoodNo as outmallGoodNo , '' as optaddPrcCnt, '' as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].[tbl_kakaostore_regItem]"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="boribori1010") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, boriboriGoodNo as outmallGoodNo , '' as optaddPrcCnt, '' as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].[tbl_boribori_regItem]"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="bindmall1010") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, bindmallGoodNo as outmallGoodNo , '' as optaddPrcCnt, '' as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].[tbl_bindmall_regItem]"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="wconcept1010") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, wconceptGoodNo as outmallGoodNo , '' as optaddPrcCnt, '' as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].[tbl_wconcept_regItem]"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
elseif (mallid="benepia1010") then
    sqlStr = ""
    sqlStr = sqlStr & " select top 1 itemid, regitemname as outmallregName, benepiaGoodNo as outmallGoodNo , '' as optaddPrcCnt, '' as optAddPrcRegType"&VbCRLF
    sqlStr = sqlStr & " ,lastStatCheckDate"
    sqlStr = sqlStr & " from [db_etcmall].[dbo].[tbl_benepia_regItem]"&VbCRLF
    sqlStr = sqlStr & " where itemid="&itemid&VbCRLF
else
    rw "미지정["&mallid&"]"
    dbget.Close() : response.end
end if
rsget.Open sqlStr, dbget
if Not(rsget.EOF or rsget.BOF) then
    regitemname     = rsget("outmallregName")
    outmallGoodNo   = rsget("outmallGoodNo")
    optaddPrcCnt    = rsget("optaddPrcCnt")
    optAddPrcRegType = rsget("optAddPrcRegType")
    lastStatCheckDate= rsget("lastStatCheckDate")
end if
rsget.close

If mallid = "shopify" Then
    sqlStr = ""
    sqlStr = sqlStr & " select top 300 '1' as Sitename,o.itemid,o.itemoption,o.isusing,o.optsellyn"&VbCRLF
    sqlStr = sqlStr & " ,o.optlimityn,o.optlimitno,o.optlimitsold"&VbCRLF
    sqlStr = sqlStr & " ,mo.optionname,o.optaddprice,o.optaddbuyprice"&VbCRLF
    sqlStr = sqlStr & " ,mo.optionTypeName"&VbCRLF
    sqlStr = sqlStr & " ,r.outmalloPtCode,r.outmallOptName,r.outmallSellyn,r.outmalllimityn,r.outmalllimitno,r.outmallAddPrice"&VbCRLF
    sqlStr = sqlStr & " from [db_item].[dbo].tbl_item_option o"&VbCRLF
    sqlStr = sqlStr & " 	left join [db_item].[dbo].tbl_OutMall_regedoption r"&VbCRLF
    sqlStr = sqlStr & " 	on r.mallid='"&mallid&"'"&VbCRLF
    sqlStr = sqlStr & " 	and o.itemid=r.itemid"&VbCRLF
    sqlStr = sqlStr & " 	and o.itemoption=r.itemoption"&VbCRLF
    sqlStr = sqlStr & " 	    LEFT JOIN db_item.[dbo].[tbl_item_multiLang_option] as mo on mo.itemid = o.itemid and mo.itemoption = IsNULL(o.itemoption,'0000') and mo.countryCd = 'EN' "
    sqlStr = sqlStr & " where o.itemid="&itemid&VbCRLF
    sqlStr = sqlStr & " Union"&VbCRLF
    sqlStr = sqlStr & " select top 100 'Z' as Sitename,r.itemid,r.itemoption,NULL as isusing,NULL as optsellyn"&VbCRLF
    sqlStr = sqlStr & " ,NULL as optlimityn,NULL as optlimitno,NULL as optlimitsold"&VbCRLF
    sqlStr = sqlStr & " ,NULL as optionname,NULL as optaddprice,NULL as optaddbuyprice"&VbCRLF
    sqlStr = sqlStr & " ,NULL as optionTypeName"
    sqlStr = sqlStr & " ,r.outmalloPtCode,r.outmallOptName,r.outmallSellyn,r.outmalllimityn,r.outmalllimitno,r.outmallAddPrice"&VbCRLF
    sqlStr = sqlStr & " from [db_item].[dbo].tbl_OutMall_regedoption r"&VbCRLF
    sqlStr = sqlStr & " 	left join [db_item].[dbo].tbl_item_option o"&VbCRLF
    sqlStr = sqlStr & " 	on o.itemid=r.itemid"&VbCRLF
    sqlStr = sqlStr & " 	and o.itemoption=r.itemoption"&VbCRLF
    sqlStr = sqlStr & " 	    LEFT JOIN db_item.[dbo].[tbl_item_multiLang_option] as mo on mo.itemid = o.itemid and mo.itemoption = IsNULL(o.itemoption,'0000') and mo.countryCd = 'EN' "
    sqlStr = sqlStr & " where r.mallid='"&mallid&"'"&VbCRLF
    sqlStr = sqlStr & " and r.itemid="&itemid&VbCRLF
    sqlStr = sqlStr & " and o.itemid is NULL"&VbCRLF
    sqlStr = sqlStr & " order by Sitename"&VbCRLF
ElseIf mallid = "kakaostore" Then
    sqlStr = ""
    sqlStr = sqlStr & " select top 300 '1' as Sitename,o.itemid,o.itemoption,o.isusing,o.optsellyn"&VbCRLF
    sqlStr = sqlStr & " ,o.optlimityn,o.optlimitno,o.optlimitsold"&VbCRLF
    sqlStr = sqlStr & " ,o.optionname,o.optaddprice,o.optaddbuyprice"&VbCRLF
    sqlStr = sqlStr & " ,o.optionTypeName"&VbCRLF
    sqlStr = sqlStr & " ,r.outmalloPtCode,r.outmallOptName,r.outmallSellyn,r.outmalllimityn,r.outmalllimitno,r.outmallAddPrice"&VbCRLF
    sqlStr = sqlStr & " from [db_item].[dbo].tbl_item_option o"&VbCRLF
    sqlStr = sqlStr & " 	left join [db_item].[dbo].tbl_kakaostore_regedoption r"&VbCRLF
    sqlStr = sqlStr & " 	on o.itemid=r.itemid"&VbCRLF
    sqlStr = sqlStr & " 	and o.itemoption=r.itemoption"&VbCRLF
    sqlStr = sqlStr & " where o.itemid="&itemid&VbCRLF
    sqlStr = sqlStr & " Union"&VbCRLF
    sqlStr = sqlStr & " select top 100 'Z' as Sitename,r.itemid,r.itemoption,NULL as isusing,NULL as optsellyn"&VbCRLF
    sqlStr = sqlStr & " ,NULL as optlimityn,NULL as optlimitno,NULL as optlimitsold"&VbCRLF
    sqlStr = sqlStr & " ,NULL as optionname,NULL as optaddprice,NULL as optaddbuyprice"&VbCRLF
    sqlStr = sqlStr & " ,NULL as optionTypeName"
    sqlStr = sqlStr & " ,r.outmalloPtCode,r.outmallOptName,r.outmallSellyn,r.outmalllimityn,r.outmalllimitno,r.outmallAddPrice"&VbCRLF
    sqlStr = sqlStr & " from [db_item].[dbo].tbl_kakaostore_regedoption r"&VbCRLF
    sqlStr = sqlStr & " 	left join [db_item].[dbo].tbl_item_option o"&VbCRLF
    sqlStr = sqlStr & " 	on o.itemid=r.itemid"&VbCRLF
    sqlStr = sqlStr & " 	and o.itemoption=r.itemoption"&VbCRLF
    sqlStr = sqlStr & " where r.itemid="&itemid&VbCRLF
    sqlStr = sqlStr & " and o.itemid is NULL"&VbCRLF
    sqlStr = sqlStr & " order by Sitename"&VbCRLF
ElseIf mallid = "boribori1010" Then
    sqlStr = ""
    sqlStr = sqlStr & " select top 300 '1' as Sitename,o.itemid,o.itemoption,o.isusing,o.optsellyn"&VbCRLF
    sqlStr = sqlStr & " ,o.optlimityn,o.optlimitno,o.optlimitsold"&VbCRLF
    sqlStr = sqlStr & " ,o.optionname,o.optaddprice,o.optaddbuyprice"&VbCRLF
    sqlStr = sqlStr & " ,o.optionTypeName"&VbCRLF
    sqlStr = sqlStr & " ,r.outmalloPtCode,r.outmallOptName,r.outmallSellyn,r.outmalllimityn,r.outmalllimitno,r.outmallAddPrice"&VbCRLF
    sqlStr = sqlStr & " from [db_item].[dbo].tbl_item_option o"&VbCRLF
    sqlStr = sqlStr & " 	left join [db_item].[dbo].tbl_boribori_regedoption r"&VbCRLF
    sqlStr = sqlStr & " 	on o.itemid=r.itemid"&VbCRLF
    sqlStr = sqlStr & " 	and o.itemoption=r.itemoption"&VbCRLF
    sqlStr = sqlStr & " where o.itemid="&itemid&VbCRLF
    sqlStr = sqlStr & " Union"&VbCRLF
    sqlStr = sqlStr & " select top 100 'Z' as Sitename,r.itemid,r.itemoption,NULL as isusing,NULL as optsellyn"&VbCRLF
    sqlStr = sqlStr & " ,NULL as optlimityn,NULL as optlimitno,NULL as optlimitsold"&VbCRLF
    sqlStr = sqlStr & " ,NULL as optionname,NULL as optaddprice,NULL as optaddbuyprice"&VbCRLF
    sqlStr = sqlStr & " ,NULL as optionTypeName"
    sqlStr = sqlStr & " ,r.outmalloPtCode,r.outmallOptName,r.outmallSellyn,r.outmalllimityn,r.outmalllimitno,r.outmallAddPrice"&VbCRLF
    sqlStr = sqlStr & " from [db_item].[dbo].tbl_boribori_regedoption r"&VbCRLF
    sqlStr = sqlStr & " 	left join [db_item].[dbo].tbl_item_option o"&VbCRLF
    sqlStr = sqlStr & " 	on o.itemid=r.itemid"&VbCRLF
    sqlStr = sqlStr & " 	and o.itemoption=r.itemoption"&VbCRLF
    sqlStr = sqlStr & " where r.itemid="&itemid&VbCRLF
    sqlStr = sqlStr & " and o.itemid is NULL"&VbCRLF
    sqlStr = sqlStr & " order by Sitename"&VbCRLF
ElseIf mallid = "bindmall1010" Then
    sqlStr = ""
    sqlStr = sqlStr & " select top 300 '1' as Sitename,o.itemid,o.itemoption,o.isusing,o.optsellyn"&VbCRLF
    sqlStr = sqlStr & " ,o.optlimityn,o.optlimitno,o.optlimitsold"&VbCRLF
    sqlStr = sqlStr & " ,o.optionname,o.optaddprice,o.optaddbuyprice"&VbCRLF
    sqlStr = sqlStr & " ,o.optionTypeName"&VbCRLF
    sqlStr = sqlStr & " ,r.outmalloPtCode,r.outmallOptName,r.outmallSellyn,r.outmalllimityn,r.outmalllimitno,r.outmallAddPrice"&VbCRLF
    sqlStr = sqlStr & " from [db_item].[dbo].tbl_item_option o"&VbCRLF
    sqlStr = sqlStr & " 	left join [db_item].[dbo].tbl_bindmall_regedoption r"&VbCRLF
    sqlStr = sqlStr & " 	on o.itemid=r.itemid"&VbCRLF
    sqlStr = sqlStr & " 	and o.itemoption=r.itemoption"&VbCRLF
    sqlStr = sqlStr & " where o.itemid="&itemid&VbCRLF
    sqlStr = sqlStr & " Union"&VbCRLF
    sqlStr = sqlStr & " select top 100 'Z' as Sitename,r.itemid,r.itemoption,NULL as isusing,NULL as optsellyn"&VbCRLF
    sqlStr = sqlStr & " ,NULL as optlimityn,NULL as optlimitno,NULL as optlimitsold"&VbCRLF
    sqlStr = sqlStr & " ,NULL as optionname,NULL as optaddprice,NULL as optaddbuyprice"&VbCRLF
    sqlStr = sqlStr & " ,NULL as optionTypeName"
    sqlStr = sqlStr & " ,r.outmalloPtCode,r.outmallOptName,r.outmallSellyn,r.outmalllimityn,r.outmalllimitno,r.outmallAddPrice"&VbCRLF
    sqlStr = sqlStr & " from [db_item].[dbo].tbl_bindmall_regedoption r"&VbCRLF
    sqlStr = sqlStr & " 	left join [db_item].[dbo].tbl_item_option o"&VbCRLF
    sqlStr = sqlStr & " 	on o.itemid=r.itemid"&VbCRLF
    sqlStr = sqlStr & " 	and o.itemoption=r.itemoption"&VbCRLF
    sqlStr = sqlStr & " where r.itemid="&itemid&VbCRLF
    sqlStr = sqlStr & " and o.itemid is NULL"&VbCRLF
    sqlStr = sqlStr & " order by Sitename"&VbCRLF
ElseIf mallid = "wconcept1010" Then
    sqlStr = ""
    sqlStr = sqlStr & " select top 300 '1' as Sitename,o.itemid,o.itemoption,o.isusing,o.optsellyn"&VbCRLF
    sqlStr = sqlStr & " ,o.optlimityn,o.optlimitno,o.optlimitsold"&VbCRLF
    sqlStr = sqlStr & " ,o.optionname,o.optaddprice,o.optaddbuyprice"&VbCRLF
    sqlStr = sqlStr & " ,o.optionTypeName"&VbCRLF
    sqlStr = sqlStr & " ,r.outmalloPtCode,r.outmallOptName,r.outmallSellyn,r.outmalllimityn,r.outmalllimitno,r.outmallAddPrice"&VbCRLF
    sqlStr = sqlStr & " from [db_item].[dbo].tbl_item_option o"&VbCRLF
    sqlStr = sqlStr & " 	left join [db_item].[dbo].tbl_wconcept_regedoption r"&VbCRLF
    sqlStr = sqlStr & " 	on o.itemid=r.itemid"&VbCRLF
    sqlStr = sqlStr & " 	and o.itemoption=r.itemoption"&VbCRLF
    sqlStr = sqlStr & " where o.itemid="&itemid&VbCRLF
    sqlStr = sqlStr & " Union"&VbCRLF
    sqlStr = sqlStr & " select top 100 'Z' as Sitename,r.itemid,r.itemoption,NULL as isusing,NULL as optsellyn"&VbCRLF
    sqlStr = sqlStr & " ,NULL as optlimityn,NULL as optlimitno,NULL as optlimitsold"&VbCRLF
    sqlStr = sqlStr & " ,NULL as optionname,NULL as optaddprice,NULL as optaddbuyprice"&VbCRLF
    sqlStr = sqlStr & " ,NULL as optionTypeName"
    sqlStr = sqlStr & " ,r.outmalloPtCode,r.outmallOptName,r.outmallSellyn,r.outmalllimityn,r.outmalllimitno,r.outmallAddPrice"&VbCRLF
    sqlStr = sqlStr & " from [db_item].[dbo].tbl_wconcept_regedoption r"&VbCRLF
    sqlStr = sqlStr & " 	left join [db_item].[dbo].tbl_item_option o"&VbCRLF
    sqlStr = sqlStr & " 	on o.itemid=r.itemid"&VbCRLF
    sqlStr = sqlStr & " 	and o.itemoption=r.itemoption"&VbCRLF
    sqlStr = sqlStr & " where r.itemid="&itemid&VbCRLF
    sqlStr = sqlStr & " and o.itemid is NULL"&VbCRLF
    sqlStr = sqlStr & " order by Sitename"&VbCRLF
ElseIf mallid = "benepia1010" Then
    sqlStr = ""
    sqlStr = sqlStr & " select top 300 '1' as Sitename,o.itemid,o.itemoption,o.isusing,o.optsellyn"&VbCRLF
    sqlStr = sqlStr & " ,o.optlimityn,o.optlimitno,o.optlimitsold"&VbCRLF
    sqlStr = sqlStr & " ,o.optionname,o.optaddprice,o.optaddbuyprice"&VbCRLF
    sqlStr = sqlStr & " ,o.optionTypeName"&VbCRLF
    sqlStr = sqlStr & " ,'' as outmallOptCode,r.outmallOptName,r.outmallSellyn,r.outmalllimityn,r.outmalllimitno,r.outmallAddPrice"&VbCRLF
    sqlStr = sqlStr & " from [db_item].[dbo].tbl_item_option o"&VbCRLF
    sqlStr = sqlStr & " 	left join [db_item].[dbo].tbl_benepia_regedoption r"&VbCRLF
    sqlStr = sqlStr & " 	on o.itemid=r.itemid"&VbCRLF
    sqlStr = sqlStr & " 	and o.itemoption=r.itemoption"&VbCRLF
    sqlStr = sqlStr & " where o.itemid="&itemid&VbCRLF
    sqlStr = sqlStr & " Union"&VbCRLF
    sqlStr = sqlStr & " select top 100 'Z' as Sitename,r.itemid,r.itemoption,NULL as isusing,NULL as optsellyn"&VbCRLF
    sqlStr = sqlStr & " ,NULL as optlimityn,NULL as optlimitno,NULL as optlimitsold"&VbCRLF
    sqlStr = sqlStr & " ,NULL as optionname,NULL as optaddprice,NULL as optaddbuyprice"&VbCRLF
    sqlStr = sqlStr & " ,NULL as optionTypeName"
    sqlStr = sqlStr & " ,'' as outmallOptCode,r.outmallOptName,r.outmallSellyn,r.outmalllimityn,r.outmalllimitno,r.outmallAddPrice"&VbCRLF
    sqlStr = sqlStr & " from [db_item].[dbo].tbl_benepia_regedoption r"&VbCRLF
    sqlStr = sqlStr & " 	left join [db_item].[dbo].tbl_item_option o"&VbCRLF
    sqlStr = sqlStr & " 	on o.itemid=r.itemid"&VbCRLF
    sqlStr = sqlStr & " 	and o.itemoption=r.itemoption"&VbCRLF
    sqlStr = sqlStr & " where r.itemid="&itemid&VbCRLF
    sqlStr = sqlStr & " and o.itemid is NULL"&VbCRLF
    sqlStr = sqlStr & " order by Sitename"&VbCRLF
Else
    sqlStr = ""
    sqlStr = sqlStr & " select top 300 '1' as Sitename,o.itemid,o.itemoption,o.isusing,o.optsellyn"&VbCRLF
    sqlStr = sqlStr & " ,o.optlimityn,o.optlimitno,o.optlimitsold"&VbCRLF
    sqlStr = sqlStr & " ,o.optionname,o.optaddprice,o.optaddbuyprice"&VbCRLF
    sqlStr = sqlStr & " ,o.optionTypeName"&VbCRLF
    sqlStr = sqlStr & " ,r.outmalloPtCode,r.outmallOptName,r.outmallSellyn,r.outmalllimityn,r.outmalllimitno,r.outmallAddPrice"&VbCRLF
    sqlStr = sqlStr & " from [db_item].[dbo].tbl_item_option o"&VbCRLF
    sqlStr = sqlStr & " 	left join [db_item].[dbo].tbl_OutMall_regedoption r"&VbCRLF
    sqlStr = sqlStr & " 	on r.mallid='"&mallid&"'"&VbCRLF
    sqlStr = sqlStr & " 	and o.itemid=r.itemid"&VbCRLF
    sqlStr = sqlStr & " 	and o.itemoption=r.itemoption"&VbCRLF
    sqlStr = sqlStr & " where o.itemid="&itemid&VbCRLF
    sqlStr = sqlStr & " Union"&VbCRLF
    sqlStr = sqlStr & " select top 100 'Z' as Sitename,r.itemid,r.itemoption,NULL as isusing,NULL as optsellyn"&VbCRLF
    sqlStr = sqlStr & " ,NULL as optlimityn,NULL as optlimitno,NULL as optlimitsold"&VbCRLF
    sqlStr = sqlStr & " ,NULL as optionname,NULL as optaddprice,NULL as optaddbuyprice"&VbCRLF
    sqlStr = sqlStr & " ,NULL as optionTypeName"
    sqlStr = sqlStr & " ,r.outmalloPtCode,r.outmallOptName,r.outmallSellyn,r.outmalllimityn,r.outmalllimitno,r.outmallAddPrice"&VbCRLF
    sqlStr = sqlStr & " from [db_item].[dbo].tbl_OutMall_regedoption r"&VbCRLF
    sqlStr = sqlStr & " 	left join [db_item].[dbo].tbl_item_option o"&VbCRLF
    sqlStr = sqlStr & " 	on o.itemid=r.itemid"&VbCRLF
    sqlStr = sqlStr & " 	and o.itemoption=r.itemoption"&VbCRLF
    sqlStr = sqlStr & " where r.mallid='"&mallid&"'"&VbCRLF
    sqlStr = sqlStr & " and r.itemid="&itemid&VbCRLF
    sqlStr = sqlStr & " and o.itemid is NULL"&VbCRLF
    sqlStr = sqlStr & " order by Sitename"&VbCRLF
End If
rsget.Open sqlStr, dbget
if Not(rsget.EOF or rsget.BOF) then
    arrRows = rsget.getRows
end if
rsget.close

TEN_URI = "http://www.10x10.co.kr/shopping/category_prd.asp?itemid="&itemid
dim isOptSoldOut, isOutOptSoldOut
dim isLimit, isOutLimit
dim limitno, outLimitno

Dim isOptAddPriceExistsItem : isOptAddPriceExistsItem = false

if isArray(arrRows) then
For i =0 To UBound(ArrRows,2)
    isOptAddPriceExistsItem = (isOptAddPriceExistsItem or ArrRows(9,i)>0)
Next
end if
%>
<script language='javascript'>
function saveThis(comp){
    if (confirm('수정 하시겠습니까?')){
        comp.form.submit();
    }
}

function refreshSellStat(itemid,mallid){
    if (mallid=="lotteCom"){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTOCK";
        document.frmSvArr.cksel.value=itemid;
        document.frmSvArr.action = "<%=apiURL%>/outmall/lotteCom/actLotteComReq.asp"
        document.frmSvArr.retFlag.value="retFunc(2)";
        document.frmSvArr.submit();
    }else if (mallid=="cjmall"){
        alert('cjMall 단품 재수신의 경우 5~10분정도 지연됨');
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "QTY";                       //조회
        document.frmSvArr.cksel.value=itemid;
        document.frmSvArr.action = "/admin/etc/cjmall/actCjmallReq.asp"
        document.frmSvArr.retFlag.value="retFunc(2)";
        document.frmSvArr.submit();
   }else if (mallid=="lotteimall"){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "CHKSTOCK";
        document.frmSvArr.cksel.value=itemid;
        document.frmSvArr.action = "<%=apiURL%>/outmall/ltimall/actLotteiMallReq.asp"
        document.frmSvArr.retFlag.value="retFunc(2)";
        document.frmSvArr.submit();
   }
}

function outItemDtlProc(itemid,mallid){
    if (mallid=="lotteCom"){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.cksel.value=itemid;
        document.frmSvArr.action = "<%=apiURL%>/outmall/lotteCom/actLotteComReq.asp"
        document.frmSvArr.retFlag.value="retFunc(1)";
        document.frmSvArr.submit();
    }else if (mallid=="cjmall"){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";                       //가격/한정수량/단품수정
        document.frmSvArr.cksel.value=itemid;
        document.frmSvArr.action = "<%=apiURL%>/outmall/cjmall/actCjmallReq.asp"
        document.frmSvArr.retFlag.value="retFunc(1)";
        document.frmSvArr.submit();
    }else if (mallid=="lotteimall"){
        document.frmSvArr.target = "xLink";
        document.frmSvArr.cmdparam.value = "EDIT";
        document.frmSvArr.cksel.value=itemid;
        document.frmSvArr.action = "<%=apiURL%>/outmall/ltimall/actLotteiMallReq.asp"
        document.frmSvArr.retFlag.value="retFunc(1)";
        document.frmSvArr.submit();
    }else if (mallid=="gsshop"){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT";
		document.frmSvArr.cksel.value=itemid;
		document.frmSvArr.action = "<%=apiURL%>/outmall/gsshop/actgsshopReq.asp"
		document.frmSvArr.submit();
    }else if (mallid=="auction1010"){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT";
		document.frmSvArr.cksel.value=itemid;
		document.frmSvArr.action = "<%=apiURL%>/outmall/auction/actauctionReq.asp"
		document.frmSvArr.submit();
	}else if (mallid=="gmarket1010"){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT";
		document.frmSvArr.cksel.value=itemid;
		document.frmSvArr.action = "<%=apiURL%>/outmall/gmarket/actGmarketReq.asp"
		document.frmSvArr.submit();
    }else if (mallid=="homeplus"){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT";
		document.frmSvArr.cksel.value=itemid;
		document.frmSvArr.action = "<%=apiURL%>/outmall/homeplus/acthomeplusReq.asp"
		document.frmSvArr.submit();
    }else if (mallid=="ssg"){
		document.frmSvArr.target = "xLink";
		document.frmSvArr.cmdparam.value = "EDIT";
		document.frmSvArr.cksel.value=itemid;
		document.frmSvArr.action = "<%=apiURL%>/outmall/ssg/actSsgReq.asp"
		document.frmSvArr.submit();
    } else {
		alert('등록되지 않은 제휴몰입니다.[' + mallid + ']');
	}
}

function retFunc(retval){
    if (retval==1){
        refreshSellStat('<%=itemid%>','<%=mallid%>');
    }else if (retval==2){
        alert('ok');
        document.location.reload();
    }
}
</script>
<table width="100%" cellpadding="5" cellspacing="1" bgcolor="#BBBBBB" class="a">
<form name="frmSv" method="post" action="/admin/etc/popOptionAddPrcSet.asp">
<input type="hidden" name="mode" value="EDTRegType">
<input type="hidden" name="itemid" value="<%=itemid %>">
<input type="hidden" name="mallid" value="<%=mallid %>">
<input type="hidden" name="mngOptAdd" value="<%=mngOptAdd %>">
<tr  bgcolor="#FFFFFF" align="center">
    <td bgcolor="#FFFFFF" align="center" colspan="4"><%= mallid %></td>
    <td align="right"><a href="javascript:document.location.reload();"><img src="http://webadmin.10x10.co.kr/images/icon_reload.gif" border="0"></a></td>
</tr>
<tr  bgcolor="#FFFFFF" align="center">
    <td width="20%" colspan="2"><a href="<%= TEN_URI %>" target=_blank><%= itemid %></a></td>
    <td>
        <%= regitemname %> / <%= lastStatCheckDate %>
    </td>
    <td width="20%" colspan="2"><%= outmallGoodNo %></td>
</tr>
<% if (mngOptAdd="1") then %>
<tr>
    <td bgcolor="#FFFFFF" align="center" colspan="5">
        <input type="radio" name="optAddPrcRegType" value="0" <%=CHKIIF(optAddPrcRegType="0","checked","")%> > 미지정(자동품절)
        <input type="radio" name="optAddPrcRegType" value="1" <%=CHKIIF(optAddPrcRegType="1","checked","")%> > 옵션추가금액 없는상품만 판매
        <input type="radio" name="optAddPrcRegType" value="9" disabled > 옵션추가금액 별도 등록
    </td>
</tr>
<tr>
    <td bgcolor="#FFFFFF" align="center" colspan="5">
    <input type="button" value=" 저 장 " onClick="saveThis(this);">
    </td>
</tr>
<% end if %>
</form>
</table>
<p>
<table width="100%" cellpadding="5" cellspacing="1" bgcolor="#BBBBBB" class="a">
<tr align="center" bgcolor="#F3F3FF" height="20">
    <td colspan="7">10x10</td>
    <td width="1" bgcolor="#FFFFFF">
   	<% If mallid <> "interpark" Then %>
    <input type="button" value=">>" onClick="outItemDtlProc('<%=itemid %>','<%=mallid %>');">
    <% End If %>
    </td>
    <td colspan="4"><%= mallid %> <% If mallid <> "gsshop" and mallid <> "interpark" and mallid <> "auction1010" and mallid <> "gmarket1010" and mallid <> "homeplus" Then %> <input type="button" value="단품재수신" class="button" onClick="refreshSellStat('<%=itemid %>','<%=mallid %>')"> <% End If %>
    <br><br>(복합옵션은 정확하지 않음)
    </td>
</tr>
<tr align="center" bgcolor="#F3F3FF" height="20">
    <td>상품코드</td>
    <td>옵션코드</td>
    <td>옵션타입</td>
    <td>옵션명</td>
    <td>한정</td>
    <td>판매</td>
    <td>옵션추가액</td>
    <td width="1" bgcolor="#FFFFFF">

    </td>
    <td>옵션명</td>
    <td>옵션코드</td>
    <td>한정</td>
    <td>판매</td>
</tr>
<% if isArray(arrRows) then %>
<% For i =0 To UBound(ArrRows,2) %>
<%
    if isNULL(ArrRows(3,i)) then
        isOptSoldOut = false
    else
        isOptSoldOut = ((ArrRows(3,i)="N") or (ArrRows(4,i)="N") or (((ArrRows(5,i)="Y") and (ArrRows(6,i)-ArrRows(7,i)<1))))
    end if

    if isNULL(ArrRows(5,i)) then
        isLimit = false
        limitno = 0
    else
        isLimit = (ArrRows(5,i)="Y")
        limitno = (ArrRows(6,i)-ArrRows(7,i))

    end if

    if (limitno<1) then limitno=0

%>
<%
    if isNULL(ArrRows(14,i)) or isNULL(ArrRows(15,i)) or isNULL(ArrRows(16,i)) then
        isOutOptSoldOut = false
    else
        isOutOptSoldOut = (ArrRows(14,i)="N") or ((ArrRows(15,i)="Y") and (ArrRows(16,i)<1))
    end if

    if isNULL(ArrRows(15,i)) then
        isOutLimit = false
        outLimitno = 0
    else
        isOutLimit = (ArrRows(15,i)="Y")
        outLimitno = ArrRows(16,i)

    end if

    if (outLimitno<1) then outLimitno=0
%>
<tr align="center" bgcolor="#FFFFFF" height="20">
    <td bgcolor="<%= CHKIIF(isOptSoldOut,"#EEEEEE","#FFFFFF") %>" ><%= ArrRows(1,i) %></td>
    <td bgcolor="<%= CHKIIF(isOptSoldOut,"#EEEEEE","#FFFFFF") %>" ><%= ArrRows(2,i) %></td>
    <td bgcolor="<%= CHKIIF(isOptSoldOut,"#EEEEEE","#FFFFFF") %>" ><%= ArrRows(11,i) %></td>
    <td bgcolor="<%= CHKIIF(isOptSoldOut,"#EEEEEE","#FFFFFF") %>" ><%= ArrRows(8,i) %></td>
    <td bgcolor="<%= CHKIIF(isOptSoldOut,"#EEEEEE","#FFFFFF") %>" >
        <% if isLimit then %>
        <font color="blue"><%= limitno %></font>
        <% end if %>
    </td>
    <td bgcolor="<%= CHKIIF(isOptSoldOut,"#EEEEEE","#FFFFFF") %>" >
        <%= CHKIIF(isOptSoldOut,"<font color=red>품절</font>","") %>
    </td>
    <td bgcolor="<%= CHKIIF(isOptSoldOut,"#EEEEEE","#FFFFFF") %>" ><%= ArrRows(9,i) %></td>
    <td width="1" bgcolor="#FFFFFF"></td>
    <td bgcolor="<%= CHKIIF(isOutOptSoldOut,"#EEEEEE","#FFFFFF") %>" ><%= ArrRows(13,i) %></td>
    <td bgcolor="<%= CHKIIF(isOutOptSoldOut,"#EEEEEE","#FFFFFF") %>" ><%= ArrRows(12,i) %></td>
    <td bgcolor="<%= CHKIIF(isOutOptSoldOut,"#EEEEEE","#FFFFFF") %>" >
        <% if isOutLimit then %>
        <font color="blue"><%= outLimitno %></font>
        <% end if %>
    </td>
    <td bgcolor="<%= CHKIIF(isOutOptSoldOut,"#EEEEEE","#FFFFFF") %>" ><%= CHKIIF(isOutOptSoldOut,"<font color=red>품절</font>","") %></td>
</tr>
<% next %>
<% end if %>

</table>

<% if (mngOptAdd<>"1") then %>
<p>
<table width="100%" cellpadding="5" cellspacing="1" bgcolor="#BBBBBB" class="a">
<tr>
    <td align="center" bgcolor="#FFFFFF" height="20">
    <input type="button" value=" 닫기 " onClick="self.close();">
    </td>
</tr>
</table>
<% end if %>

<form name="frmSvArr" method="post" >
<input type="hidden" name="mode" value="">
<input type="hidden" name="cmdparam" value="">
<input type="hidden" name="cksel" value="">
<input type="hidden" name="retFlag" value="">
</form>
<iframe name="xLink" id="xLink" frameborder="1" width="100%" height="100"></iframe>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
