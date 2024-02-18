<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" --> 
<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/partners/partnerusercls.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%

dim makerid, cdl, cdm, cds, d, i, page, mstart, OnlySellyn, OnlyIsUsing, onlyOutItem, onlyOldItem, mwdiv, danjongyn, limityn
dim research, ChulgoNo, TurnOverPro, yyyy1, mm1, yyyy2, mm2, monthgubun, excBaseRegItem, dispCate
	dispCate = requestCheckvar(request("disp"),16)
	makerid = request("makerid")
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	cdl = request("cdl")
	cdm = request("cdm")
	cds = request("cds")
	page = request("page")
	research = request("research")
	OnlySellyn = request("OnlySellyn")
	OnlyIsUsing = request("OnlyIsUsing")
	onlyOutItem = request("onlyOutItem")
	onlyOldItem = request("onlyOldItem")
	mwdiv       = request("mwdiv")
	danjongyn   = request("danjongyn")
	limityn     = request("limityn")
	ChulgoNo    = request("ChulgoNo")
	TurnOverPro = request("TurnOverPro")
	monthgubun = request("monthgubun")
	excBaseRegItem = request("excBaseRegItem")

''if (research="") and (OnlyIsUsing="") then OnlyIsUsing="Y"
if (research="") and (onlyOutItem="") then onlyOutItem="on"
if (research="") and (onlyOldItem="") then onlyOldItem="on"
if (research="") and (mwdiv="") then mwdiv="MW"
if (research="") and (danjongyn="") then danjongyn="SN"
if (research="") and (excBaseRegItem="") then excBaseRegItem="Y"

if (ChulgoNo="") then ChulgoNo="5"
if (TurnOverPro="") then TurnOverPro="0.5"

if (page = "") then
        page = 1
end if

if (yyyy1 = "") then
	d = CStr(dateadd("m" ,-1, now()))
	yyyy1 = Left(d,4)
	mm1 = Mid(d,6,2)

	yyyy2 = yyyy1
	mm2   = mm1
end if

dim olistforout
set olistforout = new CSummaryItemStock
	olistforout.FRectYYYYMM = yyyy1 + "-" + mm1
	olistforout.FRectEndDate = yyyy2 + "-" + mm2
	olistforout.FRectMakerid = makerid
	olistforout.FPageSize = 5000
	olistforout.FCurrPage = page
	olistforout.FRectCD1 = cdl
	olistforout.FRectCD2 = cdm
	olistforout.FRectCD3 = cds
	olistforout.FRectOnlySellyn = OnlySellyn
	olistforout.FRectOnlyIsUsing = OnlyIsUsing
	olistforout.FRectOnlyOldItem = onlyOldItem
	olistforout.FRectOnlyOutItem = OnlyOutItem
	olistforout.FRectMwDiv = mwdiv
	olistforout.FRectDanjongyn =danjongyn
	olistforout.FRectLimityn =limityn
	olistforout.FRectChulgoNo   = ChulgoNo
	olistforout.FRectTurnOverPro = TurnOverPro
	olistforout.FRectExcBaseRegItem = excBaseRegItem
	olistforout.FRectMonthGubun = monthgubun
	olistforout.FRectDispCate		= dispCate

	if (makerid<>"") then
	    olistforout.GetItemListTurnOver
	else
	    olistforout.GetBrandListTurnOver
	end if

dim bufStr,strMW
 
'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=itemlist_"&page&".csv"
Response.CacheControl = "public"

response.write "상품코드,브랜드ID,상품명(옵션),거래구분,총출고량,ON출고,OFF출고,출고합계,월말재고,회전율,판매,사용,한정,단종" & VbCrlf
  
	if olistforout.FresultCount > 0 then 
  	for i=0 to olistforout.FresultCount-1 
  	
		bufStr = ""  
		bufStr = bufStr & olistforout.FItemList(i).FItemID
		bufStr = bufStr & "," & olistforout.FItemList(i).Fmakerid
		bufStr = bufStr & "," & replace(db2html(olistforout.FItemList(i).Fitemname),","," ")
		if olistforout.FItemList(i).FitemoptionName <> "" then
		bufStr = bufStr & " [" & replace(olistforout.FItemList(i).FitemoptionName,",","/") & "]"
		end if
		bufStr = bufStr & "," & mwdivName(olistforout.FItemList(i).Fmwdiv)
		bufStr = bufStr & "," & olistforout.FItemList(i).Faccumchulgo*(-1)
		bufStr = bufStr & "," & olistforout.FItemList(i).Fsellno*(-1)

		bufStr = bufStr & "," & olistforout.FItemList(i).Foffchulgono*(-1)
		bufStr = bufStr & "," & (olistforout.FItemList(i).Fsellno + olistforout.FItemList(i).Foffchulgono)*(-1)
		bufStr = bufStr & "," & olistforout.FItemList(i).Frealstock
		if olistforout.FItemList(i).Frealstock<>0 then
		bufStr = bufStr & "," & CLng((olistforout.FItemList(i).Fsellno+olistforout.FItemList(i).Foffchulgono)*-1/olistforout.FItemList(i).Frealstock*100)/100
		else
		bufStr = bufStr & ","
		end if
		bufStr = bufStr & "," & olistforout.FItemList(i).Fsellyn
		bufStr = bufStr & "," & olistforout.FItemList(i).Fisusing
		bufStr = bufStr & "," & olistforout.FItemList(i).Flimityn
		if (olistforout.FItemList(i).Flimityn = "Y") then
		bufStr = bufStr & " (" & olistforout.FItemList(i).GetLimitStr & ")"
		end if
		if olistforout.FItemList(i).Fdanjongyn="Y" then
			strMW = "단종"
		elseif olistforout.FItemList(i).Fdanjongyn="S" then
			strMW = "재고부족"
		elseif olistforout.FItemList(i).Fdanjongyn="M" then
			strMW = "MD품절"
		end if
		bufStr = bufStr & "," & strMW

		

    	response.write bufStr & VbCrlf	
    Next 
  end if  

 
SET olistforout = Nothing
%>
 
 
 
<!-- #include virtual="/lib/db/dbclose.asp" -->