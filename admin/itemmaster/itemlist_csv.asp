<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" --> 
<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
' 사용안함
dim itemid, itemname, makerid, sellyn, usingyn, danjongyn, mwdiv, limityn, vatyn, sailyn, overSeaYn, itemdiv
dim cdl, cdm, cds, showminusmagin, marginup, margindown, dispCate
dim page
dim infodivYn, infodiv, deliverytype, sortDiv

itemid      = requestCheckvar(request("itemid"),255)
itemname    = requestCheckvar(request("itemname"),64)
makerid     = requestCheckvar(request("makerid"),32)
sellyn      = requestCheckvar(request("sellyn"),10)
usingyn     = requestCheckvar(request("usingyn"),10)
danjongyn   = requestCheckvar(request("danjongyn"),10)
mwdiv       = requestCheckvar(request("mwdiv"),10)
limityn     = requestCheckvar(request("limityn"),10)
vatyn       = requestCheckvar(request("vatyn"),10)
sailyn      = requestCheckvar(request("sailyn"),10)
overSeaYn   = requestCheckvar(request("overSeaYn"),10)
itemdiv     = requestCheckvar(request("itemdiv"),10)
deliverytype= requestCheckvar(request("deliverytype"),10)
page = requestCheckvar(request("page"),10)
cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)
dispCate = requestCheckvar(request("disp"),16)

showminusmagin = request("showminusmagin")
marginup = request("marginup")
margindown = request("margindown")
sortDiv	= requestCheckvar(request("sortDiv"),5)
if sortDiv="" then sortDiv="new"

infodiv  = request("infodiv")
infodivYn  = requestCheckvar(request("infodivYn"),10)

If infodiv <> "" Then
	infodivYn = "Y"	
End If

If marginup <> "" AND IsNumeric(marginup) = False Then
	rw "<script>alert('마진값(이상)이 잘못되었습니다. - "&marginup&"');history.back();</script>"
	dbget.close()
	Response.End
End If

If margindown <> "" AND IsNumeric(margindown) = False Then
	rw "<script>alert('마진값(이하)이 잘못되었습니다. - "&margindown&"');history.back();</script>"
	dbget.close()
	Response.End
End If



if (page="") then page=1

if itemid<>"" then
	dim iA ,arrTemp,arrItemid
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


'==============================================================================
dim oitem

set oitem = new CItem

oitem.FPageSize         = 5000
oitem.FCurrPage         = page
oitem.FRectMakerid      = makerid
oitem.FRectItemid       = itemid
oitem.FRectItemName     = itemname

oitem.FRectSellYN       = sellyn
oitem.FRectIsUsing      = usingyn
oitem.FRectDanjongyn    = danjongyn
oitem.FRectLimityn      = limityn
oitem.FRectMWDiv        = mwdiv
oitem.FRectVatYn        = vatyn
oitem.FRectSailYn       = sailyn
oitem.FRectIsOversea	= overSeaYn

oitem.FRectCate_Large   = cdl
oitem.FRectCate_Mid     = cdm
oitem.FRectCate_Small   = cds
oitem.FRectDispCate		= dispCate
oitem.FRectItemDiv      = itemdiv

oitem.FRectMinusMigin = showminusmagin
oitem.FRectMarginUP = marginup
oitem.FRectMarginDown = margindown
oitem.FRectInfodivYn    = infodivYn
oitem.FRectInfodiv    = infodiv 
oitem.FRectDeliverytype = deliverytype
oitem.FRectSortDiv		= sortDiv

oitem.GetItemList

dim i,bufStr
 
'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=상품목록_"&page&".csv"
Response.CacheControl = "public"

response.write "상품코드,브랜드,상품명,소비자가, 매입가, 마진, 할인율, 할인가, 할인매입가, 할인마진,쿠폰할인율,쿠폰적용판매가, 쿠폰적용매입가,쿠폰적용마진, 거래구분,배송구분, 판매여부, 사용여부, 한정여부,한정수량" & VbCrlf
  
	if oitem.FresultCount > 0 then 
  	for i=0 to oitem.FresultCount-1 
  			bufStr = ""  
			 	bufStr = bufStr & oitem.FItemList(i).Fitemid
			 	bufStr = bufStr & "," & oitem.FItemList(i).Fmakerid
			 	bufStr = bufStr & "," & replace(db2html(oitem.FItemList(i).Fitemname),","," ")
			 	bufStr = bufStr & "," & oitem.FItemList(i).Forgprice
			 	bufStr = bufStr & "," & oitem.FItemList(i).Forgsuplycash
			 	bufStr = bufStr & "," &  fnPercent(oitem.FItemList(i).Forgsuplycash,oitem.FItemList(i).Forgprice,1)
			 	 	if oitem.FItemList(i).Fsailyn="Y" then
			 		bufStr = bufStr & "," & CLng((oitem.FItemList(i).Forgprice-oitem.FItemList(i).Fsailprice)/oitem.FItemList(i).Forgprice*100) & "%"
				else
					bufStr = bufStr & ",0%"  
			  end if
			 	if oitem.FItemList(i).Fsailyn="Y" then
			 		bufStr = bufStr & "," & oitem.FItemList(i).Fsailprice
				else
					bufStr = bufStr & ","  
			  end if
			 		if oitem.FItemList(i).Fsailyn="Y" then
			 		bufStr = bufStr & "," &  oitem.FItemList(i).Fsailsuplycash
				else
					bufStr = bufStr & ","  
			  end if
			 	if oitem.FItemList(i).Fsailyn="Y" then
			 		bufStr = bufStr & "," & fnPercent(oitem.FItemList(i).Fsailsuplycash,oitem.FItemList(i).Fsailprice,1)
				else
					bufStr = bufStr & ","  
			  end if
				' 쿠폰할인율
				if oitem.FItemList(i).FitemCouponYn="Y" then
					if oitem.FItemList(i).FitemCouponType =1 or oitem.FItemList(i).FitemCouponType =2 then			
						bufStr = bufStr & "," & CLng((oitem.FItemList(i).Forgprice-oitem.FItemList(i).GetCouponAssignPrice)/oitem.FItemList(i).Forgprice*100) & "%"
					else
						bufStr = bufStr & ",0%"
					end if
				else
					bufStr = bufStr & ",0%"
				end if
			  if oitem.FItemList(i).FitemCouponYn="Y" then
				 	if oitem.FItemList(i).FitemCouponType =1 or oitem.FItemList(i).FitemCouponType =2 then 
						bufStr = bufStr & "," & oitem.FItemList(i).GetCouponAssignPrice()   
					else
						bufStr = bufStr & ","  		
					end if
				else 
					bufStr = bufStr & ","  
				end if
			 	if oitem.FItemList(i).FitemCouponYn="Y" then
					if oitem.FItemList(i).FitemCouponType="1" or oitem.FItemList(i).FitemCouponType="2" then
						if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then
								bufStr = bufStr & "," & oitem.FItemList(i).Forgsuplycash 
						else
								bufStr = bufStr & "," & oitem.FItemList(i).Fcouponbuyprice 
						end if
					else	
						bufStr = bufStr & ","  
					end if 
				else	
					bufStr = bufStr & "," 
				end if
				if oitem.FItemList(i).FitemCouponYn="Y" then
				Select Case oitem.FItemList(i).FitemCouponType
					Case "1"
						if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then
							bufStr = bufStr & "," & fnPercent(oitem.FItemList(i).Forgsuplycash,oitem.FItemList(i).GetCouponAssignPrice(),1) 
						else
							bufStr = bufStr & "," & fnPercent(oitem.FItemList(i).Fcouponbuyprice,oitem.FItemList(i).GetCouponAssignPrice(),1) 
						end if
					Case "2"
						if oitem.FItemList(i).Fcouponbuyprice=0 or isNull(oitem.FItemList(i).Fcouponbuyprice) then
							bufStr = bufStr & "," & fnPercent(oitem.FItemList(i).Forgsuplycash,oitem.FItemList(i).GetCouponAssignPrice(),1)  
						else
							bufStr = bufStr & "," & fnPercent(oitem.FItemList(i).Fcouponbuyprice,oitem.FItemList(i).GetCouponAssignPrice(),1)  
						end if
					CASE Else
							bufStr = bufStr & "," 	
					end Select
				else
						bufStr = bufStr & "," 
				end if
			 	bufStr = bufStr & "," & mwdivName(oitem.FItemList(i).Fmwdiv)
				bufStr = bufStr & "," & getBeadalDivname(oitem.FItemList(i).Fdeliverytype)
			 	bufStr = bufStr & "," & oitem.FItemList(i).Fsellyn
			 	bufStr = bufStr & "," & oitem.FItemList(i).Fisusing
			 	bufStr = bufStr & "," & oitem.FItemList(i).Flimityn
			 	if  oitem.FItemList(i).Flimityn ="Y" then
			 	bufStr = bufStr & "," & (oitem.FItemList(i).Flimitno-oitem.FItemList(i).Flimitsold )
				else
				bufStr = bufStr & ","  	
				end if
    	response.write bufStr & VbCrlf	
    Next 
  end if  

 
SET oitem = Nothing
%>
 
 
 
<!-- #include virtual="/lib/db/dbclose.asp" -->