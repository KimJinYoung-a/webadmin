 <%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ���� ���� ��ǰ�ٿ�
' History : 2016.07.25 ������ ���� 
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->
<!-- #include virtual="/lib/function.asp"--> 
<!-- #include virtual="/lib/classes/items/itemsalecls.asp"-->
<!-- #include virtual="/lib/classes/event/eventmanageCls.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->

<%dim eCode, sCode 
	dim iCurrPage,iPageSize,iTotalPage,iTotCnt
	dim makerid, sailyn,invalidmargin, sRectItemidArr  
	dim  bufStr 
	dim arrList, intLoop
	dim clsSaleItem
	dim isRate, isMargin,isMValue
	dim sSalestatus, sItemSale
	iPageSize = 1000
	iCurrpage = 1
	eCode     = requestCheckVar(Request("eC"),10)
	sCode     = requestCheckVar(Request("sC"),10) 
	makerid =  requestCheckVar(Request("makerid"),32)
	sailyn	=  requestCheckVar(Request("sailyn"),1)
 	isRate=  requestCheckVar(Request("iSR"),10)
 	isMargin=  requestCheckVar(Request("salemargin"),10)
 	isMValue=  requestCheckVar(Request("isMValue"),10)
	invalidmargin=  requestCheckVar(Request("invalidmargin"),1)
	sRectItemidArr=  requestCheckVar(Request("sRectItemidArr"),400)
	sSalestatus 	= requestCheckVar(Request("salestatus"),4)
  sItemSale	= requestCheckVar(Request("selItemStatus"),4)

	if sRectItemidArr<>"" then
	dim iA ,arrTemp,arrItemid
	sRectItemidArr = replace(sRectItemidArr,",",chr(10)) 
	sRectItemidArr = replace(sRectItemidArr,chr(13),"") 
	arrTemp = Split(sRectItemidArr,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp) 
		if trim(arrTemp(iA))<>"" then 
			'��ǰ�ڵ� ��ȿ�� �˻�(2008.08.05;������)
			if Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
				dbget.close()	:	response.End
			else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			end if
		end if
		iA = iA + 1
	loop

	sRectItemidArr = left(arrItemid,len(arrItemid)-1)
end if
 
 '���� ��ǰ����
	set clsSaleItem = new CSaleItem 
	clsSaleItem.FCPage = iCurrpage
	clsSaleItem.FPSize = iPageSize
	clsSaleItem.FSCode = sCode
	clsSaleItem.FRectMakerid = makerid
	clsSaleItem.FRectsailyn = sailyn
	clsSaleItem.FRectinvalidmargin =invalidmargin
	clsSaleItem.FRectItemidArr = sRectItemidArr 
	clsSaleItem.FRectSaleStatus = sSalestatus
  clsSaleItem.FRectItemSaleStatus = sItemSale
  
	arrList = clsSaleItem.fnGetSaleItemList
	iTotCnt = clsSaleItem.FTotCnt	'��ü ������  ��
	 
	
	'���Ⱓ�� ��ǰ���� ���� ����
	Dim arrItemCoupon, iclp
	arrItemCoupon = clsSaleItem.fnGetCouponListBySaleInfo
	set clsSaleItem = nothing
	  
	 '�������¿� ���� ���԰� ����-------------------------------------------------------
Function fnSetSaleSupplyPrice(ByVal MarginType, ByVal MarginValue, ByVal orgPrice, ByVal orgSupplyPrice, ByVal salePrice)
	Dim orgMRate
	if orgPrice <>0 then '�� ������
		orgMRate = 100-fix(orgSupplyPrice/orgPrice*10000)/100
	end if

	SELECT CASE MarginType
		Case 1	'���ϸ���
			fnSetSaleSupplyPrice = salePrice- fix(salePrice*(orgMRate/100))
		Case 2	'��ü�δ�
			fnSetSaleSupplyPrice = salePrice-(orgPrice-orgSupplyPrice)
		Case 3	'�ݹݺδ�
			fnSetSaleSupplyPrice = orgSupplyPrice- fix((orgPrice-salePrice)/2)
		Case 4	'10x10�δ�
			fnSetSaleSupplyPrice = orgSupplyPrice
		Case 5	'��������
			fnSetSaleSupplyPrice = salePrice - fix(salePrice*(MarginValue/100))
	END SELECT
End Function
'-----------------------------------------------------------------------------------   
Dim arrsalemargin, arrsalestatus
	arrsalemargin = fnSetCommonCodeArr("salemargin",False)
	arrsalestatus= fnSetCommonCodeArr("salestatus",False)
 
'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=���λ�ǰ���_"&sCode&".csv"
Response.CacheControl = "public"

response.write "��ǰ�ڵ�,�귣��,��ǰ��,��౸��,���λ���,�����ǸŰ�,������԰�,���縶����,���Һ��ڰ�,�����԰�,��������,�Һ��ڰ����������,�����ǸŰ�,���θ��԰�,���θ�����" & VbCrlf

 
	%>
		<%	Dim mSPrice, mSBPrice, iSaleMargin, iOrgMargin, iSalePercent
			Dim cpSP, cpSB, cpSM, strCpDesc, strCpList
			dim mOrgSailPrice, mOrgSailSuplyCash, sOrgSailYn, iOrgSailMargin
			dim strMW
			
			iSaleMargin=0
			iOrgMargin = 0
			iOrgSailMargin= 0
			
			Function numBerBurim(idx, sosu)
				Dim tmpSu
				tmpSu = FormatNumber(idx - 0.5/10^sosu, sosu)
				If cstr(int(tmpSu)) = cstr(formatnumber(tmpSu,0)) Then
					numBerBurim = formatnumber(tmpSu,0)
				Else
					numBerBurim = tmpSu
				End If
			End Function

		IF isArray(arrList) THEN
		 	
			For intLoop = 0 To UBound(arrList,2)
			mSPrice  =arrList(13,intLoop) - (arrList(13,intLoop)*(isRate/100))
			mSBPrice = fnSetSaleSupplyPrice(isMargin,isMValue,arrList(13,intLoop),arrList(14,intLoop),mSPrice)
	 
			if mSPrice<>0 then iSaleMargin =  100-fix(mSBPrice/mSPrice*10000)/100
			 if arrList(13,intLoop)<>0 then 
			 	iOrgMargin= 100-fix(arrList(14,intLoop)/arrList(13,intLoop)*10000)/100
			 	iSalePercent =numBerBurim(((arrList(13,intLoop)-arrList(2,intLoop))/arrList(13,intLoop))*100,2)
			 end if
			 		
			'���� ���λ�ǰ�� ��� ���� ���ΰ��� �������� 
			'���� ���λ�ǰ�� ��쿡�� ������ �� �����ǸŰ�.���԰� ����� ���Һ��ڰ������� �Ѵ�
			sOrgSailYn = arrList(24,intLoop) 
			mOrgSailPrice = arrList(22,intLoop)
			mOrgSailSuplycash = arrList(23,intLoop) 
			if mOrgSailPrice <>0 then 
			 	iOrgSailMargin= 100-fix(mOrgSailSuplycash/mOrgSailPrice*10000)/100 
			 end if 
			 
			cpSP=0: cpSB=0: cpSM=0: strCpDesc="": strCpList=""
			if isArray(arrItemCoupon) then

				for icLp=0 to ubound(arrItemCoupon,2)
					if cStr(arrItemCoupon(4,icLp))=cStr(arrList(1,intLoop)) then
						'��ǰ�����ǸŰ�
						Select Case arrItemCoupon(1,icLp)
							Case "1"
								cpSP = mSPrice- CLng(arrItemCoupon(2,icLp)*mSPrice/100)
							Case "2"
								cpSP = mSPrice- arrItemCoupon(2,icLp)
							Case Else
								cpSP = mSPrice
						End Select
						'��ǰ�������԰�
						cpSB = arrItemCoupon(5,icLp)
			 				'��ǰ��������
						if cpSB>0 then
							 cpSM = formatNumber(100-fix(cpSB/cpSP*10000)/100,0)
						end if	
						
						strCpList = strCpList & "[" & arrItemCoupon(0,icLp) & "] "
					end if
				next	
				
				if strCpList<>"" then
					strCpDesc = "( ��ǰ����"&strCpList&")"
				end if	
			end if
			
			strMW = ""	
				if  arrList(17,intLoop) ="U" then
					strMW = "��ü"
				elseif arrList(17,intLoop) ="M" then
					strMW = "����"
				elseif arrList(17,intLoop) ="W" then 
					strMW = "��Ź"
				end if
 
	 
			  bufStr = ""  
			 	bufStr = bufStr & arrList(1,intLoop) 
        bufStr = bufStr & "," &   db2html(arrList(7,intLoop))                                                                     
        bufStr = bufStr & "," &   db2html(arrList(8,intLoop))                                                                     
        bufStr = bufStr & "," &  strMW                                                                   
        bufStr = bufStr & "," &    arrList(10,intLoop)&" "&fnGetCommCodeArrDesc(arrsalestatus,arrList(4,intLoop)) &" "& chkIIF(strCpDesc>"",strCpDesc,"") 
        bufStr = bufStr & "," &    arrList(11,intLoop) 
        bufStr = bufStr & "," &  arrList(12,intLoop)
         if arrList(11,intLoop)<>0 then                                                                   
        bufStr = bufStr & "," &   100-fix(arrList(12,intLoop)/arrList(11,intLoop)*10000)/100 &"%"
      		else
      	bufStr = bufStr & "," &"0"
					end if                                                            
        bufStr = bufStr & "," &  arrList(13,intLoop) 
         if sOrgSailYn ="Y" then 
			  bufStr = bufStr &"("&(arrList(13,intLoop)-mOrgSailPrice)/arrList(13,intLoop)*100 &" %��)"& mOrgSailPrice 
			    end if 			    	                                                                      
        bufStr = bufStr & "," &  arrList(14,intLoop) 
        	if sOrgSailYn ="Y" then 
			  bufStr = bufStr & mOrgSailSuplycash 
			  	end if                                                                      
        bufStr = bufStr & "," &  iOrgMargin &"%"
			    if sOrgSailYn ="Y" then 
			  bufStr = bufStr & iOrgSailMargin &"%"
			    end if        
			  bufStr = bufStr & "," & iSalePercent &"%"                           
        bufStr = bufStr & "," &   arrList(2,intLoop)                                                                    
        bufStr = bufStr & "," &    arrList(3,intLoop)   
        	if arrList(2,intLoop)<>0 then                                                                  
        bufStr = bufStr & "," &   100-fix(arrList(3,intLoop)/arrList(2,intLoop)*10000)/100&"%"                                                                          
      		else
      	bufStr = bufStr & "," &"0"
    			end if
    		response.write bufStr & VbCrlf	
     NEXT                                                                                              
	 END IF
	        

   
       %> 
<!-- #include virtual="/lib/db/dbclose.asp" -->
