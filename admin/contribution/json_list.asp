<%@Language="VBScript" CODEPAGE="65001" %>
<% option explicit %>
<%
Response.CharSet="utf-8" 
Response.codepage="65001"
Response.ContentType="text/html;charset=utf-8"

%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/db/dbSTSopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"--> 
<!-- #include virtual="/lib/classes/contribution/contributionCls.asp"--> 
<%
dim sdatetype, dstdate,deddate , sdispcate,catekind
Dim  syear,smonth, sday, eday   
dim itotCnt
dim arrcate, intc 
dim clsCMeachulLog 

    sdatetype     = requestCheckvar(request("dategbn"),32)
    sdispcate = requestCheckvar(request("blnDp"),1) 
    catekind =  requestCheckvar(request("rdoCK"),1) 
   	syear     = requestcheckvar(request("sy"),4)
	smonth     = requestcheckvar(request("sM"),2)
    sday     = requestcheckvar(request("sd"),2)
    eday     = requestCheckvar(request("ed"),2)  

 if sdatetype="" then sdatetype="ipkumdate" 
 if  sdispcate ="" then  sdispcate = "N" 
if syear ="" then  syear = Cstr(Year( dateadd("m",-1,date()) ))
if smonth ="" then smonth = Cstr(Month( dateadd("m",-1,date()) ))
if sday ="" then sday =  "01" 
  dstdate = DateSerial(syear, smonth, sday) 
if eday ="" then  eday =  Cstr(Day( dateadd("d",-1,DateSerial(Year(Date()), Month(Date()), 1)) )) 
 deddate = DateSerial(syear, smonth, eday)    
 
 if dateadd("m",1,dstdate) <= deddate then deddate = dateadd("d",dateadd("m",1,dstdate),-1)   
 
 
dim i, tmpJSON, j,m
dim gubunNM
dim buyPer, BCPer, MPPer, MPerPer
dim hidx_buy1,hidx_buy2 ,hidx_buy3
dim hidx_bc1, hidx_bc2
dim hidx_MP1, hidx_MP2
dim hidx_MPPer1, hidx_MPPer2  
dim mper
dim imax
dim catecnt
dim buyVar, buyVarCate, itemVar, odsVar, itemVarcate
Set tmpJSON = New aspJSON 
With tmpJSON.data 
	set clsCMeachulLog = new CMeachulLog
	clsCMeachulLog.FdateType =sdatetype
	clsCMeachulLog.FstDate =dstdate
	clsCMeachulLog.FedDate =deddate
	clsCMeachulLog.FDispCate =sdispcate
    clsCMeachulLog.Fcatekind =catekind
	clsCMeachulLog.fnGetOrerLogData

    if sdispcate ="Y" then
        iTotCnt = clsCMeachulLog.FCateRow
        arrcate =clsCMeachulLog.fnGetCateList
        if isarray(arrcate) then
         catecnt =ubound(arrcate,2) +1 
        else
        catecnt =0
        end if 
    else
	    iTotCnt = clsCMeachulLog.FTotCnt
    end if
    j = 0
	if iTotCnt > 0 then	  
        '------------------구매총액 ----------------------------------------------------
            '구매총액 
            .Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", -1 
			.Add "구분", "구매총액"
            .Add "전체_주문건수", clsCMeachulLog.FOds_sum
            .Add "전체_상품수량", clsCMeachulLog.FItem_sum
			.Add "전체_금액", clsCMeachulLog.Fbuy_sum
			.Add "전체_율",  ""
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2)  
                .Add arrcate(1,intC)&"_상품수량", clsCMeachulLog.FcateItem_Sum(arrcate(3,intc))
			    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.Fcatebuy_Sum(arrcate(3,intc))
                .Add arrcate(1,intC)&"_율", ""
                next
             end if   
			end if
			End With     
            hidx_buy1 = J
			j = j+ 1  

            
            '10x10 구매총액
            i=0
            
			.Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_buy1
			.Add "구분", "10x10" 
             .Add "전체_주문건수", clsCMeachulLog.FOds_10
            .Add "전체_상품수량", clsCMeachulLog.FItem_10
			.Add "전체_금액", clsCMeachulLog.Fbuy_10
			.Add "전체_율", ""
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2) 
                .Add arrcate(1,intC)&"_상품수량", clsCMeachulLog.FcateItem_10(arrcate(3,intc))
			    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.Fcatebuy_10(arrcate(3,intc))
                .Add arrcate(1,intC)&"_율", ""
                next
             end if   
			end if
			End With    
            hidx_buy2 = J
			j = j+ 1
           m = 0
            
          '10x10 상품 구매총액
            i=0 
			.Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_buy2
			.Add "구분", "상품" 
            .Add "전체_주문건수", clsCMeachulLog.FOds_10I
            .Add "전체_상품수량", clsCMeachulLog.FItem_10I
			.Add "전체_금액", clsCMeachulLog.Fbuy_10I
			.Add "전체_율", ""
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2) 
                .Add arrcate(1,intC)&"_상품수량", clsCMeachulLog.FcateItem_10I(arrcate(3,intc))
			    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.Fcatebuy_10I(arrcate(3,intc))
                .Add arrcate(1,intC)&"_율", ""
                next
             end if   
			end if
			End With    
            hidx_buy3 = J
			j = j+ 1
           m = 0
          
         For i = 0 to iTotCnt 
            if i > 0 then
             '제휴 구매총액
                if clsCMeachulLog.Fsitename(i) <> clsCMeachulLog.Fsitename(i-1)  then
                
                    .Add j, tmpJSON.Collection()
                    With .item(j) 
                    .Add "ID", j 
                    .Add "Head_ID", hidx_buy1
                    .Add "구분",  "제휴" 
                    .Add "전체_주문건수", clsCMeachulLog.FOds_P
                    .Add "전체_상품수량", clsCMeachulLog.FItem_P
                    .Add "전체_금액", clsCMeachulLog.Fbuy_P
                    .Add "전체_율", ""
                    if sdispcate ="Y" then
                        if isArray(arrcate) then 
                            for intC= 0 To ubound(arrcate,2) 
                            .Add arrcate(1,intC)&"_상품수량", clsCMeachulLog.FcateItem_P(arrcate(3,intc))
                            .Add arrcate(1,intC)&"_금액", clsCMeachulLog.Fcatebuy_P(arrcate(3,intc))
                            .Add arrcate(1,intC)&"_율", "" 
                            next
                        end if  
                    end if
                    End With     
                    hidx_buy2 = J
                    j = j+ 1  
                    m = 0 
                end if  

                if fnGetMoneyType(clsCMeachulLog.Fmwdiv(i))<>fnGetMoneyType(clsCMeachulLog.Fmwdiv(i-1))  then  '매입구분별 구매총액   
                   if clsCMeachulLog.Fsitename(i) = "10x10" then
                        if clsCMeachulLog.Fmwdiv(i) ="M" then
                            buyVar = clsCMeachulLog.Fbuy_10I  
                            itemVar = clsCMeachulLog.FItem_10I
                            odsVar = clsCMeachulLog.FOds_10I
                        elseif  clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U"  then
                            buyVar = clsCMeachulLog.Fbuy_10C
                            itemVar = clsCMeachulLog.FItem_10C
                            odsVar = clsCMeachulLog.FOds_10C
                        else
                            buyVar = clsCMeachulLog.Fbuy_10E
                            itemVar = clsCMeachulLog.FItem_10E
                            odsVar = clsCMeachulLog.FOds_10E
                        end if       
                    else
                        if clsCMeachulLog.Fmwdiv(i) ="M" then
                            buyVar = clsCMeachulLog.Fbuy_PI 
                            itemVar = clsCMeachulLog.FItem_PI
                            odsVar = clsCMeachulLog.FOds_PI
                        elseif  clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U"  then
                            buyVar = clsCMeachulLog.Fbuy_PC
                            itemVar = clsCMeachulLog.FItem_PC
                            odsVar = clsCMeachulLog.FOds_PC
                        else
                            buyVar = clsCMeachulLog.Fbuy_PE
                            itemVar = clsCMeachulLog.FItem_PE
                            odsVar = clsCMeachulLog.FOds_PE
                        end if    
                    end if   
  
                    .Add j, tmpJSON.Collection()
                    With .item(j) 
                    .Add "ID", j 
                    .Add "Head_ID", hidx_buy2
                    .Add "구분",  fnGetMoneyType(clsCMeachulLog.Fmwdiv(i)) 
                    .Add "전체_주문건수", odsVar
                    .Add "전체_상품수량", itemVar
                    .Add "전체_금액", buyVar
                    .Add "전체_율", ""
                        
                    if sdispcate ="Y" then
                        if isArray(arrcate) then 
                            for intC= 0 To ubound(arrcate,2)
                                 if clsCMeachulLog.Fsitename(i) = "10x10" then
                                    if clsCMeachulLog.Fmwdiv(i) ="M" then
                                        buyVarCate = clsCMeachulLog.Fcatebuy_10I(arrcate(3,intc))  
                                        itemVarcate = clsCMeachulLog.FcateItem_10I(arrcate(3,intc))  
                                    elseif  clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U"  then
                                        buyVarCate = clsCMeachulLog.Fcatebuy_10C(arrcate(3,intc))
                                        itemVarcate = clsCMeachulLog.FcateItem_10C(arrcate(3,intc)) 
                                    else
                                        buyVarCate = clsCMeachulLog.Fcatebuy_10E(arrcate(3,intc))
                                        itemVarcate = clsCMeachulLog.FcateItem_10E(arrcate(3,intc)) 
                                    end if       
                                else
                                    if clsCMeachulLog.Fmwdiv(i) ="M" then
                                        buyVarCate = clsCMeachulLog.Fcatebuy_PI(arrcate(3,intc)) 
                                        itemVarcate = clsCMeachulLog.FcateItem_PI(arrcate(3,intc)) 
                                    elseif  clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U"  then
                                        buyVarCate = clsCMeachulLog.Fcatebuy_PC(arrcate(3,intc))
                                        itemVarcate = clsCMeachulLog.FcateItem_PC(arrcate(3,intc)) 
                                    else
                                        buyVarCate = clsCMeachulLog.Fcatebuy_PE(arrcate(3,intc))
                                        itemVarcate = clsCMeachulLog.FcateItem_PE(arrcate(3,intc)) 
                                    end if    
                                end if  
                                .Add arrcate(1,intC)&"_상품수량",  itemVarCate    
                                .Add arrcate(1,intC)&"_금액",  buyVarCate
                                .Add arrcate(1,intC)&"_율", ""
                            next
                        end if  
                    end if
                
                    End With   
                    hidx_buy3 = J
                    j = j+ 1   
                end if     
            end if
 
            
            '매입구분별 구매총액    
            .Add j, tmpJSON.Collection()
            With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_buy3
            .Add "구분",getmwdiv_beasongdivname(clsCMeachulLog.Fmwdiv(i)) 
            .Add "전체_주문건수", clsCMeachulLog.FOds(i)  
            .Add "전체_상품수량", clsCMeachulLog.FItem(i)  
            .Add "전체_금액", clsCMeachulLog.Fbuy(i)  
            .Add "전체_율", ""
                
             if sdispcate ="Y" then
                if isArray(arrcate) then 
                    for intC= 0 To ubound(arrcate,2)
                        if clsCMeachulLog.Fbuy_sum > 0 then  
                            mper = (clsCMeachulLog.Fcatebuy(i,arrcate(3,intc))/clsCMeachulLog.Fbuy_sum)*100
                        end if
                        .Add arrcate(1,intC)&"_상품수량",  clsCMeachulLog.FcateItem(i,arrcate(3,intc))
                        .Add arrcate(1,intC)&"_금액",  clsCMeachulLog.Fcatebuy(i,arrcate(3,intc))
                        .Add arrcate(1,intC)&"_율",""
                    next
                end if  
            end if
           
            End With    
            j = j+ 1  
	     Next
       
       '------------------구매총액 수익----------------------------------------------------
            '구매총액수익
              buyPer= 0 
			if clsCMeachulLog.Fbuy_sum > 0 then 
			buyPer = (clsCMeachulLog.FbuyPF_sum/clsCMeachulLog.Fbuy_sum)*100
			end if
            .Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", -1 
			.Add "구분", "구매총액수익" 
            .Add "전체_주문건수", clsCMeachulLog.FOds_sum
            .Add "전체_상품수량", clsCMeachulLog.FItem_sum
			.Add "전체_금액", clsCMeachulLog.FbuyPF_sum
			.Add "전체_율",  buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2)  
                 if clsCMeachulLog.Fcatebuy_Sum(arrcate(3,intc)) > 0 then 
                        mper = (clsCMeachulLog.FcatebuyPF_Sum(arrcate(3,intc))/clsCMeachulLog.Fcatebuy_Sum(arrcate(3,intc)))*100
                 end if 
                .Add arrcate(1,intC)&"_상품수량", clsCMeachulLog.FcateItem_Sum(arrcate(3,intc))
			    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FcatebuyPF_Sum(arrcate(3,intc))
                .Add arrcate(1,intC)&"_율",  mper/100 
                next
             end if   
			end if
			End With     
            hidx_buy1 = J
			j = j+ 1  

            
            '10x10 구매총액수익
            i=0
            buyPer= 0 
			if clsCMeachulLog.Fbuy_10 > 0 then 
			buyPer = (clsCMeachulLog.FbuyPF_10/clsCMeachulLog.Fbuy_10)*100
			end if
			.Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_buy1
			.Add "구분", "10x10" 
             .Add "전체_주문건수", clsCMeachulLog.FOds_10
            .Add "전체_상품수량", clsCMeachulLog.FItem_10
			.Add "전체_금액", clsCMeachulLog.FbuyPF_10
			.Add "전체_율", buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2) 
                    if clsCMeachulLog.Fcatebuy_10(arrcate(3,intc)) > 0 then
                        mper = (clsCMeachulLog.FcatebuyPF_10(arrcate(3,intc))/clsCMeachulLog.Fcatebuy_10(arrcate(3,intc)))*100
                    end if
                .Add arrcate(1,intC)&"_상품수량", clsCMeachulLog.FcateItem_10(arrcate(3,intc))    
			    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FcatebuyPF_10(arrcate(3,intc))
                .Add arrcate(1,intC)&"_율",  mper/100 
                next
             end if   
			end if
			End With    
            hidx_buy2 = J
			j = j+ 1
           m = 0
            
          '10x10 상품 구매총액
            i=0 
			.Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_buy2
			.Add "구분", "상품" 
            .Add "전체_주문건수", clsCMeachulLog.FOds_10I
            .Add "전체_상품수량", clsCMeachulLog.FItem_10I
			.Add "전체_금액", clsCMeachulLog.FbuyPF_10I
			.Add "전체_율",""
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2) 
                .Add arrcate(1,intC)&"_상품수량", clsCMeachulLog.FcateItem_10I(arrcate(3,intc))
			    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FcatebuyPF_10I(arrcate(3,intc))
                .Add arrcate(1,intC)&"_율", ""
                next
             end if   
			end if
			End With    
            hidx_buy3 = J
			j = j+ 1
           m = 0
          
         For i = 0 to iTotCnt 
            if i > 0 then
             '제휴 구매총액
                if clsCMeachulLog.Fsitename(i) <> clsCMeachulLog.Fsitename(i-1)  then
                buyPer= 0 
                    if clsCMeachulLog.Fbuy_P > 0 then 
                    buyPer = (clsCMeachulLog.FbuyPF_P/clsCMeachulLog.Fbuy_P)*100
                    end if
                    .Add j, tmpJSON.Collection()
                    With .item(j) 
                    .Add "ID", j 
                    .Add "Head_ID", hidx_buy1
                    .Add "구분",  "제휴" 
                    .Add "전체_주문건수", clsCMeachulLog.FOds_P
                    .Add "전체_상품수량", clsCMeachulLog.FItem_P
                    .Add "전체_금액", clsCMeachulLog.FbuyPF_P
                    .Add "전체_율", buyPer/100
                    if sdispcate ="Y" then
                        if isArray(arrcate) then 
                            for intC= 0 To ubound(arrcate,2) 
                               if clsCMeachulLog.Fcatebuy_P(arrcate(3,intc)) > 0 then
                                    mper = (clsCMeachulLog.FcatebuyPF_P(arrcate(3,intc))/clsCMeachulLog.Fcatebuy_P(arrcate(3,intc)))*100
                            end if
                            .Add arrcate(1,intC)&"_상품수량", clsCMeachulLog.FcateItem_P(arrcate(3,intc))
                            .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FcatebuyPF_P(arrcate(3,intc))
                            .Add arrcate(1,intC)&"_율",mper/100 
                            next
                        end if  
                    end if
                    End With     
                    hidx_buy2 = J
                    j = j+ 1  
                    m = 0 
                end if  

                if fnGetMoneyType(clsCMeachulLog.Fmwdiv(i))<>fnGetMoneyType(clsCMeachulLog.Fmwdiv(i-1))  then  '매입구분별 구매총액   
                   if clsCMeachulLog.Fsitename(i) = "10x10" then
                        if clsCMeachulLog.Fmwdiv(i) ="M" then
                            buyVar = clsCMeachulLog.FbuyPF_10I  
                            itemVar = clsCMeachulLog.FItem_10I
                            odsVar = clsCMeachulLog.FOds_10I
                        elseif  clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U"  then
                            buyVar = clsCMeachulLog.FbuyPF_10C
                            itemVar = clsCMeachulLog.FItem_10C
                            odsVar = clsCMeachulLog.FOds_10C
                        else
                            buyVar = clsCMeachulLog.FbuyPF_10E
                             itemVar = clsCMeachulLog.FItem_10E
                            odsVar = clsCMeachulLog.FOds_10E
                        end if       
                    else
                        if clsCMeachulLog.Fmwdiv(i) ="M" then
                            buyVar = clsCMeachulLog.FbuyPF_PI 
                             itemVar = clsCMeachulLog.FItem_PI
                            odsVar = clsCMeachulLog.FOds_PI
                        elseif  clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U"  then
                            buyVar = clsCMeachulLog.FbuyPF_PC
                            itemVar = clsCMeachulLog.FItem_PC
                            odsVar = clsCMeachulLog.FOds_PC
                        else
                            buyVar = clsCMeachulLog.FbuyPF_PE
                            itemVar = clsCMeachulLog.FItem_PE
                            odsVar = clsCMeachulLog.FOds_PE
                        end if    
                    end if   
 

                    .Add j, tmpJSON.Collection()
                    With .item(j) 
                    .Add "ID", j 
                    .Add "Head_ID", hidx_buy2
                    .Add "구분",  fnGetMoneyType(clsCMeachulLog.Fmwdiv(i)) 
                    .Add "전체_주문건수", odsVar
                    .Add "전체_상품수량", itemVar
                    .Add "전체_금액", buyVar
                    .Add "전체_율", ""
                        
                    if sdispcate ="Y" then
                        if isArray(arrcate) then 
                            for intC= 0 To ubound(arrcate,2)
                                 if clsCMeachulLog.Fsitename(i) = "10x10" then
                                    if clsCMeachulLog.Fmwdiv(i) ="M" then
                                        buyVarCate = clsCMeachulLog.FcatebuyPF_10I(arrcate(3,intc))  
                                        itemVarcate = clsCMeachulLog.FcateItem_10I(arrcate(3,intc))  
                                    elseif  clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U"  then
                                        buyVarCate = clsCMeachulLog.FcatebuyPF_10C(arrcate(3,intc))
                                        itemVarcate = clsCMeachulLog.FcateItem_10C(arrcate(3,intc)) 
                                    else
                                        buyVarCate = clsCMeachulLog.FcatebuyPF_10E(arrcate(3,intc))
                                        itemVarcate = clsCMeachulLog.FcateItem_10E(arrcate(3,intc)) 
                                    end if       
                                else
                                    if clsCMeachulLog.Fmwdiv(i) ="M" then
                                        buyVarCate = clsCMeachulLog.FcatebuyPF_PI(arrcate(3,intc)) 
                                         itemVarcate = clsCMeachulLog.FcateItem_PI(arrcate(3,intc)) 
                                    elseif  clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U"  then
                                        buyVarCate = clsCMeachulLog.FcatebuyPF_PC(arrcate(3,intc))
                                         itemVarcate = clsCMeachulLog.FcateItem_PC(arrcate(3,intc)) 
                                    else
                                        buyVarCate = clsCMeachulLog.FcatebuyPF_PE(arrcate(3,intc))
                                        itemVarcate = clsCMeachulLog.FcateItem_PE(arrcate(3,intc)) 
                                    end if    
                                end if   
                                .Add arrcate(1,intC)&"_상품수량",  itemVarCate    
                                .Add arrcate(1,intC)&"_금액",  buyVarCate
                                .Add arrcate(1,intC)&"_율", ""
                            next
                        end if  
                    end if
                
                    End With   
                    hidx_buy3 = J
                    j = j+ 1   
                end if     
            end if
 
            
            '매입구분별 구매총액    
            .Add j, tmpJSON.Collection()
            With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_buy3
            .Add "구분",getmwdiv_beasongdivname(clsCMeachulLog.Fmwdiv(i)) 
            .Add "전체_주문건수", clsCMeachulLog.FOds(i)  
            .Add "전체_상품수량", clsCMeachulLog.FItem(i)  
            .Add "전체_금액", clsCMeachulLog.FbuyPF(i)  
            .Add "전체_율", ""
                
             if sdispcate ="Y" then
                if isArray(arrcate) then 
                    for intC= 0 To ubound(arrcate,2) 
                        .Add arrcate(1,intC)&"_상품수량",  clsCMeachulLog.FcateItem(i,arrcate(3,intc))
                        .Add arrcate(1,intC)&"_금액",  clsCMeachulLog.FcatebuyPF(i,arrcate(3,intc))
                        .Add arrcate(1,intC)&"_율", ""
                    next
                end if  
            end if
           
            End With    
            j = j+ 1  
	     Next
       

       '------------------보너스쿠폰 ----------------------------------------------------
       '보너스쿠폰 
        buyPer= 0 
			if clsCMeachulLog.Fbuy_sum > 0 then 
			buyPer = (clsCMeachulLog.FBC_sum/clsCMeachulLog.Fbuy_sum)*100
			end if

            .Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", -1 
			.Add "구분", "보너스쿠폰"  
             .Add "전체_주문건수", clsCMeachulLog.FOds_sum
            .Add "전체_상품수량", clsCMeachulLog.FItem_sum
			.Add "전체_금액", clsCMeachulLog.FBC_sum
			.Add "전체_율",  buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2)  
                  if clsCMeachulLog.Fcatebuy_Sum(arrcate(3,intc)) > 0 then 
                        mper = (clsCMeachulLog.FcateBC_Sum(arrcate(3,intc))/clsCMeachulLog.Fcatebuy_Sum(arrcate(3,intc)))*100
                    end if 
                .Add arrcate(1,intC)&"_상품수량", clsCMeachulLog.FcateItem_Sum(arrcate(3,intc))    
			    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FcateBC_Sum(arrcate(3,intc))
                .Add arrcate(1,intC)&"_율", mper/100 
                next
             end if   
			end if
			End With     
            hidx_buy1 = J
			j = j+ 1  

            
            '10x10  
            i=0
            buyPer= 0 
			if clsCMeachulLog.Fbuy_10 > 0 then 
			buyPer = (clsCMeachulLog.FBC_10/clsCMeachulLog.Fbuy_10)*100
			end if
			.Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_buy1
			.Add "구분", "10x10" 
             .Add "전체_주문건수", clsCMeachulLog.FOds_10
            .Add "전체_상품수량", clsCMeachulLog.FItem_10
			.Add "전체_금액", clsCMeachulLog.FBC_10
			.Add "전체_율", buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2) 
                 if clsCMeachulLog.Fcatebuy_10(arrcate(3,intc)) > 0 then
                        mper = (clsCMeachulLog.FcateBC_10(arrcate(3,intc))/clsCMeachulLog.Fcatebuy_10(arrcate(3,intc)))*100
                    end if
                 .Add arrcate(1,intC)&"_상품수량", clsCMeachulLog.FcateItem_10(arrcate(3,intc))    
			    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FcateBC_10(arrcate(3,intc))
                .Add arrcate(1,intC)&"_율",  mper/100 
                next
             end if   
			end if
			End With    
            hidx_buy2 = J
			j = j+ 1
           m = 0
            
          '10x10 상품  
            i=0
            
			.Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_buy2
			.Add "구분", "상품" 
             .Add "전체_주문건수", clsCMeachulLog.FOds_10I
            .Add "전체_상품수량", clsCMeachulLog.FItem_10I
			.Add "전체_금액", clsCMeachulLog.FBC_10I
			.Add "전체_율", ""
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2) 
                 .Add arrcate(1,intC)&"_상품수량", clsCMeachulLog.FcateItem_10I(arrcate(3,intc))
			    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FcateBC_10I(arrcate(3,intc))
                .Add arrcate(1,intC)&"_율",""
                next
             end if   
			end if
			End With    
            hidx_buy3 = J
			j = j+ 1
           m = 0
          
         For i = 0 to iTotCnt 
            if i > 0 then
             '제휴  
                if clsCMeachulLog.Fsitename(i) <> clsCMeachulLog.Fsitename(i-1)  then
                buyPer= 0 
                    if clsCMeachulLog.Fbuy_P > 0 then 
                    buyPer = (clsCMeachulLog.FBC_P/clsCMeachulLog.Fbuy_P)*100
                    end if
                    .Add j, tmpJSON.Collection()
                    With .item(j) 
                    .Add "ID", j 
                    .Add "Head_ID", hidx_buy1
                    .Add "구분",  "제휴" 
                     .Add "전체_주문건수", clsCMeachulLog.FOds_P
                    .Add "전체_상품수량", clsCMeachulLog.FItem_P
                    .Add "전체_금액", clsCMeachulLog.FBC_P
                    .Add "전체_율", buyPer/100
                    if sdispcate ="Y" then
                        if isArray(arrcate) then 
                            for intC= 0 To ubound(arrcate,2) 
                             if clsCMeachulLog.Fcatebuy_P(arrcate(3,intc)) > 0 then
                                    mper = (clsCMeachulLog.FcateBC_P(arrcate(3,intc))/clsCMeachulLog.Fcatebuy_P(arrcate(3,intc)))*100
                            end if
                             .Add arrcate(1,intC)&"_상품수량", clsCMeachulLog.FcateItem_P(arrcate(3,intc))
                            .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FcateBC_P(arrcate(3,intc))
                            .Add arrcate(1,intC)&"_율", mper/100 
                            next
                        end if  
                    end if
                    End With     
                    hidx_buy2 = J
                    j = j+ 1  
                    m = 0 
                end if  

                if fnGetMoneyType(clsCMeachulLog.Fmwdiv(i))<>fnGetMoneyType(clsCMeachulLog.Fmwdiv(i-1))  then  '매입구분별 구매총액   
                   if clsCMeachulLog.Fsitename(i) = "10x10" then
                        if clsCMeachulLog.Fmwdiv(i) ="M" then
                            buyVar = clsCMeachulLog.FBC_10I  
                            itemVar = clsCMeachulLog.FItem_10I
                            odsVar = clsCMeachulLog.FOds_10I
                        elseif  clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U"  then
                            buyVar = clsCMeachulLog.FBC_10C
                             itemVar = clsCMeachulLog.FItem_10C
                            odsVar = clsCMeachulLog.FOds_10C
                        else
                            buyVar = clsCMeachulLog.FBC_10E
                            itemVar = clsCMeachulLog.FItem_10E
                            odsVar = clsCMeachulLog.FOds_10E
                        end if       
                    else
                        if clsCMeachulLog.Fmwdiv(i) ="M" then
                            buyVar = clsCMeachulLog.FBC_PI 
                            itemVar = clsCMeachulLog.FItem_PI
                            odsVar = clsCMeachulLog.FOds_PI
                        elseif  clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U"  then
                            buyVar = clsCMeachulLog.FBC_PC
                             itemVar = clsCMeachulLog.FItem_PC
                            odsVar = clsCMeachulLog.FOds_PC
                        else
                            buyVar = clsCMeachulLog.FBC_PE
                             itemVar = clsCMeachulLog.FItem_PE
                            odsVar = clsCMeachulLog.FOds_PE
                        end if    
                    end if   
  
                    .Add j, tmpJSON.Collection()
                    With .item(j) 
                    .Add "ID", j 
                    .Add "Head_ID", hidx_buy2
                    .Add "구분",  fnGetMoneyType(clsCMeachulLog.Fmwdiv(i)) 
                    .Add "전체_주문건수", odsVar
                    .Add "전체_상품수량", itemVar
                    .Add "전체_금액", buyVar
                    .Add "전체_율", ""
                        
                    if sdispcate ="Y" then
                        if isArray(arrcate) then 
                            for intC= 0 To ubound(arrcate,2)
                                 if clsCMeachulLog.Fsitename(i) = "10x10" then
                                    if clsCMeachulLog.Fmwdiv(i) ="M" then
                                        buyVarCate = clsCMeachulLog.FcateBC_10I(arrcate(3,intc)) 
                                        itemVarcate = clsCMeachulLog.FcateItem_10I(arrcate(3,intc))   
                                    elseif  clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U"  then
                                        buyVarCate = clsCMeachulLog.FcateBC_10C(arrcate(3,intc))
                                         itemVarcate = clsCMeachulLog.FcateItem_10C(arrcate(3,intc))
                                    else
                                        buyVarCate = clsCMeachulLog.FcateBC_10E(arrcate(3,intc))
                                        itemVarcate = clsCMeachulLog.FcateItem_10E(arrcate(3,intc)) 
                                    end if       
                                else
                                    if clsCMeachulLog.Fmwdiv(i) ="M" then
                                        buyVarCate = clsCMeachulLog.FcateBC_PI(arrcate(3,intc)) 
                                         itemVarcate = clsCMeachulLog.FcateItem_PI(arrcate(3,intc)) 
                                    elseif  clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U"  then
                                        buyVarCate = clsCMeachulLog.FcateBC_PC(arrcate(3,intc))
                                        itemVarcate = clsCMeachulLog.FcateItem_PC(arrcate(3,intc)) 
                                    else
                                        buyVarCate = clsCMeachulLog.FcateBC_PE(arrcate(3,intc))
                                         itemVarcate = clsCMeachulLog.FcateItem_PE(arrcate(3,intc)) 
                                    end if    
                                end if  
                                 .Add arrcate(1,intC)&"_상품수량",  itemVarCate    
                                .Add arrcate(1,intC)&"_금액",  buyVarCate
                                .Add arrcate(1,intC)&"_율", ""
                            next
                        end if  
                    end if
                
                    End With   
                    hidx_buy3 = J
                    j = j+ 1   
                end if     
            end if
 
            
            '매입구분별    
            .Add j, tmpJSON.Collection()
            With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_buy3
            .Add "구분",getmwdiv_beasongdivname(clsCMeachulLog.Fmwdiv(i)) 
            .Add "전체_주문건수", clsCMeachulLog.FOds(i)  
            .Add "전체_상품수량", clsCMeachulLog.FItem(i)  
            .Add "전체_금액", clsCMeachulLog.FBC(i)  
            .Add "전체_율",""
                
             if sdispcate ="Y" then
                if isArray(arrcate) then 
                    for intC= 0 To ubound(arrcate,2) 
                        .Add arrcate(1,intC)&"_상품수량",  clsCMeachulLog.FcateItem(i,arrcate(3,intc))
                        .Add arrcate(1,intC)&"_금액",  clsCMeachulLog.FcateBC(i,arrcate(3,intc))
                        .Add arrcate(1,intC)&"_율", ""
                    next
                end if  
            end if
           
            End With    
            j = j+ 1  
	     Next

         '------------------취급액 ---------------------------------------------------- 
         
            .Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", -1 
			.Add "구분", "취급액"  
             .Add "전체_주문건수", clsCMeachulLog.FOds_sum
            .Add "전체_상품수량", clsCMeachulLog.FItem_sum
			.Add "전체_금액", clsCMeachulLog.FMP_sum
			.Add "전체_율",  ""
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2)  
                .Add arrcate(1,intC)&"_상품수량", clsCMeachulLog.FcateItem_Sum(arrcate(3,intc))
			    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FcateMP_Sum(arrcate(3,intc))
                .Add arrcate(1,intC)&"_율", ""
                next
             end if   
			end if
			End With     
            hidx_buy1 = J
			j = j+ 1  

            
            '10x10  
            i=0 
			.Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_buy1
			.Add "구분", "10x10" 
             .Add "전체_주문건수", clsCMeachulLog.FOds_10
            .Add "전체_상품수량", clsCMeachulLog.FItem_10
			.Add "전체_금액", clsCMeachulLog.FMP_10
			.Add "전체_율", ""
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2) 
                 .Add arrcate(1,intC)&"_상품수량", clsCMeachulLog.FcateItem_10(arrcate(3,intc))
			    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FcateMP_10(arrcate(3,intc))
                .Add arrcate(1,intC)&"_율", ""
                next
             end if   
			end if
			End With    
            hidx_buy2 = J
			j = j+ 1
           m = 0
            
          '10x10 상품  
            i=0 
			.Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_buy2
			.Add "구분", "상품" 
             .Add "전체_주문건수", clsCMeachulLog.FOds_10I
            .Add "전체_상품수량", clsCMeachulLog.FItem_10I
			.Add "전체_금액", clsCMeachulLog.FMP_10I
			.Add "전체_율", ""
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2) 
                 .Add arrcate(1,intC)&"_상품수량", clsCMeachulLog.FcateItem_10I(arrcate(3,intc))
			    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FcateMP_10I(arrcate(3,intc))
                .Add arrcate(1,intC)&"_율", ""
                next
             end if   
			end if
			End With    
            hidx_buy3 = J
			j = j+ 1
           m = 0
          
         For i = 0 to iTotCnt 
            if i > 0 then
             '제휴  
                if clsCMeachulLog.Fsitename(i) <> clsCMeachulLog.Fsitename(i-1)  then
               
                    .Add j, tmpJSON.Collection()
                    With .item(j) 
                    .Add "ID", j 
                    .Add "Head_ID", hidx_buy1
                    .Add "구분",  "제휴" 
                     .Add "전체_주문건수", clsCMeachulLog.FOds_P
                    .Add "전체_상품수량", clsCMeachulLog.FItem_P
                    .Add "전체_금액", clsCMeachulLog.FMP_P
                    .Add "전체_율", ""
                    if sdispcate ="Y" then
                        if isArray(arrcate) then 
                            for intC= 0 To ubound(arrcate,2)
                             .Add arrcate(1,intC)&"_상품수량", clsCMeachulLog.FcateItem_P(arrcate(3,intc)) 
                            .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FcateMP_P(arrcate(3,intc))
                            .Add arrcate(1,intC)&"_율",""
                            next
                        end if  
                    end if
                    End With     
                    hidx_buy2 = J
                    j = j+ 1  
                    m = 0 
                end if  

                if fnGetMoneyType(clsCMeachulLog.Fmwdiv(i))<>fnGetMoneyType(clsCMeachulLog.Fmwdiv(i-1))  then  '매입구분별 구매총액   
                   if clsCMeachulLog.Fsitename(i) = "10x10" then
                        if clsCMeachulLog.Fmwdiv(i) ="M" then
                            buyVar = clsCMeachulLog.FMP_10I  
                            itemVar = clsCMeachulLog.FItem_10I
                            odsVar = clsCMeachulLog.FOds_10I
                        elseif  clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U"  then
                            buyVar = clsCMeachulLog.FMP_10C
                            itemVar = clsCMeachulLog.FItem_10C
                            odsVar = clsCMeachulLog.FOds_10C
                        else
                            buyVar = clsCMeachulLog.FMP_10E
                             itemVar = clsCMeachulLog.FItem_10E
                            odsVar = clsCMeachulLog.FOds_10E
                        end if       
                    else
                        if clsCMeachulLog.Fmwdiv(i) ="M" then
                            buyVar = clsCMeachulLog.FMP_PI 
                            itemVar = clsCMeachulLog.FItem_PI
                            odsVar = clsCMeachulLog.FOds_PI
                        elseif  clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U"  then
                            buyVar = clsCMeachulLog.FMP_PC
                            itemVar = clsCMeachulLog.FItem_PC
                            odsVar = clsCMeachulLog.FOds_PC
                        else
                            buyVar = clsCMeachulLog.FMP_PE
                            itemVar = clsCMeachulLog.FItem_PE
                            odsVar = clsCMeachulLog.FOds_PE
                        end if    
                    end if   
 

                    .Add j, tmpJSON.Collection()
                    With .item(j) 
                    .Add "ID", j 
                    .Add "Head_ID", hidx_buy2
                    .Add "구분",  fnGetMoneyType(clsCMeachulLog.Fmwdiv(i)) 
                    .Add "전체_금액", buyVar
                    .Add "전체_율", ""
                        
                    if sdispcate ="Y" then
                        if isArray(arrcate) then 
                            for intC= 0 To ubound(arrcate,2)
                                 if clsCMeachulLog.Fsitename(i) = "10x10" then
                                    if clsCMeachulLog.Fmwdiv(i) ="M" then
                                        buyVarCate = clsCMeachulLog.FcateMP_10I(arrcate(3,intc))  
                                        itemVarcate = clsCMeachulLog.FcateItem_10I(arrcate(3,intc))  
                                    elseif  clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U"  then
                                        buyVarCate = clsCMeachulLog.FcateMP_10C(arrcate(3,intc))
                                         itemVarcate = clsCMeachulLog.FcateItem_10C(arrcate(3,intc))
                                    else
                                        buyVarCate = clsCMeachulLog.FcateMP_10E(arrcate(3,intc))
                                        itemVarcate = clsCMeachulLog.FcateItem_10E(arrcate(3,intc)) 
                                    end if       
                                else
                                    if clsCMeachulLog.Fmwdiv(i) ="M" then
                                        buyVarCate = clsCMeachulLog.FcateMP_PI(arrcate(3,intc)) 
                                         itemVarcate = clsCMeachulLog.FcateItem_PI(arrcate(3,intc)) 
                                    elseif  clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U"  then
                                        buyVarCate = clsCMeachulLog.FcateMP_PC(arrcate(3,intc))
                                         itemVarcate = clsCMeachulLog.FcateItem_PC(arrcate(3,intc)) 
                                    else
                                        buyVarCate = clsCMeachulLog.FcateMP_PE(arrcate(3,intc))
                                        itemVarcate = clsCMeachulLog.FcateItem_PE(arrcate(3,intc)) 
                                    end if    
                                end if  
                                 .Add arrcate(1,intC)&"_상품수량",  itemVarCate 
                                .Add arrcate(1,intC)&"_금액",  buyVarCate
                                .Add arrcate(1,intC)&"_율", ""
                            next
                        end if  
                    end if
                
                    End With   
                    hidx_buy3 = J
                    j = j+ 1   
                end if     
            end if
 
            
            '매입구분별    
            .Add j, tmpJSON.Collection()
            With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_buy3
            .Add "구분",getmwdiv_beasongdivname(clsCMeachulLog.Fmwdiv(i)) 
            .Add "전체_주문건수", clsCMeachulLog.FOds(i)  
            .Add "전체_상품수량", clsCMeachulLog.FItem(i) 
            .Add "전체_금액", clsCMeachulLog.FMP(i)  
            .Add "전체_율", ""
                
             if sdispcate ="Y" then
                if isArray(arrcate) then 
                    for intC= 0 To ubound(arrcate,2) 
                        .Add arrcate(1,intC)&"_상품수량",  clsCMeachulLog.FcateItem(i,arrcate(3,intc))
                        .Add arrcate(1,intC)&"_금액",  clsCMeachulLog.FcateMP(i,arrcate(3,intc))
                        .Add arrcate(1,intC)&"_율", ""
                    next
                end if  
            end if
           
            End With    
            j = j+ 1  
	     Next

         '------------------취급액 수익 ---------------------------------------------------- 
            buyPer= 0 
			if clsCMeachulLog.FMP_sum > 0 then 
			buyPer = (clsCMeachulLog.FMPPF_sum/clsCMeachulLog.FMP_sum)*100
			end if
            .Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", -1 
			.Add "구분", "취급액 수익"  
            .Add "전체_주문건수", clsCMeachulLog.FOds_sum
            .Add "전체_상품수량", clsCMeachulLog.FItem_sum
			.Add "전체_금액", clsCMeachulLog.FMPPF_sum
			.Add "전체_율", buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2) 
                  if clsCMeachulLog.FcateMP_Sum(arrcate(3,intc)) > 0 then 
                        mper = (clsCMeachulLog.FcateMPPF_Sum(arrcate(3,intc))/clsCMeachulLog.FcateMP_Sum(arrcate(3,intc)))*100
                    end if 
                .Add arrcate(1,intC)&"_상품수량", clsCMeachulLog.FcateItem_Sum(arrcate(3,intc))    
			    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FcateMPPF_Sum(arrcate(3,intc))
                .Add arrcate(1,intC)&"_율", mper/100 
                next
             end if   
			end if
			End With     
            hidx_buy1 = J
			j = j+ 1  

            
            '10x10  
            i=0
            buyPer= 0 
			if clsCMeachulLog.FMP_10 > 0 then 
			buyPer = (clsCMeachulLog.FMPPF_10/clsCMeachulLog.FMP_10)*100
			end if
			.Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_buy1
			.Add "구분", "10x10" 
            .Add "전체_주문건수", clsCMeachulLog.FOds_10
            .Add "전체_상품수량", clsCMeachulLog.FItem_10
			.Add "전체_금액", clsCMeachulLog.FMPPF_10
			.Add "전체_율", buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2) 
                  if clsCMeachulLog.FcateMP_10(arrcate(3,intc)) > 0 then
                        mper = (clsCMeachulLog.FcateMPPF_10(arrcate(3,intc))/clsCMeachulLog.FcateMP_10(arrcate(3,intc)))*100
                   end if
                 .Add arrcate(1,intC)&"_상품수량", clsCMeachulLog.FcateItem_10(arrcate(3,intc))   
			    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FcateMPPF_10(arrcate(3,intc))
                .Add arrcate(1,intC)&"_율", mper/100 
                next
             end if   
			end if
			End With    
            hidx_buy2 = J
			j = j+ 1
           m = 0
            
          '10x10 상품  
            i=0 
			.Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_buy2
			.Add "구분", "상품" 
            .Add "전체_주문건수", clsCMeachulLog.FOds_10I
            .Add "전체_상품수량", clsCMeachulLog.FItem_10I
			.Add "전체_금액", clsCMeachulLog.FMPPF_10I
			.Add "전체_율", ""
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2) 
                 .Add arrcate(1,intC)&"_상품수량", clsCMeachulLog.FcateItem_10I(arrcate(3,intc))
			    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FcateMPPF_10I(arrcate(3,intc))
                .Add arrcate(1,intC)&"_율", ""
                next
             end if   
			end if
			End With    
            hidx_buy3 = J
			j = j+ 1
           m = 0
          
         For i = 0 to iTotCnt 
            if i > 0 then
             '제휴  
                if clsCMeachulLog.Fsitename(i) <> clsCMeachulLog.Fsitename(i-1)  then
                buyPer= 0 
                    if clsCMeachulLog.FMP_P > 0 then 
                    buyPer = (clsCMeachulLog.FMPPF_P/clsCMeachulLog.FMP_P)*100
                    end if
                    .Add j, tmpJSON.Collection()
                    With .item(j) 
                    .Add "ID", j 
                    .Add "Head_ID", hidx_buy1
                    .Add "구분",  "제휴" 
                     .Add "전체_주문건수", clsCMeachulLog.FOds_P
                    .Add "전체_상품수량", clsCMeachulLog.FItem_P
                    .Add "전체_금액", clsCMeachulLog.FMPPF_P
                    .Add "전체_율", buyPer/100
                    if sdispcate ="Y" then
                        if isArray(arrcate) then 
                            for intC= 0 To ubound(arrcate,2) 
                             if  clsCMeachulLog.FcateMP_P(arrcate(3,intc)) > 0 then
                                    mper = (clsCMeachulLog.FcateMPPF_P(arrcate(3,intc))/ clsCMeachulLog.FcateMP_P(arrcate(3,intc)))*100
                            end if
                            .Add arrcate(1,intC)&"_상품수량", clsCMeachulLog.FcateItem_P(arrcate(3,intc))
                            .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FcateMPPF_P(arrcate(3,intc))
                            .Add arrcate(1,intC)&"_율", mper/100 
                            next
                        end if  
                    end if
                    End With     
                    hidx_buy2 = J
                    j = j+ 1  
                    m = 0 
                end if  

                if fnGetMoneyType(clsCMeachulLog.Fmwdiv(i))<>fnGetMoneyType(clsCMeachulLog.Fmwdiv(i-1))  then  '매입구분별 구매총액   
                   if clsCMeachulLog.Fsitename(i) = "10x10" then
                        if clsCMeachulLog.Fmwdiv(i) ="M" then
                            buyVar = clsCMeachulLog.FMPPF_10I  
                            itemVar = clsCMeachulLog.FItem_10I
                            odsVar = clsCMeachulLog.FOds_10I
                        elseif  clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U"  then
                            buyVar = clsCMeachulLog.FMPPF_10C
                             itemVar = clsCMeachulLog.FItem_10C
                            odsVar = clsCMeachulLog.FOds_10C
                        else
                            buyVar = clsCMeachulLog.FMPPF_10E
                             itemVar = clsCMeachulLog.FItem_10E
                            odsVar = clsCMeachulLog.FOds_10E
                        end if       
                    else
                        if clsCMeachulLog.Fmwdiv(i) ="M" then
                            buyVar = clsCMeachulLog.FMPPF_PI 
                            itemVar = clsCMeachulLog.FItem_PI
                            odsVar = clsCMeachulLog.FOds_PI
                        elseif  clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U"  then
                            buyVar = clsCMeachulLog.FMPPF_PC
                            itemVar = clsCMeachulLog.FItem_PC
                            odsVar = clsCMeachulLog.FOds_PC
                        else
                            buyVar = clsCMeachulLog.FMPPF_PE
                             itemVar = clsCMeachulLog.FItem_PE
                            odsVar = clsCMeachulLog.FOds_PE
                        end if    
                    end if   

                    

                    .Add j, tmpJSON.Collection()
                    With .item(j) 
                    .Add "ID", j 
                    .Add "Head_ID", hidx_buy2
                    .Add "구분",  fnGetMoneyType(clsCMeachulLog.Fmwdiv(i)) 
                     .Add "전체_주문건수", odsVar
                    .Add "전체_상품수량", itemVar
                    .Add "전체_금액", buyVar
                    .Add "전체_율",""
                        
                    if sdispcate ="Y" then
                        if isArray(arrcate) then 
                            for intC= 0 To ubound(arrcate,2)
                                 if clsCMeachulLog.Fsitename(i) = "10x10" then
                                    if clsCMeachulLog.Fmwdiv(i) ="M" then
                                        buyVarCate = clsCMeachulLog.FcateMPPF_10I(arrcate(3,intc))  
                                         itemVarcate = clsCMeachulLog.FcateItem_10I(arrcate(3,intc))  
                                    elseif  clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U"  then
                                        buyVarCate = clsCMeachulLog.FcateMPPF_10C(arrcate(3,intc))
                                         itemVarcate = clsCMeachulLog.FcateItem_10C(arrcate(3,intc)) 
                                    else
                                        buyVarCate = clsCMeachulLog.FcateMPPF_10E(arrcate(3,intc))
                                         itemVarcate = clsCMeachulLog.FcateItem_10E(arrcate(3,intc)) 
                                    end if       
                                else
                                    if clsCMeachulLog.Fmwdiv(i) ="M" then
                                        buyVarCate = clsCMeachulLog.FcateMPPF_PI(arrcate(3,intc)) 
                                         itemVarcate = clsCMeachulLog.FcateItem_PI(arrcate(3,intc)) 
                                    elseif  clsCMeachulLog.Fmwdiv(i) ="W" or clsCMeachulLog.Fmwdiv(i) ="U"  then
                                        buyVarCate = clsCMeachulLog.FcateMPPF_PC(arrcate(3,intc))
                                          itemVarcate = clsCMeachulLog.FcateItem_PC(arrcate(3,intc)) 
                                    else
                                        buyVarCate = clsCMeachulLog.FcateMPPF_PE(arrcate(3,intc))
                                         itemVarcate = clsCMeachulLog.FcateItem_PE(arrcate(3,intc)) 
                                    end if    
                                end if   
                                 .Add arrcate(1,intC)&"_상품수량",  itemVarCate    
                                .Add arrcate(1,intC)&"_금액",  buyVarCate 
                                .Add arrcate(1,intC)&"_율", ""
                            next
                        end if  
                    end if
                
                    End With   
                    hidx_buy3 = J
                    j = j+ 1   
                end if     
            end if
 
            
            '매입구분별   
           
            .Add j, tmpJSON.Collection()
            With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_buy3
            .Add "구분",getmwdiv_beasongdivname(clsCMeachulLog.Fmwdiv(i)) 
             .Add "전체_주문건수", clsCMeachulLog.FOds(i)  
            .Add "전체_상품수량", clsCMeachulLog.FItem(i)  
            .Add "전체_금액", clsCMeachulLog.FMPPF(i)  
            .Add "전체_율", ""
                
             if sdispcate ="Y" then
                if isArray(arrcate) then 
                    for intC= 0 To ubound(arrcate,2) 
                      .Add arrcate(1,intC)&"_상품수량",  clsCMeachulLog.FcateItem(i,arrcate(3,intc))
                        .Add arrcate(1,intC)&"_금액",  clsCMeachulLog.FcateMPPF(i,arrcate(3,intc))
                        .Add arrcate(1,intC)&"_율", ""
                    next
                end if  
            end if
           
            End With    
            j = j+ 1  
	     Next
    end if     

    '------ 변동비1 ----------	
    dim  totvc,totvc_10, totvc_p, totsc_10, totsc_P , totwh_10, totwh_p
    dim arrPF, intp
    dim hidx_vc1, hidx_vc1_10, hidx_vc1_P, hidx_sc_10 ,hidx_wh_10, hidx_sc_p ,hidx_wh_p ,cper
        totvc = 0 : totvc_10=0:totvc_p=0: totsc_10=0: totsc_P=0 : totwh_10=0:totwh_p=0

	   clsCMeachulLog.fnGetprofitlossdata   
       totsc_10 = clsCMeachulLog.FTotSC_10
       totsc_P = clsCMeachulLog.FTotSC_P
      
       totwh_10= clsCMeachulLog.FTotWH_10
       totwh_p= clsCMeachulLog.FTotWH_P

       totvc_10 = totsc_10 + totwh_10
       totvc_P = totsc_P + totwh_p

       totvc = totvc_10 + totvc_P

         buyPer= 0 
            if clsCMeachulLog.FMP_sum > 0 then 
                buyPer = (totvc/clsCMeachulLog.FMP_sum)*100
            end if
         .Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", -1 
			.Add "구분", "변동비1(수수료,물류비)"  
			.Add "전체_금액", totvc
			.Add "전체_율",  buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) and itotcnt>0 then 
                for intC= 0 To ubound(arrcate,2)  
                    if  clsCMeachulLog.FcateMP_Sum(arrcate(3,intc)) > 0 then  
                    cper = ( clsCMeachulLog.FTotVC1_Cate(arrcate(3,intC))/ clsCMeachulLog.FcateMP_Sum(arrcate(3,intc)))*100
                     end if 
			    .Add arrcate(1,intC)&"_금액",  clsCMeachulLog.FTotVC1_Cate(arrcate(3,intC))
                .Add arrcate(1,intC)&"_율",cper/100
                next
             end if   
			end if
			End With    
            hidx_vc1 = J
			j = j+ 1  

             buyPer= 0 
            if clsCMeachulLog.FMP_10  > 0 then 
                buyPer = (totvc_10/clsCMeachulLog.FMP_10)*100
            end if
         .Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_vc1
			.Add "구분", "10x10" 
			.Add "전체_금액", totvc_10
			.Add "전체_율",  buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2)   
                 if clsCMeachulLog.FcateMP_10(arrcate(3,intc)) > 0 then  
                  cper = ( clsCMeachulLog.FTotVC1_Cate_10(arrcate(3,intC))/clsCMeachulLog.FcateMP_10(arrcate(3,intc)))*100
               end if
			    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FTotVC1_Cate_10(arrcate(3,intC))
                .Add arrcate(1,intC)&"_율",cper/100
                next
             end if   
			end if
			End With    
            hidx_vc1_10 = J
			j = j+ 1  

         buyPer= 0 
            if clsCMeachulLog.FMP_P > 0 then 
                buyPer = (totvc_P/clsCMeachulLog.FMP_P)*100
            end if
         .Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_vc1
			.Add "구분", "제휴"  
			.Add "전체_금액", totvc_P
			.Add "전체_율",  buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2)  
                 if clsCMeachulLog.FcateMP_P(arrcate(3,intc)) > 0 then  
                  cper = ( clsCMeachulLog.FTotVC1_Cate_P(arrcate(3,intC))/clsCMeachulLog.FcateMP_P(arrcate(3,intc)))*100
                end if
			    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FTotVC1_Cate_P(arrcate(3,intC))
                .Add arrcate(1,intC)&"_율",cper/100
                next
             end if   
			end if
			End With    
            hidx_vc1_P = J
			j = j+ 1 


        buyPer= 0 
            if clsCMeachulLog.FMP_10 > 0 then 
                buyPer = (totsc_10/clsCMeachulLog.FMP_10)*100
            end if
        .Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_vc1_10
			.Add "구분", "판매수수료" 
			.Add "전체_금액", totsc_10 
			.Add "전체_율",  buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2)  
                    if clsCMeachulLog.FcateMP_10(arrcate(3,intc)) > 0 then  
                    cper = ( clsCMeachulLog.FTotSC_Cate_10(arrcate(3,intC))/clsCMeachulLog.FcateMP_10(arrcate(3,intc)))*100
                    end if
                    .Add arrcate(1,intC)&"_금액",  clsCMeachulLog.FTotSC_Cate_10(arrcate(3,intC))
                    .Add arrcate(1,intC)&"_율", cper/100
                next
             end if   
			end if
			End With    
            hidx_sc_10= J
			j = j+ 1 

         buyPer= 0 
            if clsCMeachulLog.FMP_10 > 0 then 
                buyPer = (totwh_10/clsCMeachulLog.FMP_10)*100
            end if
        .Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_vc1_10
			.Add "구분", "물류비"  
			.Add "전체_금액", totwh_10 
			.Add "전체_율",  buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2)  
                     if clsCMeachulLog.FcateMP_10(arrcate(3,intc)) > 0 then  
                    cper = ( clsCMeachulLog.FTotWH_Cate_10(arrcate(3,intC))/clsCMeachulLog.FcateMP_10(arrcate(3,intc)))*100
                    end if
                    .Add arrcate(1,intC)&"_금액",  clsCMeachulLog.FTotWH_Cate_10(arrcate(3,intC))
                    .Add arrcate(1,intC)&"_율",cper/100
                next
             end if   
			end if
			End With    
            hidx_wh_10= J
			j = j+ 1 


         buyPer= 0 
            if clsCMeachulLog.FMP_P > 0 then 
                buyPer = (totsc_P/clsCMeachulLog.FMP_P)*100
            end if
        .Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_vc1_P
			.Add "구분", "판매수수료"  
			.Add "전체_금액", totsc_P
			.Add "전체_율",  buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2)  
                    if clsCMeachulLog.FcateMP_P(arrcate(3,intc)) > 0 then  
                    cper = ( clsCMeachulLog.FTotSC_Cate_P(arrcate(3,intC))/clsCMeachulLog.FcateMP_P(arrcate(3,intc)))*100
                    end if
                    .Add arrcate(1,intC)&"_금액",  clsCMeachulLog.FTotSC_Cate_P(arrcate(3,intC))
                    .Add arrcate(1,intC)&"_율",cper/100
                next
             end if   
			end if
			End With    
            hidx_sc_P= J
			j = j+ 1 

         buyPer= 0 
            if clsCMeachulLog.FMP_P > 0 then 
                buyPer = (totwh_P/clsCMeachulLog.FMP_P)*100
            end if
        .Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_vc1_P
			.Add "구분", "물류비"   
			.Add "전체_금액", totwh_P 
			.Add "전체_율",  buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2)  
                     if clsCMeachulLog.FcateMP_P(arrcate(3,intc)) > 0 then  
                    cper = ( clsCMeachulLog.FTotWH_Cate_P(arrcate(3,intC))/clsCMeachulLog.FcateMP_P(arrcate(3,intc)))*100
                    end if
                    .Add arrcate(1,intC)&"_금액",  clsCMeachulLog.FTotWH_Cate_P(arrcate(3,intC))
                    .Add arrcate(1,intC)&"_율",cper/100
                next
             end if   
			end if
			End With    
            hidx_wh_p= J
			j = j+ 1 
 
 
dim hidx1  

         buyPer= 0 
			if (clsCMeachulLog.FMP_sum) > 0 then 
			buyPer = ((clsCMeachulLog.FMPPF_sum- totvc)/clsCMeachulLog.FMP_sum)*100
			end if
         .Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", -1 
			.Add "구분", "공헌이익 1"   
			.Add "전체_금액", clsCMeachulLog.FMPPF_sum- totvc
			.Add "전체_율",  buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) and iTotCnt > 0 then 
                for intC= 0 To ubound(arrcate,2)   
                    if clsCMeachulLog.FcateMP_Sum(arrcate(3,intc)) > 0 then  
                    cper = ((clsCMeachulLog.FcateMPPF_Sum(arrcate(3,intc))- clsCMeachulLog.FTotVC1_Cate(arrcate(3,intC)))/clsCMeachulLog.FcateMP_Sum(arrcate(3,intc)))*100
                    end if
                    .Add arrcate(1,intC)&"_금액",  (clsCMeachulLog.FcateMPPF_Sum(arrcate(3,intc))- clsCMeachulLog.FTotVC1_Cate(arrcate(3,intC))) 
                   .Add arrcate(1,intC)&"_율",cper/100
                next
             end if   
			end if
			End With    
            hidx1 = J
			j = j+ 1   


          buyPer= 0 
			if (clsCMeachulLog.FMP_10) > 0 then 
			buyPer = ((clsCMeachulLog.FMPPF_10-totvc_10)/clsCMeachulLog.FMP_10)*100
			end if
			.Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx1
			.Add "구분", "공헌이익 1(10x10)" 
			.Add "전체_금액", clsCMeachulLog.FMPPF_10-totvc_10
			.Add "전체_율", buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) and iTotCnt > 0 then 
                for intC= 0 To ubound(arrcate,2) 
                    if clsCMeachulLog.FcateMP_10(arrcate(3,intc)) > 0 then
                        cper = ((clsCMeachulLog.FcateMPPF_10(arrcate(3,intc))-clsCMeachulLog.FTotVC1_Cate_10(arrcate(3,intC)))/clsCMeachulLog.FcateMP_10(arrcate(3,intc)))*100
                    end if
			        .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FcateMPPF_10(arrcate(3,intc))-clsCMeachulLog.FTotVC1_Cate_10(arrcate(3,intC)) 
                    .Add arrcate(1,intC)&"_율", cper/100 
                next
             end if   
			end if
			End With   
			j = j+ 1


             buyPer= 0 
            if clsCMeachulLog.FMP_P > 0 then 
                buyPer = ((clsCMeachulLog.FMPPF_P-totvc_P)/clsCMeachulLog.FMP_P)*100
            end if
         .Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx1
			.Add "구분", "공헌이익 1(제휴)"  
			.Add "전체_금액", clsCMeachulLog.FMPPF_P-totvc_P
			.Add "전체_율",  buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) and itotCnt >0 then 
                for intC= 0 To ubound(arrcate,2)  
                    if clsCMeachulLog.FcateMP_P(arrcate(3,intc)) > 0 then  
                    cper = ((clsCMeachulLog.FcateMPPF_P(arrcate(3,intc))-clsCMeachulLog.FTotVC1_Cate_P(arrcate(3,intC)))/clsCMeachulLog.FcateMP_P(arrcate(3,intc)))*100
                    end if
                    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FcateMPPF_P(arrcate(3,intc))-clsCMeachulLog.FTotVC1_Cate_P(arrcate(3,intC))  
                    .Add arrcate(1,intC)&"_율",cper/100 
                next
             end if   
			end if
			End With     
			j = j+ 1 

dim hidx_vc2,hidx_vc2_10,hidx_vc2_P

             buyPer= 0 
			if (clsCMeachulLog.FMP_sum) > 0 then  
			buyPer = ((clsCMeachulLog.FTotMF_10 + clsCMeachulLog.FTotMF_P)/clsCMeachulLog.FMP_sum)*100
			end if
            .Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", -1 
			.Add "구분", "변동비2(광고판촉비)"   
			.Add "전체_금액", clsCMeachulLog.FTotMF_10 + clsCMeachulLog.FTotMF_P
			.Add "전체_율",  buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) and itotcnt > 0 then 
                for intC= 0 To ubound(arrcate,2)  
                    if clsCMeachulLog.FcateMP_Sum(arrcate(3,intc)) > 0 then  
                    cper = ( clsCMeachulLog.FTotVC2_Cate(arrcate(3,intC))/clsCMeachulLog.FcateMP_Sum(arrcate(3,intc)))*100
                    end if
                    .Add arrcate(1,intC)&"_금액",  clsCMeachulLog.FTotVC2_Cate(arrcate(3,intC))
                    .Add arrcate(1,intC)&"_율",cper/100 
                next
             end if   
			end if
			End With    
            hidx_vc2 = J
			j = j+ 1  

             buyPer= 0 
            if clsCMeachulLog.FMP_10 > 0 then 
                buyPer = (clsCMeachulLog.FTotMF_10/clsCMeachulLog.FMP_10)*100
            end if
         .Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_vc2
			.Add "구분", "10x10" 
			.Add "전체_금액", clsCMeachulLog.FTotMF_10
			.Add "전체_율",  buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) and itotcnt>0 then 
                for intC= 0 To ubound(arrcate,2) 
			      if clsCMeachulLog.FcateMP_10(arrcate(3,intc)) > 0 then  
                  cper = ( clsCMeachulLog.FTotVC2_Cate_10(arrcate(3,intC))/clsCMeachulLog.FcateMP_10(arrcate(3,intc)))*100
                 end if
			    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FTotVC2_Cate_10(arrcate(3,intC))
                .Add arrcate(1,intC)&"_율",cper/100
                next
             end if   
			end if
			End With    
            hidx_vc2_10 = J
			j = j+ 1  

              buyPer= 0 
            if  clsCMeachulLog.FMP_P > 0 then 
                buyPer = (clsCMeachulLog.FTotMF_P/ clsCMeachulLog.FMP_P)*100
            end if
         .Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx_vc2
			.Add "구분", "제휴"   
			.Add "전체_금액", clsCMeachulLog.FTotMF_P
			.Add "전체_율",  buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) and itotcnt>0 then 
                for intC= 0 To ubound(arrcate,2) 
			      if clsCMeachulLog.FcateMP_P(arrcate(3,intc)) > 0 then  
                  cper = ( clsCMeachulLog.FTotVC2_Cate_P(arrcate(3,intC))/clsCMeachulLog.FcateMP_P(arrcate(3,intc)))*100
                end if
			    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FTotVC2_Cate_P(arrcate(3,intC))
                .Add arrcate(1,intC)&"_율",cper/100
                next
             end if   
			end if
			End With    
            hidx_vc2_P = J
			j = j+ 1 

             dim hidxNM
        for intp = 0 to clsCMeachulLog.Fscrow-1  
          .Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
             if clsCMeachulLog.FAccCIdx(intp) = 5 and clsCMeachulLog.FsiteNM(intp) ="10x10" then 
            .Add "Head_ID",  hidx_sc_10
             elseif clsCMeachulLog.FAccCIdx(intp) = 5 and clsCMeachulLog.FsiteNM(intp) ="제휴" then 
            .Add "Head_ID",  hidx_sc_p
            elseif clsCMeachulLog.FAccCIdx(intp) = 6 and clsCMeachulLog.FsiteNM(intp) ="10x10" then 
            .Add "Head_ID",  hidx_wh_10
            elseif clsCMeachulLog.FAccCIdx(intp) = 6 and clsCMeachulLog.FsiteNM(intp) ="제휴" then 
             .Add "Head_ID",  hidx_wh_p
             elseif clsCMeachulLog.FAccCIdx(intp) = 9 and clsCMeachulLog.FsiteNM(intp) ="10x10" then 
             .Add "Head_ID",  hidx_vc2_10 
             elseif clsCMeachulLog.FAccCIdx(intp) = 9 and clsCMeachulLog.FsiteNM(intp) ="제휴" then 
             .Add "Head_ID",  hidx_vc2_p 
            end if 
			.Add "구분", clsCMeachulLog.FAccNM(intp)
			.Add "전체_금액", clsCMeachulLog.FPFPrice(intp)
			.Add "전체_율",  ""
			if sdispcate ="Y" then
             if isArray(arrcate) then 
                for intC= 0 To ubound(arrcate,2)  
                    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FCatePrice(intp, arrcate(3,intC))
                    .Add arrcate(1,intC)&"_율",""
                next
             end if   
			end if
			End With     
			j = j+ 1    
        next


         buyPer= 0 
			if (clsCMeachulLog.FMP_sum) > 0 then 
			buyPer = ((clsCMeachulLog.FMPPF_sum- totvc- clsCMeachulLog.FTotMF_10 + clsCMeachulLog.FTotMF_P )/clsCMeachulLog.FMP_sum)*100
			end if
         .Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", -1 
			.Add "구분", "공헌이익 2" 
			.Add "전체_금액", clsCMeachulLog.FMPPF_sum- totvc-(clsCMeachulLog.FTotMF_10 + clsCMeachulLog.FTotMF_P) '취급액수익-변동비1-변동비2
			.Add "전체_율", buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) and itotCnt>0 then 
                for intC= 0 To ubound(arrcate,2)    
                    if clsCMeachulLog.FcateMP_Sum(arrcate(3,intc)) > 0 then  
                    cper = ((clsCMeachulLog.FcateMPPF_Sum(arrcate(3,intc))- clsCMeachulLog.FTotVC1_Cate(arrcate(3,intC))- clsCMeachulLog.FTotVC2_Cate(arrcate(3,intC)))/clsCMeachulLog.FcateMP_Sum(arrcate(3,intc)))*100
                    end if
                    .Add arrcate(1,intC)&"_금액",  (clsCMeachulLog.FcateMPPF_Sum(arrcate(3,intc))- clsCMeachulLog.FTotVC1_Cate(arrcate(3,intC))- clsCMeachulLog.FTotVC2_Cate(arrcate(3,intC)))
                   .Add arrcate(1,intC)&"_율",cper/100
                next
             end if   
			end if
			End With    
            hidx1 = J
			j = j+ 1   


          buyPer= 0 
			if clsCMeachulLog.FMP_10 > 0 then 
			buyPer = ((clsCMeachulLog.FMPPF_10-totvc_10-clsCMeachulLog.FTotMF_10)/clsCMeachulLog.FMP_10)*100
			end if
			.Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx1
			.Add "구분", "공헌이익 2(10x10)" 
			.Add "전체_금액", clsCMeachulLog.FMPPF_10-totvc_10-clsCMeachulLog.FTotMF_10
			.Add "전체_율", buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) and iTotCnt > 0  then 
                for intC= 0 To ubound(arrcate,2) 
                    if clsCMeachulLog.FcateMP_10(arrcate(3,intc)) > 0 then
                        cper = ((clsCMeachulLog.FcateMPPF_10(arrcate(3,intc))-clsCMeachulLog.FTotVC1_Cate_10(arrcate(3,intC))-clsCMeachulLog.FTotVC2_Cate_10(arrcate(3,intC)))/clsCMeachulLog.FcateMP_10(arrcate(3,intc)))*100
                    end if
			        .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FcateMPPF_10(arrcate(3,intc))-clsCMeachulLog.FTotVC1_Cate_10(arrcate(3,intC))-clsCMeachulLog.FTotVC2_Cate_10(arrcate(3,intC))
                    .Add arrcate(1,intC)&"_율", cper/100 
                next
             end if   
			end if
			End With   
			j = j+ 1


             buyPer= 0 
            if  clsCMeachulLog.FMP_P > 0 then 
                buyPer = ((clsCMeachulLog.FMPPF_P-totvc_P-clsCMeachulLog.FTotMF_P )/ clsCMeachulLog.FMP_P)*100
            end if
         .Add j, tmpJSON.Collection()
			With .item(j) 
            .Add "ID", j 
            .Add "Head_ID", hidx1
			.Add "구분", "공헌이익 2(제휴)"  
			.Add "전체_금액", clsCMeachulLog.FMPPF_P-totvc_P-clsCMeachulLog.FTotMF_P 
			.Add "전체_율",  buyPer/100
			if sdispcate ="Y" then
             if isArray(arrcate) and itotCnt > 0 then 
                for intC= 0 To ubound(arrcate,2)  
                    if clsCMeachulLog.FcateMP_P(arrcate(3,intc)) > 0 then  
                    cper = ((clsCMeachulLog.FcateMPPF_P(arrcate(3,intc))-clsCMeachulLog.FTotVC1_Cate_P(arrcate(3,intC))-clsCMeachulLog.FTotVC2_Cate_P(arrcate(3,intC)))/clsCMeachulLog.FcateMP_P(arrcate(3,intc)))*100
                    end if
                    .Add arrcate(1,intC)&"_금액", clsCMeachulLog.FcateMPPF_P(arrcate(3,intc))-clsCMeachulLog.FTotVC1_Cate_P(arrcate(3,intC))-clsCMeachulLog.FTotVC2_Cate_P(arrcate(3,intC))
                    .Add arrcate(1,intC)&"_율",cper/100
                next
             end if   
			end if
			End With     
			j = j+ 1 
    set clsCMeachulLog  = nothing 
    End With
	Response.Write tmpJSON.JSONoutput() 
	
	Set tmpJSON = Nothing
 
 
 %> 
 <!-- #include virtual="/lib/db/dbSTSclose.asp" -->
 <!-- #include virtual="/lib/db/db3close.asp" --> 
 <!-- #include virtual="/lib/db/dbclose.asp" -->