<%
'######################################
' 공헌이익 Class
'######################################
'매입구분에 따른 금액 부분명 가져오기
Function fnGetMoneyType(ByVal mwdiv)
dim moneytype
 if mwdiv ="M" then 
  moneytype = "상품"
 elseif mwdiv = "U" or mwdiv ="W" then
  moneytype = "수수료"
 else
  moneytype = "기타"
 end IF 
 fnGetMoneyType = moneytype
End Function



class CMeachulLog

  public FdateType
  public FstDate
  public FedDate
  public FDispCate
  public Fcatekind

  public FTotCnt

  public FItem_sum 
  public FItem_10
  public FItem_10I
  public FItem_10C
  public FItem_10E 
  public FItem_p
  public FItem_pI
  public FItem_pC
  public FItem_pE

  public FOds_sum
  public FOds_10
  public FOds_10I
  public FOds_10C
  public FOds_10E 
  public FOds_p
  public FOds_pI
  public FOds_pC
  public FOds_pE 

  public Fbuy_sum
  public Fbuy_10
  public Fbuy_10I
  public Fbuy_10C
  public Fbuy_10E 
  public Fbuy_p
  public Fbuy_pI
  public Fbuy_pC
  public Fbuy_pE 
  
  public FbuyPF_sum
  public FbuyPF_10
  public FbuyPF_10I
  public FbuyPF_10C
  public FbuyPF_10E 
  public FbuyPF_p
  public FbuyPF_pI
  public FbuyPF_pC
  public FbuyPF_pE
   
  public FBC_sum
  public FBC_10
  public FBC_p
  public FBC_10I
  public FBC_pI
  public FBC_10C
  public FBC_pC
  public FBC_10E
  public FBC_pE
 

  public FMP_sum
  public FMP_10
  public FMP_p
  public FMP_10I
  public FMP_pI
  public FMP_10C
  public FMP_pC
  public FMP_10E
  public FMP_pE
 

  public FMPPF_sum
  public FMPPF_10
  public FMPPF_p
  public FMPPF_10I
  public FMPPF_pI
  public FMPPF_10C
  public FMPPF_pC
  public FMPPF_10E
  public FMPPF_pE
 

  public FDDate()
  public Fsitename()
  public Fmwdiv()
  public FItem()
  public FOds()
  public Fbuy()
  public FbuyPF()
  public FBC()
  public FMP()
  public FMPPF()

  public FBuyPer_10
  public FBuyPer_p
  public FBCPer_10
  public FBCPer_p
  public FMPricePer_10
  public FMPricePer_p
  public FMPerPer_10
  public FMPerPer_p

  public Fcatecode()
  public Fcatename()
  public Fcateno()
  public Fcatebuy()
  public FcatebuyPF()
  public FcateBC()
  public FcateMP()
  public FcateMPPF()
  public FcateItem()

  public Fcatebuy_Sum()
  public FcatebuyPF_Sum()
  public FcateBC_Sum()
  public FcateMP_Sum()
  public FcateMPPF_Sum()
  public FcateItem_Sum()
          
  public Fcatebuy_10()
  public FcatebuyPF_10()
  public FcateBC_10()
  public FcateMP_10()
  public FcateMPPF_10()
  public FcateItem_10()

  public Fcatebuy_P()
  public FcatebuyPF_P()
  public FcateBC_P()
  public FcateMP_P()
  public FcateMPPF_P()
  public FcateItem_P()

  public Fcatebuy_10I()
  public FcatebuyPF_10I()
  public FcateBC_10I()
  public FcateMP_10I()
  public FcateMPPF_10I()
  public FcateItem_10I()

  public Fcatebuy_PI()
  public FcatebuyPF_PI()
  public FcateBC_PI()
  public FcateMP_PI()
  public FcateMPPF_PI()
  public FcateItem_PI()

  public Fcatebuy_10C()
  public FcatebuyPF_10C()
  public FcateBC_10C()
  public FcateMP_10C()
  public FcateMPPF_10C()
  public FcateItem_10C()

  public Fcatebuy_PC()
  public FcatebuyPF_PC()
  public FcateBC_PC()
  public FcateMP_PC()
  public FcateMPPF_PC()
  public FcateItem_PC()

  public Fcatebuy_10E()
  public FcatebuyPF_10E()
  public FcateBC_10E()
  public FcateMP_10E()
  public FcateMPPF_10E()
  public FcateItem_10E()

  public Fcatebuy_PE()
  public FcatebuyPF_PE()
  public FcateBC_PE()
  public FcateMP_PE()
  public FcateMPPF_PE()
  public FcateItem_PE()
 
  public FCateCnt
  public FCateRow
  public Farrcate

  public FcatecodeSearch
  public ForderGubun
  public FstartYearMonth
  public FendYearMonth

'카테고리 리스트 가져오기
Function fnGetCateList
 dim strSql 
 dim arrList
  if Fcatekind ="" then Fcatekind ="D"
    strSql = "exec db_statistics.[dbo].usp_Ten_OrderLog_getCategory '"&FdateType&"','"&FstDate&"','"&FedDate&"','"&Fcatekind&"','"&FcatecodeSearch&"' "    
    rsSTSget.CursorLocation = adUseClient
    rsSTSget.open strSql,dbSTSget,adOpenForwardOnly, adLockReadOnly
   
		IF not rsSTSget.EOF THEN
           fnGetCateList = rsSTSget.getRows() 
        END IF
    rsSTSget.close 
End Function

    '매출총액 가져오기
  public Function fnGetOrerLogData
    Dim strSql
    dim arrList , intLoop 
    dim arrcate, intc
    dim vatinclude
    FTotCnt = 0 
	strSql = "exec db_statistics.[dbo].usp_Ten_orderLog_getMeachulData '"&FdateType&"','"&FstDate&"','"&FedDate&"','"&FDispCate&"','"&Fcatekind&"' "    
  	rsSTSget.CursorLocation = adUseClient
    rsSTSget.open strSql,dbSTSget,adOpenForwardOnly, adLockReadOnly 
'	rsSTSget.Open strSql, dbSTSget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsSTSget.EOF OR rsSTSget.BOF) THEN
			arrList = rsSTSget.getRows()
		END IF
		rsSTSget.close  
     

  IF isArray(arrList) then
  
    FTotCnt = Ubound(arrList,2) 

    FItem_sum = 0
    FOds_sum = 0
    Fbuy_sum = 0
    Fbuy_10  = 0
    Fbuy_10I = 0
    Fbuy_10C = 0
    Fbuy_10E = 0
    Fbuy_p   = 0 
    Fbuy_pI   = 0 
    Fbuy_pC   = 0 
    Fbuy_pE   = 0 

    FbuyPer_10 = 0
    FbuyPer_p = 0

    FbuyPF_sum = 0
    FbuyPF_10  = 0
    FbuyPF_10I = 0
    FbuyPF_10C = 0
    FbuyPF_10E = 0
    FbuyPF_p   = 0 
    FbuyPF_pI   = 0 
    FbuyPF_pC   = 0 
    FbuyPF_pE   = 0 

    FBC_sum = 0
    FBC_10 = 0
    FBC_p = 0 
    FBC_10I = 0
    FBC_pI = 0 
    FBC_10C = 0
    FBC_pC = 0 
    FBC_10E = 0
    FBC_pE = 0 

    FBCPer_10 = 0
    FBCPer_p = 0 

    FMP_sum = 0
    FMP_10 = 0
    FMP_p = 0 
    FMP_10I = 0
    FMP_pI = 0 
    FMP_10C = 0
    FMP_pC = 0 
    FMP_10 = 0
    FMP_p = 0 
    
    FMPPF_sum = 0
    FMPPF_10 = 0
    FMPPF_p = 0 
    FMPPF_10I = 0
    FMPPF_pI = 0 
    FMPPF_10C = 0
    FMPPF_pC = 0 
    FMPPF_10E = 0
    FMPPF_pE = 0 

    redim FDDate(FTotCnt)
    redim Fsitename(FTotCnt)
    redim Fmwdiv(FTotCnt)
    redim Fbuy(FTotCnt)
    redim FbuyPF(FTotCnt)
    redim FBC(FTotCnt)
    redim FMP(FTotCnt)
    redim FMPPF(FTotCnt)
    redim FItem(FTotCnt)
    redim FOds(FTotCnt)
   
    dim icateno
    dim ic : ic = 0
    dim i : i = 0
    dim oldmwdiv : oldmwdiv = ""
    dim oldsitename : oldsitename = ""
    
    if  FDispCate="Y" then  

        Farrcate =  fnGetCateList
        FCateCnt = ubound(Farrcate,2) +1 
        FCateRow =  arrList(12,0)-1
 
        redim Fcatecode(FTotCnt)
        redim Fcatename(FTotCnt) 
        redim Fcateno(FTotCnt) 
       
        redim Fcatebuy(FCateRow,FCateCnt) 
        redim FcatebuyPF(FCateRow,FCateCnt) 
        redim FcateBC(FCateRow,FCateCnt) 
        redim FcateMP(FCateRow,FCateCnt) 
        redim FcateMPPF(FCateRow,FCateCnt) 
        redim FcateItem(FCateRow,FCateCnt)

        redim Fcatebuy_Sum(FCateCnt)
        redim FcatebuyPF_Sum(FCateCnt)
        redim FcateBC_Sum(FCateCnt)
        redim FcateMP_Sum(FCateCnt)
        redim FcateMPPF_Sum(FCateCnt)
        redim FcateItem_Sum(FCateCnt)

        redim Fcatebuy_10(FCateCnt)
        redim FcatebuyPF_10(FCateCnt)
        redim FcateBC_10(FCateCnt)
        redim FcateMP_10(FCateCnt)
        redim FcateMPPF_10(FCateCnt)
        redim FcateItem_10(FCateCnt)

        redim Fcatebuy_P(FCateCnt)
        redim FcatebuyPF_P(FCateCnt)
        redim FcateBC_P(FCateCnt)
        redim FcateMP_P(FCateCnt)
        redim FcateMPPF_P(FCateCnt)
        redim FcateItem_P(FCateCnt)

        redim Fcatebuy_10I(FCateCnt)
        redim FcatebuyPF_10I(FCateCnt)
        redim FcateBC_10I(FCateCnt)
        redim FcateMP_10I(FCateCnt)
        redim FcateMPPF_10I(FCateCnt)
        redim FcateItem_10I(FCateCnt)

        redim Fcatebuy_PI(FCateCnt)
        redim FcatebuyPF_PI(FCateCnt)
        redim FcateBC_PI(FCateCnt)
        redim FcateMP_PI(FCateCnt)
        redim FcateMPPF_PI(FCateCnt)
        redim FcateItem_PI(FCateCnt)

        redim Fcatebuy_10C(FCateCnt)
        redim FcatebuyPF_10C(FCateCnt)
        redim FcateBC_10C(FCateCnt)
        redim FcateMP_10C(FCateCnt)
        redim FcateMPPF_10C(FCateCnt)
        redim FcateItem_10C(FCateCnt)

        redim Fcatebuy_PC(FCateCnt)
        redim FcatebuyPF_PC(FCateCnt)
        redim FcateBC_PC(FCateCnt)
        redim FcateMP_PC(FCateCnt)
        redim FcateMPPF_PC(FCateCnt)
        redim FcateItem_PC(FCateCnt)

        redim Fcatebuy_10E(FCateCnt)
        redim FcatebuyPF_10E(FCateCnt)
        redim FcateBC_10E(FCateCnt)
        redim FcateMP_10E(FCateCnt)
        redim FcateMPPF_10E(FCateCnt)
        redim FcateItem_10E(FCateCnt)

        redim Fcatebuy_PE(FCateCnt)
        redim FcatebuyPF_PE(FCateCnt)
        redim FcateBC_PE(FCateCnt)
        redim FcateMP_PE(FCateCnt)
        redim FcateMPPF_PE(FCateCnt)
        redim FcateItem_PE(FCateCnt)

       ' 초기값 설정
        dim a,b
        for a= 0 to FCateRow
          for b = 1 to FCateCnt
            Fcatebuy(a,b) = 0
            FcatebuyPF(a,b) = 0
            FcateBC(a,b) = 0
            FcateMP(a,b) = 0
            FcateMPPF(a,b) = 0
            FcateItem(a,b) = 0
          next
        next   

        for b = 1 to FCateCnt
            Fcatebuy_Sum(b) =  0
            FcatebuyPF_Sum(b) =  0
            FcateBC_Sum(b)=0
            FcateMP_Sum(b)=0
            FcateMPPF_Sum(b)=0
            FcateItem_Sum(b) = 0

            Fcatebuy_10(b) =0
            FcatebuyPF_10(b) =0
            FcateBC_10(b)=0
            FcateMP_10(b)=0
            FcateMPPF_10(b)=0
            FcateItem_10(b) =0

            Fcatebuy_P(b) =0
            FcatebuyPF_P(b) =0
            FcateBC_P(b)=0
            FcateMP_P(b)=0
            FcateMPPF_P(b)=0
            FcateItem_P(b) =0

            Fcatebuy_10I(b) =0
            FcatebuyPF_10I(b) =0
            FcateBC_10I(b)=0
            FcateMP_10I(b)=0
            FcateMPPF_10I(b)=0
            FcateItem_10I(b) =0

            Fcatebuy_PI(b) =0
            FcatebuyPF_PI(b) =0
            FcateBC_PI(b)=0
            FcateMP_PI(b)=0
            FcateMPPF_PI(b)=0
            FcateItem_PI(b) =0

            Fcatebuy_10C(b) =0
            FcatebuyPF_10C(b) =0
            FcateBC_10C(b)=0
            FcateMP_10C(b)=0
            FcateMPPF_10C(b)=0
            FcateItem_10C(b) =0

            Fcatebuy_PC(b) =0
            FcatebuyPF_PC(b) =0
            FcateBC_PC(b)=0
            FcateMP_PC(b)=0
            FcateMPPF_PC(b)=0
            FcateItem_PC(b) =0

            Fcatebuy_10E(b) =0
            FcatebuyPF_10E(b) =0
            FcateBC_10E(b)=0
            FcateMP_10E(b)=0
            FcateMPPF_10E(b)=0
            FcateItem_10E(b) =0

            Fcatebuy_PE(b) =0
            FcatebuyPF_PE(b) =0
            FcateBC_PE(b)=0
            FcateMP_PE(b)=0
            FcateMPPF_PE(b)=0
            FcateItem_PE(b) =0
        next
     end if    


    For intLoop = 0 To Ubound(arrList,2) 
        if FDispCate="Y" then 
        
            Fcatecode(intLoop) = arrList(10,intLoop)
            Fcatename(intLoop) = arrList(11,intLoop)  
            Fcateno(intLoop)   = arrList(13,intLoop)  
 
            '카테고리는 매입구분값으로 배열처리 
            if intLoop >0 then
              if (arrList(2,intLoop) <> arrList(2,intLoop-1)) or (arrList(1,intLoop) <> arrList(1,intLoop-1)) then
                i = i+ 1
                Fitem(i) = 0
                FOds(i)=0
                Fbuy(i)  = 0 
                FbuyPF(i)  = 0 
                FBC(i)   = 0
                FMP(i)   = 0
                FMPPF(i) = 0
                Fsitename(i) = arrList(1,intLoop)
                Fmwdiv(i)    = arrList(2,intLoop)
              end if
            else
                i = 0
                Fitem(i) = 0
                FOds(i)=0
                Fbuy(i)  = 0 
                FbuyPF(i)  = 0 
                FBC(i)   = 0
                FMP(i)   = 0
                FMPPF(i) = 0
                Fsitename(i) = arrList(1,intLoop)
                Fmwdiv(i)    = arrList(2,intLoop)
            end if

           
            Fbuy(i) =  Fbuy(i) + arrList(3,intLoop)
            FbuyPF(i) =  FbuyPF(i) + cdbl(arrList(4,intLoop))
            FBC(i) = FBC(i) + arrList(5,intLoop)
            FMP(i) = FMP(i)  + arrList(6,intLoop)
            FMPPF(i)= FMPPF(i) +cdbl(arrList(7,intLoop)) 
            FItem(i) = FItem(i) + arrList(8, intLoop)
            FOds(i) = FOds(i) + arrList(9, intLoop)

            Fcatebuy(i,Fcateno(intLoop)) =  arrList(3,intLoop) 
            FcatebuyPF(i,Fcateno(intLoop)) =  cdbl(arrList(4,intLoop))
            FcateBC(i,Fcateno(intLoop))  =  arrList(5,intLoop)  
            FcateMP(i,Fcateno(intLoop))  =  arrList(6,intLoop)
            FcateMPPF(i,Fcateno(intLoop))= cdbl(arrList(7,intLoop))     
            FcateItem(i,Fcateno(intLoop)) = arrList(8,intLoop)
       
        else
            FDDate(intLoop)    = arrList(0,intLoop)
            Fsitename(intLoop) = arrList(1,intLoop)
            Fmwdiv(intLoop)    = arrList(2,intLoop) 
            Fbuy(intLoop)      = arrList(3,intLoop) '구매총액 
            FbuyPF(intLoop)    = cdbl(arrList(4,intLoop)) '구매총액 수익
            FBC(intLoop)       = arrList(5,intLoop) '보너스쿠폰액 
            FMP(intLoop)       = arrList(6,intLoop) '취급액 
            FMPPF(intLoop)     = cdbl(arrList(7,intLoop)) '취급액 수익  
            FItem(intLoop)     = arrList(8,intLoop)
            FOds(intLoop)     = arrList(9,intLoop)
        end if

      ' 구분별 총액
         Fbuy_sum = Fbuy_sum + arrList(3,intLoop)
         FbuyPF_sum = FbuyPF_sum + cdbl(arrList(4,intLoop))
         FBC_sum  = FBC_sum +  arrList(5,intLoop)
         FMP_sum = FMP_sum + arrList(6,intLoop) 
         FMPPF_sum = FMPPF_sum + cdbl(arrList(7,intLoop)) 
         
         FItem_sum = FItem_sum + arrList(8,intLoop)
         FOds_sum = FOds_sum + arrList(9,intLoop) 
         
      '매출처별 총액   
        if  arrList(1,intLoop) ="10x10" then
          Fbuy_10    = Fbuy_10 + arrList(3,intLoop)
          FbuyPF_10 = FbuyPF_10 + cdbl(arrList(4,intLoop))
          FBC_10     = FBC_10 + arrList(5,intLoop)
          FMP_10     = FMP_10 + arrList(6,intLoop) 
          FMPPF_10   = FMPPF_10 + cdbl(arrList(7,intLoop)) 
         
          FItem_10   = FItem_10 + arrList(8,intLoop)
          FOds_10   = FOds_10 + arrList(9,intLoop)
 
          if arrList(2,intLoop) ="M" then
            Fbuy_10I    = Fbuy_10I + arrList(3,intLoop)
            FbuyPF_10I = FbuyPF_10I + cdbl(arrList(4,intLoop))
            FBC_10I     = FBC_10I + arrList(5,intLoop)
            FMP_10I     = FMP_10I + arrList(6,intLoop) 
            FMPPF_10I  = FMPPF_10I + cdbl(arrList(7,intLoop)) 

            FItem_10I   = FItem_10I + arrList(8,intLoop)
            FOds_10I   = FOds_10I + arrList(9,intLoop)
          elseif arrList(2,intLoop) ="W" or arrList(2,intLoop) ="U" then
            Fbuy_10C    = Fbuy_10C + arrList(3,intLoop) 
            FbuyPF_10C = FbuyPF_10C + cdbl(arrList(4,intLoop))
            FBC_10C     = FBC_10C + arrList(5,intLoop)
            FMP_10C     = FMP_10C + arrList(6,intLoop) 
            FMPPF_10C   = FMPPF_10C + cdbl(arrList(7,intLoop)) 
            FItem_10C   = FItem_10C + arrList(8,intLoop)
            FOds_10C   = FOds_10C + arrList(9,intLoop)
          else
            Fbuy_10E    = Fbuy_10E + arrList(3,intLoop)
            FbuyPF_10E = FbuyPF_10E + cdbl(arrList(4,intLoop))
            FBC_10E     = FBC_10E + arrList(5,intLoop)
            FMP_10E     = FMP_10E + arrList(6,intLoop) 
            FMPPF_10E   = FMPPF_10E + cdbl(arrList(7,intLoop)) 
            FItem_10E   = FItem_10E + arrList(8,intLoop)
            FOds_10E  = FOds_10E + arrList(9,intLoop)
          end if 
        else
          Fbuy_p   =  Fbuy_p + arrList(3,intLoop)
          FbuyPF_p = FbuyPF_p + cdbl(arrList(4,intLoop))
          FBC_p    = FBC_p +  arrList(5,intLoop)
          FMP_p    = FMP_p +  arrList(6,intLoop) 
          FMPPF_p  = FMPPF_p + cdbl(arrList(7,intLoop)) 
          FItem_p   = FItem_p + arrList(8,intLoop)
          FOds_p   = FOds_p + arrList(9,intLoop)

          if arrList(2,intLoop) ="M" then
            Fbuy_PI    = Fbuy_PI + arrList(3,intLoop)
            FbuyPF_PI = FbuyPF_PI + cdbl(arrList(4,intLoop))
            FBC_PI     = FBC_PI + arrList(5,intLoop)
            FMP_PI     = FMP_PI + arrList(6,intLoop) 
            FMPPF_PI  = FMPPF_PI + cdbl(arrList(7,intLoop)) 
            FItem_PI   = FItem_PI + arrList(8,intLoop)
            FOds_PI   = FOds_PI + arrList(9,intLoop)

          elseif arrList(2,intLoop) ="W" or arrList(2,intLoop) ="U" then
            Fbuy_PC    = Fbuy_PC + arrList(3,intLoop)
            FbuyPF_PC  = FbuyPF_PC + cdbl(arrList(4,intLoop))
            FBC_PC     = FBC_PC + arrList(5,intLoop)
            FMP_PC     = FMP_PC + arrList(6,intLoop) 
            FMPPF_PC   = FMPPF_PC + cdbl(arrList(7,intLoop)) 
            FItem_PC   = FItem_PC + arrList(8,intLoop)
            FOds_PC   = FOds_PC + arrList(9,intLoop)

          else
            Fbuy_PE    = Fbuy_PE + arrList(3,intLoop)
            FbuyPF_PE  = FbuyPF_PE + cdbl(arrList(4,intLoop))
            FBC_PE     = FBC_PE + arrList(5,intLoop)
            FMP_PE     = FMP_PE + arrList(6,intLoop) 
            FMPPF_PE   = FMPPF_PE + cdbl(arrList(7,intLoop)) 
            FItem_PE   = FItem_PE + arrList(8,intLoop)
            FOds_PE   = FOds_PE + arrList(9,intLoop)

          end if  
        end if 
      
    Next 

  

      '카테고리별 총액 
       if FDispCate="Y" then  
         for ic =1 to FCateCnt
          for i = 0 to FCateRow
            Fcatebuy_Sum(ic) =  Fcatebuy_Sum(ic) +  Fcatebuy(i,ic)  
            FcatebuyPF_Sum(ic) =  FcatebuyPF_Sum(ic) +  FcatebuyPF(i,ic) 
            FcateBC_Sum(ic)=FcateBC_Sum(ic) + FcateBC(i,ic)
            FcateMP_Sum(ic)=FcateMP_Sum(ic) + FcateMP(i,ic)
            FcateMPPF_Sum(ic)=FcateMPPF_Sum(ic) + FcateMPPF(i,ic)
            FcateItem_Sum(ic) =  FcateItem_Sum(ic) +  FcateItem(i,ic)

          if  Fsitename(i) = "10x10" then
            Fcatebuy_10(ic) =Fcatebuy_10(ic)+  Fcatebuy(i,ic) 
            FcatebuyPF_10(ic) =FcatebuyPF_10(ic)+  FcatebuyPF(i,ic) 
            FcateBC_10(ic)=FcateBC_10(ic) + FcateBC(i,ic)
            FcateMP_10(ic)=FcateMP_10(ic)+ FcateMP(i,ic)
            FcateMPPF_10(ic)=FcateMPPF_10(ic)+ FcateMPPF(i,ic)
            FcateItem_10(ic) =FcateItem_10(ic)+  FcateItem(i,ic) 
               if Fmwdiv(i) ="M" then
                  Fcatebuy_10I(ic) =Fcatebuy_10I(ic)+  Fcatebuy(i,ic) 
                  FcatebuyPF_10I(ic) =FcatebuyPF_10I(ic)+  FcatebuyPF(i,ic) 
                  FcateBC_10I(ic)=FcateBC_10I(ic) + FcateBC(i,ic)
                  FcateMP_10I(ic)=FcateMP_10I(ic)+ FcateMP(i,ic)
                  FcateMPPF_10I(ic)=FcateMPPF_10I(ic)+ FcateMPPF(i,ic)
                  FcateItem_10I(ic) =FcateItem_10I(ic)+  FcateItem(i,ic) 
                elseif Fmwdiv(i) ="W" or Fmwdiv(i) ="U" then
                  Fcatebuy_10C(ic) =Fcatebuy_10C(ic)+  Fcatebuy(i,ic)  
                  FcatebuyPF_10C(ic) =FcatebuyPF_10C(ic)+  FcatebuyPF(i,ic) 
                  FcateBC_10C(ic)=FcateBC_10C(ic) + FcateBC(i,ic)
                  FcateMP_10C(ic)=FcateMP_10C(ic)+ FcateMP(i,ic)
                  FcateMPPF_10C(ic)=FcateMPPF_10C(ic)+ FcateMPPF(i,ic)
                  FcateItem_10C(ic) =FcateItem_10C(ic)+  FcateItem(i,ic) 
                else
                  Fcatebuy_10E(ic) =Fcatebuy_10E(ic)+  Fcatebuy(i,ic) 
                  FcatebuyPF_10E(ic) =FcatebuyPF_10E(ic)+  FcatebuyPF(i,ic) 
                  FcateBC_10E(ic)=FcateBC_10E(ic) + FcateBC(i,ic)
                  FcateMP_10E(ic)=FcateMP_10E(ic)+ FcateMP(i,ic)
                  FcateMPPF_10E(ic)=FcateMPPF_10E(ic)+ FcateMPPF(i,ic)
                  FcateItem_10E(ic) =FcateItem_10E(ic)+  FcateItem(i,ic) 
                end if  
          else
            Fcatebuy_P(ic) =Fcatebuy_P(ic)+  Fcatebuy(i,ic) 
            FcatebuyPF_P(ic) =FcatebuyPF_P(ic)+  FcatebuyPF(i,ic) 
            FcateBC_P(ic)=FcateBC_P(ic) + FcateBC(i,ic)
            FcateMP_P(ic)=FcateMP_P(ic)+ FcateMP(i,ic)
            FcateMPPF_P(ic)=FcateMPPF_P(ic)+ FcateMPPF(i,ic)
            FcateItem_P(ic) =FcateItem_P(ic)+  FcateItem(i,ic) 
                 if Fmwdiv(i) ="M" then
                  Fcatebuy_PI(ic) =Fcatebuy_PI(ic)+  Fcatebuy(i,ic) 
                  FcatebuyPF_PI(ic) =FcatebuyPF_PI(ic)+  FcatebuyPF(i,ic) 
                  FcateBC_PI(ic)=FcateBC_PI(ic) + FcateBC(i,ic)
                  FcateMP_PI(ic)=FcateMP_PI(ic)+ FcateMP(i,ic)
                  FcateMPPF_PI(ic)=FcateMPPF_PI(ic)+ FcateMPPF(i,ic)
                  FcateItem_PI(ic) =FcateItem_PI(ic)+  FcateItem(i,ic) 
                elseif Fmwdiv(i) ="W" or Fmwdiv(i) ="U" then
                  Fcatebuy_PC(ic) =Fcatebuy_PC(ic)+  Fcatebuy(i,ic) 
                  FcatebuyPF_PC(ic) =FcatebuyPF_PC(ic)+  FcatebuyPF(i,ic) 
                  FcateBC_PC(ic)=FcateBC_PC(ic) + FcateBC(i,ic)
                  FcateMP_PC(ic)=FcateMP_PC(ic)+ FcateMP(i,ic)
                  FcateMPPF_PC(ic)=FcateMPPF_PC(ic)+ FcateMPPF(i,ic)
                  FcateItem_PC(ic) =FcateItem_PC(ic)+  FcateItem(i,ic) 
                else
                  Fcatebuy_PE(ic) =Fcatebuy_PE(ic)+  Fcatebuy(i,ic) 
                  FcatebuyPF_PE(ic) =FcatebuyPF_PE(ic)+  FcatebuyPF(i,ic) 
                  FcateBC_PE(ic)=FcateBC_PE(ic) + FcateBC(i,ic)
                  FcateMP_PE(ic)=FcateMP_PE(ic)+ FcateMP(i,ic)
                  FcateMPPF_PE(ic)=FcateMPPF_PE(ic)+ FcateMPPF(i,ic)
                  FcateItem_PE(ic) =FcateItem_PE(ic)+  FcateItem(i,ic) 
                end if  
          end if
          next
         next
       end if
 
     
  end IF
 
End Function 
 

 public FTotVC1   
 public FTotVC1_10   
 public FTotVC1_P  
 public FTotSC_10   
 public FTotSC_P 
 public FTotWH_10
 public FTotWH_P
 public FTotVC2
 public FTotVC2_10
 public FTotVC2_P
 public FTotMF_10
 public FTotMF_P

 public FTotSC_Cate_10()
 public FTotSC_Cate_P()
 public FTotWH_Cate_10()
 public FTotWH_Cate_P()
 public FTotVC1_Cate_10()
 public FTotVC1_Cate_P()
 public FTotVC1_Cate() 
 public FTotVC2_Cate_10()
 public FTotVC2_Cate_P()
 public FTotVC2_Cate()
 public FTotMF_Cate_10()
 public FTotMF_Cate_P()

 public FPFList

 public FAccCIdx()
 public FsiteNM()
 public FAccNM()
 public FPFPrice()
 public FCatePrice()
  public Fscrow
'안분데이터  저장내용 가져오기
public Function fnGetprofitlossdata
dim strSql   
dim  arrPFSum ,intLoop
dim ic, arrcate, catecnt
FTotVC1 = 0
FTotVC1_10 = 0
FTotVC1_P = 0
FTotSC_10 = 0
FTotSC_P = 0
FTotWH_10 = 0
FTotWH_P = 0
FTotVC2 = 0
FTotVC2_10 = 0
FTotVC2_P = 0
FTotMF_10 = 0
FTotMF_P = 0

if FDispCate ="Y" then
arrcate =  fnGetCateList
if isarray(arrcate) then
catecnt = ubound(arrcate,2)+1
else 
catecnt= 0  
end if
 redim FTotSC_Cate_10(catecnt)
 redim FTotSC_Cate_P(catecnt)
 redim FTotWH_Cate_10(catecnt)
 redim FTotWH_Cate_P(catecnt)
 redim FTotVC1_Cate_10(catecnt)
 redim FTotVC1_Cate_P(catecnt)
 redim FTotVC1_Cate(catecnt)
 redim FTotVC2_Cate_10(catecnt)
 redim FTotVC2_Cate_P(catecnt)
 redim FTotVC2_Cate(catecnt)
 redim FTotMF_Cate_10(catecnt)
 redim FTotMF_Cate_P(catecnt)
 if isArray(arrcate) then
 for ic = 0 to ubound(arrcate,2)
  FTotSC_Cate_10(ic) = 0
  FTotSC_Cate_P(ic) = 0
  FTotWH_Cate_10(ic) =0
  FTotWH_Cate_P(ic)=0
  FTotVC1_Cate_10(ic) = 0
  FTotVC1_Cate_P(ic) = 0
  FTotVC1_Cate(ic) = 0
  FTotVC2_Cate_10(ic) = 0
  FTotVC2_Cate_P(ic) = 0
  FTotVC2_Cate(ic) = 0
  FTotMF_Cate_10(ic) =0
  FTotMF_Cate_P(ic)=0
 next
 end if
 end if 
strSql = "exec db_datamart.dbo.usp_Ten_profitloss_getSumData '"&FstDate&"','"&FedDate&"','"&FDispCate&"','"&Fcatekind&"' " 
db3_rsget.CursorLocation = adUseClient
db3_rsget.open strSql,db3_dbget,adOpenForwardOnly, adLockReadOnly
   	if  not (db3_rsget.EOF  OR db3_rsget.BOF) then  
      arrPFSum = db3_rsget.getRows() 
    end if
db3_rsget.Close  

IF isArray(arrPFSum) then
  for intLoop = 0 To UBound(arrPFSum,2)
    if arrPFSum(0,intLoop)=5 then '변동비1>판매수수료
        if arrPFSum(1,intLoop) ="10x10" then
            if FDispCate = "N" then 
            FTotSC_10 = arrPFSum(2,intLoop)
            else 
                if isArray(arrcate) then
                    for ic = 0 to ubound(arrcate,2) 
                        if Cstr(arrcate(0,ic)) = Cstr(arrPFSum(2,intLoop)) then
                          FTotSC_Cate_10(arrcate(3,ic)) = arrPFSum(3,intLoop) 
                          FTotSC_10 = FTotSC_10 +  FTotSC_Cate_10(arrcate(3,ic)) 
                          FTotVC1_Cate_10(arrcate(3,ic)) = FTotVC1_Cate_10(arrcate(3,ic)) +  FTotSC_Cate_10(arrcate(3,ic)) 
                          FTotVC1_Cate(arrcate(3,ic)) = FTotVC1_Cate(arrcate(3,ic)) +  FTotSC_Cate_10(arrcate(3,ic)) 
                        end if   
                    next
                end if
            end if 
        else
            if FDispCate = "N" then
            FTotSC_P = arrPFSum(2,intLoop)
            else
                if isArray(arrcate) then
                    for ic = 0 to ubound(arrcate,2) 
                        if Cstr(arrcate(0,ic)) = Cstr(arrPFSum(2,intLoop)) then
                          FTotSC_Cate_P(arrcate(3,ic)) = arrPFSum(3,intLoop) 
                          FTotSC_P = FTotSC_P +  FTotSC_Cate_P(arrcate(3,ic))
                          FTotVC1_Cate_P(arrcate(3,ic)) = FTotVC1_Cate_P(arrcate(3,ic)) +  FTotSC_Cate_P(arrcate(3,ic)) 
                          FTotVC1_Cate(arrcate(3,ic)) = FTotVC1_Cate(arrcate(3,ic)) +  FTotSC_Cate_P(arrcate(3,ic)) 
                        end if   
                    next
                end if
            end if  
        end if  
    elseif  arrPFSum(0,intLoop)=6 then '변동비1>물류비
        if arrPFSum(1,intLoop) ="10x10" then
            if FDispCate = "N" then 
              FTotWH_10 = arrPFSum(2,intLoop)
            else 
              if isArray(arrcate) then
                    for ic = 0 to ubound(arrcate,2) 
                        if Cstr(arrcate(0,ic)) = Cstr(arrPFSum(2,intLoop)) then
                          FTotWH_Cate_10(arrcate(3,ic)) = arrPFSum(3,intLoop) 
                          FTotWH_10 = FTotWH_10 +  FTotWH_Cate_10(arrcate(3,ic)) 
                          FTotVC1_Cate_10(arrcate(3,ic)) = FTotVC1_Cate_10(arrcate(3,ic)) +  FTotWH_Cate_10(arrcate(3,ic)) 
                          FTotVC1_Cate(arrcate(3,ic)) = FTotVC1_Cate(arrcate(3,ic)) +  FTotWH_Cate_10(arrcate(3,ic)) 
                        end if   
                    next
                end if
            end if
        else
           if FDispCate = "N" then 
              FTotWH_P = arrPFSum(2,intLoop)
            else 
              if isArray(arrcate) then
                    for ic = 0 to ubound(arrcate,2) 
                        if Cstr(arrcate(0,ic)) = Cstr(arrPFSum(2,intLoop)) then
                          FTotWH_Cate_P(arrcate(3,ic)) = arrPFSum(3,intLoop) 
                          FTotWH_P = FTotWH_P +  FTotWH_Cate_P(arrcate(3,ic))
                          FTotVC1_Cate_P(arrcate(3,ic)) = FTotVC1_Cate_P(arrcate(3,ic)) +  FTotWH_Cate_P(arrcate(3,ic)) 
                          FTotVC1_Cate(arrcate(3,ic)) = FTotVC1_Cate(arrcate(3,ic)) +  FTotWH_Cate_P(arrcate(3,ic)) 
                        end if   
                    next
                end if
            end if
        end if   
    elseif  arrPFSum(0,intLoop)=9 then '변동비2>광고판촉비
         if arrPFSum(1,intLoop) ="10x10" then
            if FDispCate = "N" then 
              FTotMF_10 = arrPFSum(2,intLoop)
            else 
              if isArray(arrcate) then
                    for ic = 0 to ubound(arrcate,2) 
                        if Cstr(arrcate(0,ic)) = Cstr(arrPFSum(2,intLoop)) then
                          FTotMF_Cate_10(arrcate(3,ic)) = arrPFSum(3,intLoop) 
                          FTotMF_10 = FTotMF_10 +  FTotMF_Cate_10(arrcate(3,ic)) 
                          FTotVC2_Cate_10(arrcate(3,ic)) = FTotVC2_Cate_10(arrcate(3,ic)) +  FTotMF_Cate_10(arrcate(3,ic)) 
                          FTotVC2_Cate(arrcate(3,ic)) = FTotVC2_Cate(arrcate(3,ic)) +  FTotMF_Cate_10(arrcate(3,ic)) 
                        end if   
                    next
                end if
            end if
        else
           if FDispCate = "N" then 
              FTotMF_P = arrPFSum(2,intLoop)
            else 
              if isArray(arrcate) then
                    for ic = 0 to ubound(arrcate,2) 
                        if Cstr(arrcate(0,ic)) = Cstr(arrPFSum(2,intLoop)) then
                          FTotMF_Cate_P(arrcate(3,ic)) = arrPFSum(3,intLoop) 
                          FTotMF_P = FTotMF_P +  FTotMF_Cate_P(arrcate(3,ic))
                          FTotVC2_Cate_P(arrcate(3,ic)) = FTotVC2_Cate_P(arrcate(3,ic)) +  FTotMF_Cate_P(arrcate(3,ic)) 
                          FTotVC2_Cate(arrcate(3,ic)) = FTotVC2_Cate(arrcate(3,ic)) +  FTotMF_Cate_P(arrcate(3,ic)) 
                        end if   
                    next
                end if
            end if
        end if    
    end if
  next
end if

dim  arrList 
strSql = "exec db_datamart.dbo.usp_Ten_profitloss_getData '"&FstDate&"','"&FedDate&"','"&FDispCate&"' ,'"&Fcatekind&"'" 
db3_rsget.CursorLocation = adUseClient
db3_rsget.open strSql,db3_dbget,adOpenForwardOnly, adLockReadOnly
   	if  not (db3_rsget.EOF  OR db3_rsget.BOF) then  
      arrList = db3_rsget.getRows() 
    end if
db3_rsget.Close   
 
 dim scrow,irow
    scrow = 0
    irow= 0
if isArray(arrList) then
    if FDispCate = "N" then
       scrow = ubound(arrList,2) +1 
    else
       scrow = arrList(6,0)
       redim FCatePrice(scrow, catecnt)
    end if
     Fscrow = scrow
      redim FAccCIdx(scrow)
      redim FsiteNM(scrow)
      redim FAccNM(scrow)
      redim FPFPrice(scrow)

    for intLoop = 0  to ubound(arrList,2) 
       if FDispCate = "N" then  
             FAccCIdx(intLoop) = arrList(4,intLoop)
             FsiteNM(intLoop) = arrList(0,intLoop)
             FAccNM(intLoop) = arrList(2,intLoop)
             FPFPrice(intLoop) = arrList(3,intLoop)
        else 
             if intLoop > 0 then
                if arrList(1,intLoop) <> arrList(1,intLoop-1) then 
                  irow = irow + 1  
                  FPFPrice(irow)  = 0 
                end if
             else 
               FPFPrice(irow)    = 0
            end if
           
                  
                FAccCIdx(irow) = arrList(4,intLoop)
                FsiteNM(irow) = arrList(0,intLoop)
                FAccNM(irow) = arrList(2,intLoop)
                if isArray(arrcate) then
                  for ic = 0 to ubound(arrcate,2)  
                    if Cstr(arrcate(0,ic)) = Cstr(arrList(5,intLoop)) then
                      FCatePrice(irow,arrcate(3,ic))  = arrList(3,intLoop)
                      FPFPrice(irow)  = FPFPrice(irow) +  FCatePrice(irow,arrcate(3,ic)) 
                    end if  
                 next
                end if 
        end if
    next
end if

End Function

'공헌이익 총 매출 데이터 가져오기
Function fnGetTotalContributionProfit
  dim strSql 
  dim arrList
  strSql = " SELECT idx, YYYYMM, totalPurchase, totalPurchaseIncome, bonusCoupon, handllingAmount, handllingAmountIncome, productQuantity, numberOfOrders, variableCost1, variableCost2 "
  strSql = strSql & ", contributionProfit1, contributionProfit2, totalPurchaseRate, bonusCouponRate, handllingAmountRate, variableCostRate, contributionProfitRate, regdate, lastupdate "
  strSql = strSql & " FROM db_statistics.dbo.tbl_totalContributionProfit_stats WITH(NOLOCK) WHERE 1=1 "
  If FstartYearMonth <> "" or FendYearMonth <> "" Then
    strSql = strSql & " And YYYYMM >= '"&FstartYearMonth&"' And YYYYMM <= '"&FendYearMonth&"' "
  End If
  strSql = strSql & " ORDER BY YYYYMM "&ForderGubun
  rsSTSget.CursorLocation = adUseClient
  rsSTSget.open strSql,dbSTSget,adOpenForwardOnly, adLockReadOnly

  IF not rsSTSget.EOF THEN
    fnGetTotalContributionProfit = rsSTSget.getRows() 
  END IF
  rsSTSget.close
End Function

'공헌이익 카테고리별 총 매출 데이터 가져오기
Function fnGetCategoryTotalContributionProfit

  If Trim(FcatecodeSearch)="" Then
    exit function
  End If

  dim strSql 
  dim arrList
  strSql = " SELECT idx, YYYYMM, totalPurchase, totalPurchaseIncome, bonusCoupon, handllingAmount, handllingAmountIncome, productQuantity, numberOfOrders, variableCost1, variableCost2 "
  strSql = strSql & ", contributionProfit1, contributionProfit2, totalPurchaseRate, bonusCouponRate, handllingAmountRate, variableCostRate, contributionProfitRate, regdate, lastupdate, catecode, catename "
  strSql = strSql & " FROM db_statistics.dbo.tbl_CategoryContributionProfit_stats WITH(NOLOCK) WHERE 1=1 "
  strSql = strSql & " And catecode = '"&FcatecodeSearch&"' "
  If FstartYearMonth <> "" or FendYearMonth <> "" Then
    strSql = strSql & " And YYYYMM >= '"&FstartYearMonth&"' And YYYYMM <= '"&FendYearMonth&"' "
  End If
  strSql = strSql & " ORDER BY YYYYMM "&ForderGubun
  rsSTSget.CursorLocation = adUseClient
  rsSTSget.open strSql,dbSTSget,adOpenForwardOnly, adLockReadOnly

  IF not rsSTSget.EOF THEN
    fnGetCategoryTotalContributionProfit = rsSTSget.getRows() 
  END IF
  rsSTSget.close
End Function

End Class
%> 