<%
'######################################
' 공헌이익 Class
'######################################

'카테고리 리스트 가져오기
Function fnGetCateList
 dim strSql 
 dim arrList
    strSql = " select catecode, catename from db_item.dbo.tbl_display_cate where useyn='Y' and depth=1 order by sortno "
    rsget.Open strSql,dbget
		IF not rsget.EOF THEN
            fnGetCateList = rsget.getRows()
        END IF
    rsget.close

End Function

class CMeachulLog

  public FdateType
  public FstDate
  public FedDate
  public FDispCate

  public FTotCnt
  public Fbuytot_sum
  public Fbuytot_10
  public FbuytotPer_10
  public Fbuytot_p
  public FbuytotPer_p

  public FBCtot_sum
  public FBCtot_10
  public FBCtot_p

  public FMPricetot_sum
  public FMPricetot_10
  public FMPricetot_p

  public FMPertot_sum
  public FMPertot_10
  public FMPertot_p

  public FDDate()
  public Fsitename()
  public Fmwdiv()
  public Fbuytot()
  public FBCtot()
  public FMPricetot()
  public FMPertot()

  public FBCtotPer_10
  public FBCtotPer_p
  public FMPricetotPer_10
  public FMPricetotPer_p
  public FMPertotPer_10
  public FMPertotPer_p

    '구매총액
  public Function fnGetOrerLogData
    Dim strSql
    dim arrList , intLoop 
    if FdateType ="" then FdateType ="DTLactdate"
    if FstDate ="" then FstDate =  "2018-09-01"
    if FedDate = "" then FedDate ="2018-09-30"
    if FDispCate ="" then FDispCate ="N"
    FTotCnt = 0
	strSql = "exec db_statistics.[dbo].usp_Ten_orderLog_getMeachulData '"&FdateType&"','"&FstDate&"','"&FedDate&"','"&FDispCate&"' "   
 
  	rsSTSget.CursorLocation = adUseClient
    rsSTSget.open strSql,dbSTSget,adOpenForwardOnly, adLockReadOnly
   
'	rsSTSget.Open strSql, dbSTSget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsSTSget.EOF OR rsSTSget.BOF) THEN
			arrList = rsSTSget.getRows()
		END IF
		rsSTSget.close  
  

  IF isArray(arrList) then
  
    FTotCnt = Ubound(arrList,2) 
    Fbuytot_sum = 0
    Fbuytot_10  = 0
    Fbuytot_p   = 0
    
    FbuytotPer_10 = 0
    FbuyTotPer_p = 0

    FBCtot_sum = 0
    FBCtot_10 = 0
    FBCtot_p = 0 
    
    FMPricetot_sum = 0
    FMPricetot_10 = 0
    FMPricetot_p = 0 
    
    FMPertot_sum = 0
    FMPertot_10 = 0
    FMPertot_p = 0 

    redim FDDate(FTotCnt)
    redim Fsitename(FTotCnt)
    redim Fmwdiv(FTotCnt)
    redim Fbuytot(FTotCnt)
    redim FBCtot(FTotCnt)
    redim FMPricetot(FTotCnt)
    redim FMPertot(FTotCnt)

    For intLoop = 0 To Ubound(arrList,2)
      if FDispCate="N" then
         FDDate(intLoop)= arrList(0,intLoop)
         Fsitename(intLoop) = arrList(1,intLoop)
         Fmwdiv(intLoop) = arrList(2,intLoop)
         Fbuytot(intLoop) = arrList(3,intLoop) '구매총액 
         FBCtot(intLoop) =arrList(4,intLoop) '보너스쿠폰액 
         FMPricetot(intLoop) = arrList(5,intLoop) '취급액 
         FMPertot(intLoop) = cdbl(arrList(6,intLoop)) '취급액 수익  
      end if
        Fbuytot_sum = Fbuytot_sum + Fbuytot(intLoop)
        FBCtot_sum = FBCtot_sum + FBCtot(intLoop)
        FMPricetot_sum = FMPricetot_sum + FMPricetot(intLoop)
        FMPertot_sum = FMPertot_sum + Cdbl(FMPertot(intLoop))
         
        if  Fsitename(intLoop) ="10x10" then
          Fbuytot_10 =  Fbuytot_10 + Fbuytot(intLoop)
          FBCtot_10 = FBCtot_10 + FBCtot(intLoop)
          FMPricetot_10 = FMPricetot_10 + FMPricetot(intLoop)
           FMPertot_10 = FMPertot_10 + Cdbl(FMPertot(intLoop)) 
        else
          Fbuytot_p =  Fbuytot_p + Fbuytot(intLoop)
          FBCtot_p = FBCtot_p + FBCtot(intLoop)
          FMPricetot_p= FMPricetot_p + FMPricetot(intLoop)
           FMPertot_p = FMPertot_p +Cdbl(FMPertot(intLoop))
        end if
     
    Next
  

      if Fbuytot_sum >0 then    
        FbuytotPer_10 = round((Fbuytot_10/Fbuytot_sum)*100)
        FbuyTotPer_p=  round((Fbuytot_p/Fbuytot_sum)*100)
      end if 

      if FBCtot_sum>0 then 
      FBCtotPer_10 = round((FBCtot_10/FBCtot_sum)*100)
      FBCtotPer_p= round((FBCtot_p/FBCtot_sum)*100)
      end if
    
      if FMPricetot_sum>0 then 
      FMPricetotPer_10 = round((FMPricetot_10/FMPricetot_sum)*100)
      FMPricetotPer_p= round((FMPricetot_p/FMPricetot_sum)*100)
      end if
      
      if FMPertot_sum>0 then 
      FMPertotPer_10 = round((FMPertot_10/FMPertot_sum)*100)
      FMPertotPer_p= round((FMPertot_p/FMPertot_sum)*100)
      end if

  end IF
 
End Function 

public FliComm10
public FliCommP
'변동비 - 라이센스 수수료
public Function fnGetLinceComm
Dim strSql
    dim arrList , intLoop 
    if FdateType ="" then FdateType ="DTLactdate" 
    if FDispCate ="" then FDispCate ="N"
    FliComm10 = 0
    FliCommP = 0
	strSql = "exec db_statistics.[dbo].usp_Ten_orderLog_getLicenseCommData '"&FdateType&"','"&FstDate&"','"&FedDate&"','"&FDispCate&"' "   
  
  	rsSTSget.CursorLocation = adUseClient
    rsSTSget.open strSql,dbSTSget,adOpenForwardOnly, adLockReadOnly
   
'	rsSTSget.Open strSql, dbSTSget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsSTSget.EOF OR rsSTSget.BOF) THEN
			 FliComm10 = rsSTSget("licomm10")
       FliCommP = rsSTSget("licommP")
		END IF
		rsSTSget.close 

End Function

public FpgComm
public FCardComm
public FCpsComm

'변동비 - 수수료
public Function fnGetComm
dim strSql
strSql = "exec db_partner.dbo.usp_Ten_profitloss_getCommData '"&FdateType&"','"&FstDate&"','"&FedDate&"','"&FDispCate&"' "
 
rsget.CursorLocation = adUseClient
rsget.open strSql,dbget,adOpenForwardOnly, adLockReadOnly
   	if  not (rsget.EOF  OR rsget.BOF) then 
      FpgComm = rsget("pgcomm")
      FCardComm = rsget("cardcomm")
      FCpsComm = rsget("cpscomm") 
    end if
rsget.Close     
End Function
End Class
%>