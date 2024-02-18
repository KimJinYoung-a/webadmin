<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbiTmsOpen.asp" -->
<!-- #include virtual="/lib/db/dbiTMSHelper.asp"-->  
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim divMastKey   : divMastKey =  requestCheckvar(request("divMastKey"),10)
Dim mode         : mode =  requestCheckvar(request("mode"),32)
Dim AssignYYYYMM : AssignYYYYMM =  requestCheckvar(request("AssignYYYYMM"),7)
Dim bizSecCd     : bizSecCd =  requestCheckvar(request("bizSecCd"),32)
Dim AssignAccUseCD : AssignAccUseCD =  requestCheckvar(request("AssignAccUseCD"),32)
Dim cust_cd      : cust_cd =  requestCheckvar(request("cust_cd"),32)
Dim isForceDel   : isForceDel =  requestCheckvar(request("isForceDel"),1)
Dim dtl_bizSecCd : dtl_bizSecCd =  requestCheckvar(request("dtl_bizSecCd"),1024)
Dim dtl_dlvPro : dtl_dlvPro =  requestCheckvar(request("dtl_dlvPro"),1024)

Dim SLTRKEY      : SLTRKEY =  requestCheckvar(request("SLTRKEY"),1024)
Dim SLTRKEY_SEQ  : SLTRKEY_SEQ =  requestCheckvar(request("SLTRKEY_SEQ"),1024)
Dim chk     : chk =  request("chk")
Dim iSLTRKEY, iSLTRKEY_SEQ, ichk

'rw divMastKey
'rw mode
'rw AssignYYYYMM
'rw bizSecCd
'rw cust_cd
'rw AssignAccUseCD
'rw dtl_bizSecCd
'rw dtl_dlvPro


Dim paramInfo, sqlStr
Dim retParamInfo, RetErr, retErrStr
Dim i, RetVal, idtl_bizSecCd, idtl_dlvPro

IF (mode="regDivMast") then
    '' 사용자ACC_CD가 존재하는지 check
    paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
            ,Array("@yyyymm"	,adVarchar, adParamInput,7, AssignYYYYMM) _
            ,Array("@bizSecCd"	,adVarchar, adParamInput,10, bizSecCd) _
            ,Array("@acc_use_cd" ,adVarchar, adParamInput,13, AssignAccUseCD) _
            ,Array("@cust_cd"	,adVarchar, adParamInput,13, cust_cd) _
            ,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
    )
    
    sqlStr = "db_SCM_LINK.dbo.sp_SCM2ERP_bizProfit_regMst" 
            	
	IF application("Svr_Info")="Dev" THEN
	    sqlStr = sqlStr&"_TEST"
    END IF
            
    retParamInfo = fnExecSPOutput(sqlStr,paramInfo)
   
    RetErr       = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드 or IDX
    retErrStr    = GetValue(retParamInfo, "@retErrStr")    ' 에러내역
    
    IF (RetErr<1) then
        rw "RetErr="&RetErr
        rw "retErrStr="&retErrStr
    ELSE
        dtl_bizSecCd = split(dtl_bizSecCd,",")
        dtl_dlvPro   = split(dtl_dlvPro,",")
        
        for i=0 to Ubound(dtl_bizSecCd)
            idtl_bizSecCd = Trim(dtl_bizSecCd(i))
            idtl_dlvPro   = Trim(dtl_dlvPro(i))
            
            ''rw idtl_bizSecCd & "|" & idtl_dlvPro
            IF (idtl_bizSecCd<>"") and (idtl_dlvPro<>"") then
                
                paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
                        ,Array("@divMastKey"	,adInteger, adParamInput,, RetErr) _
                        ,Array("@BIZSECTION_CD"	,adVarchar, adParamInput,10, idtl_bizSecCd) _
                        ,Array("@divPro" ,adVarchar, adParamInput,10,idtl_dlvPro) _
                        ,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
                )
                
                sqlStr = "db_SCM_LINK.dbo.sp_SCM2ERP_bizProfit_regDtl" 
                	
            	IF application("Svr_Info")="Dev" THEN
            	    sqlStr = sqlStr&"_TEST"
                END IF
        
                retParamInfo = fnExecSPOutput(sqlStr,paramInfo)
       
                RetVal       = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드 or IDX
                retErrStr    = GetValue(retParamInfo, "@retErrStr")    ' 에러내역
                
                rw retErrStr
            end if
        next
    ENd IF
ELSEIF (mode="regPreM") THEN '이전달 내역 가져와서 등록처리 
	Dim preyyyymm,predivMastKey
	Dim arrList, j,arrDtl
	Dim paramInfo1,retParamInfo1,RetErr1
	preyyyymm = dateadd("m",-1,AssignYYYYMM) '검색 이전달
    
   '1.검색달에 데이터 있는지 확인 - 데이터 있는  경우 복사하지 않는다. 
    paramInfo1 = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
         							,Array("@yyyymm"	,adVarchar, adParamInput,7, AssignYYYYMM)_
         				)
    IF (application("Svr_Info")="Dev") THEN
        sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitDivMasterExists_TEST"
    ELSE
        sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitDivMasterExists"
    END IF  
   			retParamInfo1 = fnExecSPOutput(sqlStr,paramInfo1) 
		    RetErr1       = GetValue(retParamInfo1, "@RETURN_VALUE") 
			IF RetErr1 = 1 THEN
					Call Alert_return ("선택하신 달에 기존 데이터가 존재합니다. 이전달 내역 가져오기가 불가능합니다.")
				response.end
			END IF 
    
    '2.이전 달 master 내역 가져오기
    IF (application("Svr_Info")="Dev") THEN
        sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitDivMasterList_TEST"
    ELSE
        sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitDivMasterList"
    END IF 
   		 sqlStr =sqlStr+"('"&preyyyymm&"','','','')"	  
    dbiTms_rsget.Open sqlStr, dbiTms_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc 
		IF  not (dbiTms_rsget.EOF OR dbiTms_rsget.BOF) THEN
				ArrList = dbiTms_rsget.getRows()
				dbiTms_rsget.close
    ELSE 
    	dbiTms_rsget.close
		  Call Alert_return ("이전달 내역이 존재하지 않습니다..확인 후 다시 처리해주세요")
		 response.end
		END IF
		
		IF isArray(arrList) THEN
		  For i=0 to ubound(arrList,2)  
			  predivMastKey		= arrList(0,i)
			  bizSecCd	  		= arrList(2,i)
			  AssignAccUseCD  = arrList(5,i) 
			  cust_cd  				= arrList(3,i) 
			  '3.현재달에 이전달 내역 for 문을 통해 하나씩 복사 등록
		  	 paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
            ,Array("@yyyymm"	,adVarchar, adParamInput,7, AssignYYYYMM) _
            ,Array("@bizSecCd"	,adVarchar, adParamInput,10, bizSecCd) _
            ,Array("@acc_use_cd" ,adVarchar, adParamInput,13, AssignAccUseCD) _
            ,Array("@cust_cd"	,adVarchar, adParamInput,13, cust_cd) _
            ,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
		   	 ) 
		    sqlStr = "db_SCM_LINK.dbo.sp_SCM2ERP_bizProfit_regMst"  
					IF application("Svr_Info")="Dev" THEN
			    sqlStr = sqlStr&"_TEST"
		    END IF 
    		retParamInfo = fnExecSPOutput(sqlStr,paramInfo) 
		    RetErr       = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드 or IDX
		    retErrStr    = GetValue(retParamInfo, "@retErrStr")    ' 에러내역
     
		    IF (RetErr<1) then
		        rw "RetErr="&RetErr
		        rw "retErrStr="&retErrStr
		    ELSE
		    	'4.이전달 해당 master코드에 해당하는 detail 내역 가져오기
	    	  IF (application("Svr_Info")="Dev") THEN
	            sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitDivDetailList_TEST"
	        ELSE
	            sqlStr ="db_SCM_LINK.dbo.sp_VW_BIZ_MonthProfitDivDetailList"
	        END IF 
        	sqlStr =sqlStr+"("&predivMastKey&")"	  
       	  dbiTms_rsget.Open sqlStr, dbiTms_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
				 
				  IF Not (dbiTms_rsget.EOF OR dbiTms_rsget.BOF) THEN 
				  	arrDtl = dbiTms_rsget.getRows()
				  	
		        for j=0 to Ubound(arrDtl,2)
		            idtl_bizSecCd = Trim(arrDtl(1,j))
		            idtl_dlvPro   = Trim(arrDtl(2,j))
		            
		             
		            IF (idtl_bizSecCd<>"") and (idtl_dlvPro<>"") then
		               '5.detail 에 새로 생성된 master 코드로 이전달 detail 내역 복사등록 
		                paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
		                        ,Array("@divMastKey"	,adInteger, adParamInput,, RetErr) _
		                        ,Array("@BIZSECTION_CD"	,adVarchar, adParamInput,10, idtl_bizSecCd) _
		                        ,Array("@divPro" ,adVarchar, adParamInput,10,idtl_dlvPro) _
		                        ,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
		                )
		                
		                sqlStr = "db_SCM_LINK.dbo.sp_SCM2ERP_bizProfit_regDtl" 
		                	
		            	IF application("Svr_Info")="Dev" THEN
		            	    sqlStr = sqlStr&"_TEST"
		              END IF
		        
		                retParamInfo = fnExecSPOutput(sqlStr,paramInfo)
		       
		                RetVal       = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드 or IDX
		                retErrStr    = GetValue(retParamInfo, "@retErrStr")    ' 에러내역
		                
		                rw retErrStr
		            end if
		         next
		        ENd IF
		        dbiTms_rsget.close
		    ENd IF
			Next
		END IF
	  
ELSEIF (mode="delDivMast") then
    '' 기 안분 적용된 내역이 있으면 삭제 불가함.
    paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
            ,Array("@divMastKey"	,adInteger, adParamInput,, divMastKey) _
            ,Array("@isForceDel"	,adInteger, adParamInput,, isForceDel) _
            ,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
    )
    
    sqlStr = "db_SCM_LINK.dbo.sp_SCM2ERP_bizProfit_delMst" 
            	
	IF application("Svr_Info")="Dev" THEN
	    sqlStr = sqlStr&"_TEST"
    END IF
            
    retParamInfo = fnExecSPOutput(sqlStr,paramInfo)
   
    RetErr       = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드 or IDX
    retErrStr    = GetValue(retParamInfo, "@retErrStr")    ' 에러내역
    
    IF (RetErr<1) then
        rw "RetErr="&RetErr
        rw "retErrStr="&retErrStr
    ELSE
        
    ENd IF
ELSEIF (mode="assignDiv") then
    'rw chk
    'rw SLTRKEY
    'rw SLTRKEY_SEQ
    
    chk = split(chk,",")
    SLTRKEY = split(SLTRKEY,",")
    SLTRKEY_SEQ = split(SLTRKEY_SEQ,",")
    
    ''rw "Ubound(chk)="&Ubound(chk)
    for i=0 to Ubound(chk)
        ichk = Trim(chk(i))
        if (ichk<>"") then
            iSLTRKEY = Trim(SLTRKEY(ichk))
            iSLTRKEY_SEQ = Trim(SLTRKEY_SEQ(ichk))
            
            ''rw iSLTRKEY
            ''rw iSLTRKEY_SEQ
            IF (iSLTRKEY<>"" and iSLTRKEY_SEQ<>"") then
                paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
                        ,Array("@divMastKey"	,adInteger, adParamInput,, divMastKey) _
                        ,Array("@SLTRKEY"	,advarchar, adParamInput,12, iSLTRKEY) _
                        ,Array("@SLTRKEY_SEQ"	,adInteger, adParamInput,, iSLTRKEY_SEQ) _
                )
            
                sqlStr = "db_SCM_LINK.dbo.sp_SCM2ERP_DivAssign" 
                        	
            	IF application("Svr_Info")="Dev" THEN
            	    sqlStr = sqlStr&"_TEST"
                END IF
                        
                retParamInfo = fnExecSPOutput(sqlStr,paramInfo)
                RetErr       = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드 or IDX
                
                IF (RetErr<1) then
                    rw "RetErr="&RetErr
                    rw "retErrStr="&retErrStr
                ELSE
                    rw "RetErr="&RetErr
                    rw "retErrStr="&retErrStr
                ENd IF
            end if
        end if
    next
ELSEIF (mode="DelAssignDiv") then    
    IF (application("Svr_Info")="Dev") THEN
        sqlStr ="db_SCM_LINK.dbo.sp_SCM2ERP_bizProfit_delAssign_TEST"
    ELSE
        sqlStr ="db_SCM_LINK.dbo.sp_SCM2ERP_bizProfit_delAssign"
    END IF
    sqlStr =sqlStr+" '"&SLTRKEY&"',"&SLTRKEY_SEQ&" "
rw sqlStr
    dbiTms_dbget.Execute sqlStr
ELSE
    response.write "mode=["&mode&"] 미지정"
END IF


%>
<script type="text/javascript">
alert('ok');
<% if (mode="delDivMast") then divMastKey="" %>
location.href="/admin/analysis/popBizDivSet.asp?yyyymm=<%= AssignYYYYMM %>&divMastKey=<%=divMastKey%>"
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbiTmsClose.asp" -->