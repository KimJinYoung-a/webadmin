<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 결제요청서 ERP연동
' History : 2011.12.16 eastone  생성 erpLink_Process.asp
'###########################################################
%>

<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbiTmsOpen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->  
<!-- #include virtual="/lib/db/dbiTMSHelper.asp"-->  
<!-- #include virtual="/lib/classes/approval/payreqListCls.asp"--> 
<!-- #include virtual="/lib/classes/approval/payRequestCls.asp"-->  

<%
DIM isTESTMODE : isTESTMODE = FALSE
Dim LTp : LTp = reQuestCheckvar(request("LTp"),10)  ''타입
Dim chk : chk = request("chk")                      ''Idx Array
Dim ipFileNo : ipFileNo = request("ipFileNo")
rw LTp
rw chk

chk = split(chk,",")

dim i,jj
Dim payIdx, paramInfo, sqlStr
Dim retParamInfo, RetErr, retErrStr, retErpLinkType, ret_SLTRKEY
Dim errALL
Dim retARR,iTR_MAST_SEQ,iSETTLE_REAL_DATE, iSETTLE_YYYYMM
Dim AssignedRow, tmpAssign, erpLinkType
Dim ierpLinkSeq, iExt5

IF (LTp="A") THEN
    For i=UBound(chk) to LBound(chk) STEP -1
        IF Trim(chk(i))<>"" THEN
            if Trim(chk(i))<70 and Trim(chk(i))<>67 THen
                response.write "과거내역 업로드 불가"
                response.end
            end if
        END IF
    Next
ENd IF


IF (LTp="A") THEN
  For i=UBound(chk) to LBound(chk) STEP -1   ''역으로 등록.
    payIdx = Trim(chk(i))
    IF (payIdx<>"") Then
        ''rw payIdx
        ''검증은 ERP 프로시져에서.
        paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
            ,Array("@payRequestIdx"	,adInteger, adParamInput,, payIdx) _
            ,Array("@erpLinkType"	,adVarchar, adParamOutput,1, "") _
            ,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
    	)
    	
    	''sqlStr = "db_SCM_LINK.dbo.usp_sERP_SCM2ERP_payREQ" ''2016/03/14 아님..
    	sqlStr = "db_SCM_LINK.dbo.sp_SCM2ERP_payREQ_sERP" ''2016/03/17
    	
    	IF application("Svr_Info")="Dev" THEN
    	    sqlStr = sqlStr&"_TEST"
        END IF
        
    	retParamInfo = fnExecSPOutput(sqlStr,paramInfo)
       
        RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드 or IDX
        retErrStr  = GetValue(retParamInfo, "@retErrStr")   ' 생성된 송장번호 
        retErpLinkType = GetValue(retParamInfo, "@erpLinkType")   'S:영업,F:자금수지,C:회계
        
        rw "RetErr="&RetErr
        rw "retErrStr="&retErrStr
        
  if (NOT isTESTMODE) then     
        IF (RetErr<0) THEN
            errALL = errALL&"["&payIdx&"] (ERR:"&RetErr&") "&retErrStr&VbCRLF
        ELSE
            sqlStr = "Update db_Partner.dbo.tbl_eAppPayRequest"
            sqlStr = sqlStr & " SET erpLinkKey='"&RetErr&"'"
            sqlStr = sqlStr & " ,erpLinkType='"&retErpLinkType&"'"
            sqlStr = sqlStr & " ,payRequestState=8"
            sqlStr = sqlStr & " WHERE payREquestIdx="&payIdx
            
            dbget.Execute sqlStr
            
            IF (retErpLinkType="S") then  ''문서 업데이트
                sqlStr = " update D"
                sqlStr = sqlStr & " set erpDocLinkType='"&retErpLinkType&"'"
                sqlStr = sqlStr & " ,erpDocLinkKey="&RetErr
                sqlStr = sqlStr & " ,erpDocSendDate=getdate()"
                sqlStr = sqlStr & " from db_partner.dbo.tbl_eAppPayDoc D"
                sqlStr = sqlStr & " WHERE payREquestIdx="&payIdx
                
                dbget.Execute sqlStr
            END IF
                
        END IF
  end if
    END IF
    
    
  Next
ELSEIF (LTp="AF") THEN          '''입금확정 File - 업체 정기결제.
    IF (ipFileNo<>"") then
        ''검증은 ERP 프로시져에서.
        paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
            ,Array("@ipFileNo"	,adInteger, adParamInput,, ipFileNo) _
            ,Array("@erpLinkType"	,adVarchar, adParamOutput,1, "") _
            ,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
    	)
    	
    	sqlStr = "db_SCM_LINK.dbo.sp_SCM2ERP_ICheFile_sERP" 
    	
    	'' 2018/02/28
    	'' 은행이 새로 생길경우 --  [sERP_ENT_10BY10].dbo.BA_COM_CD where GRP_CD='S003' 의 CD_DESC 값과 우리 은행명이 일치해야함, 
    	'' sp_SCM2ERP_ICheFile_sERP 의 CASE WHEN T.ipkumbank='조흥' THEN '신한' 부분을 추가  할것.
    	
    	IF application("Svr_Info")="Dev" THEN
    	    sqlStr = sqlStr&"_TEST"
        END IF
        
    	retParamInfo = fnExecSPOutput(sqlStr,paramInfo)
       
        RetErr    = GetValue(retParamInfo, "@RETURN_VALUE")         ' 에러코드 or IDX
        retErrStr  = GetValue(retParamInfo, "@retErrStr")           '   
        retErpLinkType = GetValue(retParamInfo, "@erpLinkType")     ' S:영업,F:자금수지,C:회계
        
        rw "RetErr="&RetErr
        rw "retErrStr="&retErrStr
    if (NOT isTESTMODE) then    
        IF (RetErr<0) THEN
            errALL = errALL&"["&ipFileNo&"] (ERR:"&RetErr&") "&retErrStr&VbCRLF
        ELSE
            sqlStr = "Update db_jungsan.dbo.tbl_jungsan_ipkumFile_Master"
            sqlStr = sqlStr & " SET ipFileState=3"
            sqlStr = sqlStr & " ,ERP_TrMastSeq="&RetErr                                '''자금 수지 Key
            sqlStr = sqlStr & " WHERE ipFileNo="&ipFileNo
            
            dbget.Execute sqlStr
                
        END IF
    end if	    
        ''이체등록내역 update
   if (FALSE) then  ''필요 없을듯 함..     
        IF application("Svr_Info")="Dev" THEN
            sqlStr ="db_SCM_LINK.dbo.sp_ERP_RESULT_BY_LINKTYPE_2015_TEST (4,"&RetErr&") "
        ELSE
            sqlStr ="db_SCM_LINK.dbo.sp_ERP_RESULT_BY_LINKTYPE_2015 (4,"&RetErr&") "
        END IF
        
    	dbiTms_rsget.Open sqlStr, dbiTms_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    	IF Not (dbiTms_rsget.EOF OR dbiTms_rsget.BOF) THEN
    		retARR = dbiTms_rsget.getRows()   
    	END IF
    	dbiTms_rsget.close
    	
    	AssignedRow = 0
    	''  d.plan_seq,d.plan_list_seq,d.plan_amt,d.SCM_DTLIDX 
    if (NOT isTESTMODE) then   
    
    	IF (IsARRAY(retARR)) THEN
    	    For jj = 0 To UBound(retARR,2)
    	        ierpLinkSeq         = Trim(retARR(1,jj)) ''CStr(retARR(0,jj))+CStr(Format00(5,Trim(retARR(1,jj))))
    	        iExt5               = Trim(retARR(3,jj))
    	        
    	        if (CStr(iExt5)<>"") THEN
        	        sqlStr = " Update db_jungsan.dbo.tbl_jungsan_ipkumFile_Detail"
        	        sqlStr = sqlStr & " SET erpDTLSeq='"&ierpLinkSeq&"'"
        	        sqlStr = sqlStr & " ,ipFileDetailState=1"                                    ''전송
        	        sqlStr = sqlStr & " WHERE ipFileNo="&ipFileNo
        	        sqlStr = sqlStr & " and isNULL(refipFileDetailIDx,ipFileDetailIDx)="&iExt5
        
                    dbget.Execute sqlStr,tmpAssign
                    
                    AssignedRow = AssignedRow + tmpAssign
                end if
    	    Next
    	END IF
     end if
    end if	
    	rw AssignedRow&"건 반영됨"
    	
    end if
ELSEIF (LTp="R") THEN
    sqlStr = "db_SCM_LINK.dbo.sp_ERP_RESULT_BY_LINKTYPE 1"
    
    IF application("Svr_Info")="Dev" THEN
        sqlStr ="db_SCM_LINK.dbo.sp_ERP_RESULT_BY_LINKTYPE_TEST(1)"
    ELSE
        sqlStr ="db_SCM_LINK.dbo.sp_ERP_RESULT_BY_LINKTYPE_sERP(1)"
    END IF
    
	dbiTms_rsget.Open sqlStr, dbiTms_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
	IF Not (dbiTms_rsget.EOF OR dbiTms_rsget.BOF) THEN
		retARR = dbiTms_rsget.getRows()   
	END IF
	dbiTms_rsget.close
	
	AssignedRow = 0
	''  P.TR_MAST_SEQ, P.SETTLE_REAL_DATE, P.PLAN_SEQ, erpLinkType
	IF (IsARRAY(retARR)) THEN
	    For i = 0 To UBound(retARR,2)
	        iTR_MAST_SEQ        = retARR(0,i)
	        iSETTLE_REAL_DATE   = Trim(retARR(1,i))
	        erpLinkType         = Trim(retARR(3,i))
	        IF (LEN(iSETTLE_REAL_DATE)=8) THEN
	            iSETTLE_REAL_DATE = LEFT(iSETTLE_REAL_DATE,4)+"-"+MID(iSETTLE_REAL_DATE,5,2)+"-"+RIGHT(iSETTLE_REAL_DATE,2)
	        ENd IF
	        iSETTLE_YYYYMM      = LEFT(iSETTLE_REAL_DATE,7)         ''결제일로 할지.. 계산서 날짜로 할지.. ==>계산서 날짜로.
	        
	        if (CStr(iTR_MAST_SEQ)<>"") THEN
    	        sqlStr = " Update db_partner.dbo.tbl_eAppPayRequest"
    	        sqlStr = sqlStr & " SET payRequestState=9"
    	        sqlStr = sqlStr & " ,payRealDate='"&iSETTLE_REAL_DATE&"'"
    	        '''sqlStr = sqlStr & " ,yyyymm='"&iSETTLE_YYYYMM&"'"
    	        sqlStr = sqlStr & " WHERE payRequestType in (1,2)"
    	        sqlStr = sqlStr & " and erpLinkType='"&erpLinkType&"'"
                sqlStr = sqlStr & " and payRequestState=8"
                sqlStr = sqlStr & " and erpLinkKey='"&iTR_MAST_SEQ&"'"
    
                dbget.Execute sqlStr,tmpAssign
                
                AssignedRow = AssignedRow + tmpAssign
            end if
	    Next
	END IF
	
	rw AssignedRow&"건 반영됨"
ELSEIF (LTp="C") THEN                                           ''''운영비 전송.
  For i=UBound(chk) to LBound(chk) STEP -1   ''역으로 등록.
    payIdx = Trim(chk(i))
    IF (payIdx<>"") Then
        ''rw payIdx
        ''검증은 ERP 프로시져에서.  
        paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
            ,Array("@OpExpidx"	,adInteger, adParamInput,, payIdx) _
            ,Array("@RET_SLTRKEY"	,adVarchar, adParamOutput,12, "") _
            ,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
    	)
    	
    	sqlStr = "db_SCM_LINK.dbo.sp_SCM2ERP_SLMAST_INPUT_sERP" 
    	
    	IF application("Svr_Info")="Dev" THEN
    	    sqlStr = sqlStr&"_TEST"
        END IF
        
    	retParamInfo = fnExecSPOutput(sqlStr,paramInfo)
       
        RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드 or IDX
        ret_SLTRKEY = GetValue(retParamInfo, "@RET_SLTRKEY")   'ret_SLTRKEY
        retErrStr  = GetValue(retParamInfo, "@retErrStr")   ' 생성된 송장번호 
        
        
        ''rw "RetErr="&RetErr
        ''rw "retErrStr="&retErrStr
        
        IF  ((isTESTMODE) or (RetErr<0)) THEN
            errALL = errALL&"["&payIdx&"] (ERR:"&RetErr&") "&retErrStr&"[ret_SLTRKEY="&ret_SLTRKEY&"]"&VbCRLF
        ELSE
            sqlStr = "Update db_partner.dbo.tbl_OpExpMonthly"
            sqlStr = sqlStr & " SET erpLinkKey='"&ret_SLTRKEY&"'"
            sqlStr = sqlStr & " ,erpLinkType='"&LTp&"'"
            sqlStr = sqlStr & " ,state=10"
            sqlStr = sqlStr & " WHERE opexpidx="&payIdx
            
            dbget.Execute sqlStr
            
            '''전표 상세 업데이트
            ''sqlStr = "db_SCM_LINK.dbo.sp_ERP_RESULT_BY_LINKTYPE 3,idx"
    
            sqlStr = " update D"&VBCRLF
            sqlStr = sqlStr&" SET erpLinkSeq=M.erpLinkKey"&VBCRLF
            sqlStr = sqlStr&" from db_partner.dbo.tbl_OpExpMonthly M"&VBCRLF
            sqlStr = sqlStr&" 	Join db_partner.dbo.tbl_OpExpDaily D"&VBCRLF
            sqlStr = sqlStr&" 	on M.opexpidx="&payIdx&""&VBCRLF
            sqlStr = sqlStr&" 	and M.yyyymm=LEFT(D.yyyymmdd,7)"&VBCRLF
            sqlStr = sqlStr&" 	and M.OpExpPartidx=D.OpExpPartIdx"&VBCRLF
            
            dbget.Execute sqlStr,AssignedRow
	
'            IF application("Svr_Info")="Dev" THEN
'                sqlStr ="db_SCM_LINK.dbo.sp_ERP_RESULT_BY_LINKTYPE_TEST (3,"&RetErr&") "
'            ELSE
'                sqlStr ="db_SCM_LINK.dbo.sp_ERP_RESULT_BY_LINKTYPE_sERP (3,'"&ret_SLTRKEY&"') "
'            END IF
'            
'        	dbiTms_rsget.Open sqlStr, dbiTms_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
'        	IF Not (dbiTms_rsget.EOF OR dbiTms_rsget.BOF) THEN
'        		retARR = dbiTms_rsget.getRows()   
'        	END IF
'        	dbiTms_rsget.close
'        	
'        	AssignedRow = 0
'        	''  P.TR_MAST_SEQ, P.SETTLE_REAL_DATE, P.PLAN_SEQ, erpLinkType
'        	IF (IsARRAY(retARR)) THEN
'        	    For jj = 0 To UBound(retARR,2)
'        	        iTR_MAST_SEQ        = retARR(0,jj)
'        	        ierpLinkSeq         = Trim(retARR(1,jj))
'        	        ierpLinkSeq         = ierpLinkSeq+Format00(4,Trim(retARR(2,jj)))
'        	        iExt5               = Trim(retARR(3,jj))
'        	        
'        	        if isNULL(iExt5) then ''2016/05/03 sERP
''        	            sqlStr = " Update db_partner.dbo.tbl_OpExpDaily"
''            	        sqlStr = sqlStr & " SET erpLinkSeq='"&ret_SLTRKEY&"'"  
''            	        sqlStr = sqlStr & " WHERE opExpDailyIdx="&iExt5&""
''            
''                        dbget.Execute sqlStr,tmpAssign
''                        
''                        AssignedRow = AssignedRow + tmpAssign
'        	        else
'            	        if (CStr(iExt5)<>"") THEN
'                	        sqlStr = " Update db_partner.dbo.tbl_OpExpDaily"
'                	        sqlStr = sqlStr & " SET erpLinkSeq='"&ierpLinkSeq&"'"
'                	        sqlStr = sqlStr & " WHERE opExpDailyIdx="&iExt5&""
'                
'                            dbget.Execute sqlStr,tmpAssign
'                            
'                            AssignedRow = AssignedRow + tmpAssign
'                        end if
'                    end if
'        	    Next
'        	END IF
        	
        	rw AssignedRow&"건 반영됨"
        END IF
    END IF
    
    
  Next
ELSEIF (LTp="D") THEN                                           ''''체크카드내역 전송.
  For i=UBound(chk) to LBound(chk) STEP -1   ''역으로 등록.
    payIdx = Trim(chk(i))
    IF (payIdx<>"") Then
        ''rw payIdx
        ''검증은 ERP 프로시져에서.  
        paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
            ,Array("@OpExpidx"	,adInteger, adParamInput,, payIdx) _
            ,Array("@RET_SLTRKEY"	,adVarchar, adParamOutput,12, "") _
            ,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
    	)
    	
    	sqlStr = "db_SCM_LINK.dbo.sp_SCM2ERP_SLMAST_INPUT_CARD_sERP" 
    	
    	IF application("Svr_Info")="Dev" THEN
    	    sqlStr = sqlStr&"_TEST"
        END IF
        
    	retParamInfo = fnExecSPOutput(sqlStr,paramInfo)
       
        RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' 에러코드 or IDX
        retErrStr  = GetValue(retParamInfo, "@retErrStr")   ' 
        ret_SLTRKEY = GetValue(retParamInfo, "@RET_SLTRKEY")   'ret_SLTRKEY
        ''retErpLinkType = GetValue(retParamInfo, "@erpLinkType")   'S:영업,F:자금수지,C:회계
        
        ''rw "RetErr="&RetErr
        ''rw "retErrStr="&retErrStr
        
        IF  ((isTESTMODE) or (RetErr<0)) THEN
            errALL = errALL&"["&payIdx&"] (ERR:"&RetErr&") "&retErrStr&":ret_SLTRKEY:"&ret_SLTRKEY&VbCRLF
        ELSE
            sqlStr = "Update db_partner.dbo.tbl_OpExpMonthlyCard"
            sqlStr = sqlStr & " SET erpLinkKey='"&ret_SLTRKEY&"'"
            sqlStr = sqlStr & " ,erpLinkType='"&LTp&"'"
            sqlStr = sqlStr & " ,state=10"
            sqlStr = sqlStr & " WHERE oPExpCardIdx="&payIdx
            
            dbget.Execute sqlStr
            
            ''[ToDo]
            '''전표 상세 업데이트
            ''sqlStr = "db_SCM_LINK.dbo.sp_ERP_RESULT_BY_LINKTYPE 3,idx"
    
            IF application("Svr_Info")="Dev" THEN
                sqlStr ="db_SCM_LINK.dbo.sp_ERP_RESULT_BY_LINKTYPE_TEST (3,"&RetErr&") "
            ELSE
                sqlStr ="db_SCM_LINK.dbo.sp_ERP_RESULT_BY_LINKTYPE_sERP (3,'"&ret_SLTRKEY&"') "
            END IF
            
        	dbiTms_rsget.Open sqlStr, dbiTms_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
        	IF Not (dbiTms_rsget.EOF OR dbiTms_rsget.BOF) THEN
        		retARR = dbiTms_rsget.getRows()   
        	END IF
        	dbiTms_rsget.close
        	
        	AssignedRow = 0
        	''  P.TR_MAST_SEQ, P.SETTLE_REAL_DATE, P.PLAN_SEQ, erpLinkType
        	IF (IsARRAY(retARR)) THEN
        	    For jj = 0 To UBound(retARR,2)
        	        iTR_MAST_SEQ        = retARR(0,jj)
        	        ierpLinkSeq         = Trim(retARR(1,jj))
        	        ierpLinkSeq         = ierpLinkSeq+Format00(4,Trim(retARR(2,jj)))
        	        iExt5               = Trim(retARR(3,jj))
        	        
        	        if (CStr(iExt5)<>"") THEN
            	        sqlStr = " Update db_partner.dbo.tbl_OpExpDailyCard"
            	        sqlStr = sqlStr & " SET erpLinkSeq='"&ierpLinkSeq&"'"
            	        sqlStr = sqlStr & " WHERE opExpDailyCardIdx="&iExt5&""
            
                        dbget.Execute sqlStr,tmpAssign
                        
                        AssignedRow = AssignedRow + tmpAssign
                    end if
        	    Next
        	END IF
        	
        	rw AssignedRow&"건 반영됨"
        END IF
    END IF
    
    
  Next
ELSE
    rw "지정안됨 LTp="&LTp
END IF
rw errALL
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbiTmsClose.asp" -->