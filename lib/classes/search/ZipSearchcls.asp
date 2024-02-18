<%
'###########################################################
' Description :  SCM ������ȣ Ŭ����
' History : 2016.07.01 �ѿ�� ����Ʈ ���� ����
'###########################################################
%>
<%
''CS_EUCKR <=> CS_UTF8
''group by �ʵ�� ���� ���� ����.
'''--------------------------------------------------------------------------------------

DIM G_ORGSCH_ADDR , GG_ORGSCH_ADDR
DIM G_1STSCH_ADDR , GG_1STSCH_ADDR
DIM G_2NDSCH_ADDR , GG_2NDSCH_ADDR
DIM G_3RDSCH_ADDR , GG_3RDSCH_ADDR
Dim G_4THSCH_ADDR , GG_4THSCH_ADDR

DIM G_SCH_TIME : G_SCH_TIME=formatdatetime(now(),4)

IF (application("Svr_Info") = "Dev") THEN
     G_1STSCH_ADDR = "192.168.50.10"  ''"110.93.128.109" ''
     G_2NDSCH_ADDR = "192.168.50.10"
     G_3RDSCH_ADDR = "192.168.50.10"
     G_4THSCH_ADDR = "192.168.50.10"
     G_ORGSCH_ADDR = "192.168.50.10"
ELSE
     G_1STSCH_ADDR = "192.168.0.210"        ''192.168.0.210  :: (X)
     G_2NDSCH_ADDR = "192.168.0.207"        ''192.168.0.207  :: 
     G_3RDSCH_ADDR = "192.168.0.209"        ''192.168.0.209  :: (X)
     G_4THSCH_ADDR = "192.168.0.208"        ''192.168.0.208  :: 
     G_ORGSCH_ADDR = "192.168.0.206"        ''192.168.0.206
END IF

GG_1STSCH_ADDR = G_1STSCH_ADDR
GG_2NDSCH_ADDR = G_2NDSCH_ADDR
GG_3RDSCH_ADDR = G_3RDSCH_ADDR
GG_4THSCH_ADDR = G_4THSCH_ADDR
GG_ORGSCH_ADDR = G_ORGSCH_ADDR

if (Application("G_ZIPSCH_ADDR")="") or (Application("G_ZIPSCH_ADDR")=GG_1STSCH_ADDR) or (Application("G_ZIPSCH_ADDR")=GG_3RDSCH_ADDR) then
	Application("G_ZIPSCH_ADDR")=G_ORGSCH_ADDR
end if
G_ORGSCH_ADDR = Application("G_ZIPSCH_ADDR")

''sample in doc
function escapeQuery( istr )
	dim ret, c, i
	ret = ""
	For i=1 To Len(istr)
		c = Mid(istr,i,1)
		select case c
		case "\"
			ret = ret & "\\"
		case "'"
			ret = ret & "\'"
		case chr(34)
			ret = ret & "\" & chr(34)
		case "*"
			ret = ret & "\*"
		case else
			ret = ret & c
		end select
	Next
	escapeQuery = ret
end function

function getTimeChkAddr(defaultAddr)
    '''6��10�� 1���� �ε��� �� 2�������� Copy
    '''6��50��~ 1��=>3�������� Copy
    getTimeChkAddr = defaultAddr

    IF (defaultAddr=G_4THSCH_ADDR) THEN
        IF (G_SCH_TIME>"06:00:00") and (G_SCH_TIME<"06:40:00") then
            getTimeChkAddr = G_2NDSCH_ADDR
        END IF
    ELSE
        IF (G_SCH_TIME>"06:40:00") and (G_SCH_TIME<"07:00:00") then
            getTimeChkAddr = G_4THSCH_ADDR
        END IF
    END IF
end function

function debugQuery(iDocruzer,Scn,iSearchQuery,iSortQuery,iFTotalCount,iFResultcount)
  exit function
    IF Not (application("Svr_Info")="Dev") THEN
        exit function
    ENd IF

    dim itime
    Call iDocruzer.GetResult_SearchTime(itime) '�ҿ�ð�
    rw "-------------------------------"
    rw Scn
    rw iSearchQuery
    rw iSortQuery
    rw "FTotalCount:"&iFTotalCount
    rw "FResultcount:"&iFResultcount
    rw "GetResult_SearchTime:"&itime
end function
'''--------------------------------------------------------------------------------------

Class CZipPrdItem
    Private SUB Class_initialize()

	End SUB

	Private SUB Class_Terminate()

	End SUB
	
	PUBLIC Fidx
    PUBLIC Fzipcode
    PUBLIC Fsido
    PUBLIC Fgungu
    PUBLIC Feupmyun
    PUBLIC Froad
    PUBLIC Fisunder
    PUBLIC Fbuilding_no
    PUBLIC Fbuilding_sub
    PUBLIC Fofficial_bld
    PUBLIC Fdong
    PUBLIC Fri
    PUBLIC Fdong_admin
    PUBLIC Fissan
    PUBLIC Fjibun_main
    PUBLIC Fjibun_sub
    PUBLIC Fold_zipcode
    PUBLIC Fzipgroup   '2016/07/14 left(zipcode,2)
    
End Class
	
Class CGroupbySidoItem

	Private SUB Class_initialize()

	End SUB

	Private SUB Class_Terminate()

	End SUB

	PUBLIC Fzipgroup
	PUBLIC Fsido
	PUBLIC Fgungu
	PUBLIC FCNT
	
End Class



Class SearchItemCls

	Private SUB Class_initialize()
        ''�⺻ 1�� ����.------------------------
		SvrAddr = getTimeChkAddr(G_ORGSCH_ADDR)
		''--------------------------------------

		SvrPort = "6167"'DocSvrPort

		AuthCode = "" '������

		Logs = "" '�αװ�

		FResultCount = 0
		FTotalCount = 0
		FPageSize = 10
		FCurrPage = 1
		FPageSize = 30
		FRectColsSize =5
		FLogsAccept = false

	End SUB

	Private SUB Class_Terminate()

	End SUB

	dim FItemList
	dim FPageSize
	dim FCurrPage
	dim FScrollCount
	dim FResultCount
	dim FTotalCount
	dim FTotalPage

	dim FRectSearchTxt		'�˻���
	dim FRectZipgroup               '' �׷����ڵ�
	
	dim FRectSearchSiDo 	'SiDo
	dim FRectSearchGunGu	'gungu
	dim FRectSearchRoad     'Road
	dim FRectSearchBuilding_no  'Building_no
	dim FRectSearchDong         'Dong
	dim FRectSearchJibun_main   'Jibun_main
	dim FRectSearchJibun_sub    'Jibun_sub
	
	dim FRectPrevSearchTxt	'���� �˻���
	
	dim FRectExceptText		'���ܾ�
	
	dim FRectSortMethod		'���Ĺ�� 

	dim FCheckResearch 		'����� ��˻� üũ��
	dim FRectColsSize		'��� ����Ʈ ����
	dim FLogsAccept			'�߰� �α� ���� ����


	Private SvrAddr
	Private SvrPort
	Private AuthCode
	Private Logs
	Private Scn
	private strQuery
	Private Order
	Private StartNum

	Private SearchQuery
	Private SortQuery
    
    public function InitDocruzer(iDocruzer)
        InitDocruzer = FALSE
        IF ( iDocruzer.BeginSession() < 0 ) THEN
			EXIT function
		End If
        
        IF NOT DocSetOption(iDocruzer) THEN
			EXIT function
		End If
		InitDocruzer = TRUE
    End function

    public function DocSetOption(iDocruzer)
        dim ret 
        ret = iDocruzer.SetOption(iDocruzer.OPTION_REQUEST_CHARSET_UTF8,1)
        DocSetOption = (ret>=0)
    end function
    

	''/�˻� ���� ����
	FUNCTION getSearchQuery(byref query)
		dim strQue, arrCCD, arrSCD, arrACD, lp

		'### �˻����� ���� ###

		'@ �˻���(Ű����)
		IF FRectSearchTxt<>"" Then
			FRectSearchTxt = chgCoinedKeyword(FRectSearchTxt)
			FRectSearchTxt = escapeQuery(FRectSearchTxt)  ''2015 �߰�
		''FRectSearchTxt=replace(FRectSearchTxt," "," & ")
		''rw FRectSearchTxt
			IF FRectExceptText<>"" Then
			    FRectExceptText = escapeQuery(FRectExceptText)  ''2015 �߰�
				strQue = getQrCon(strQue) & "(idx_searchtxt='" & FRectSearchTxt & " ! " & FRectExceptText & "' BOOLEAN) "	'���ܾ�
			else
				strQue = getQrCon(strQue) & "idx_searchtxt='" & FRectSearchTxt & "'  allword synonym"	'Ű����˻�(���Ǿ� ����) synonym
				'strQue = getQrCon(strQue) & "idx_searchtxt='" & FRectSearchTxt & "'  BOOLEAN "
				'strQue = getQrCon(strQue) & "idx_searchtxt like '" & FRectSearchTxt & "*'   "
			End if
		END IF

        '@idx_zipgroup
        IF FRectZipgroup<>"" Then
			strQue = strQue & getQrCon(strQue) & "idx_zipgroup='"&FRectZipgroup&"' "
		End IF
		
		'@ SiDo
		IF FRectSearchSiDo<>"" Then
			strQue = strQue & getQrCon(strQue) & "idx_sido='"&FRectSearchSiDo&"' "
		End IF
		
		'@ GunGu
		IF FRectSearchGunGu<>"" Then
			strQue = strQue & getQrCon(strQue) & "idx_gungu='"&FRectSearchGunGu&"' "
		End IF
		
        '@ Road
		IF FRectSearchRoad<>"" Then
			strQue = strQue & getQrCon(strQue) & "idx_road='"&FRectSearchRoad&"' "
		End IF
		
		'@ Building_no
		IF FRectSearchBuilding_no<>"" Then
			strQue = strQue & getQrCon(strQue) & "idx_building_no='"&FRectSearchBuilding_no&"' "
		End IF

        '@ Dong
		IF FRectSearchDong<>"" Then
			strQue = strQue & getQrCon(strQue) & "dong='"&FRectSearchDong&"' "
		End IF
		
		'@ Jibun_main
		IF FRectSearchJibun_main<>"" Then
			strQue = strQue & getQrCon(strQue) & "jibun_main='"&FRectSearchJibun_main&"' "
		End IF
		
		'@ Jibun_sub
		IF FRectSearchJibun_sub<>"" Then
			strQue = strQue & getQrCon(strQue) & "jibun_sub='"&FRectSearchJibun_sub&"' "
		End IF
		
        
		query = strQue
	End FUNCTION

	Sub getSortQuery(byref query)
		dim strQue

        strQue = ""
        
		'// ����
		IF FRectSortMethod="zipcode" THEN 'zipcode
			strQue = strQue & " ORDER BY zipcode"
		ELSEIF FRectSortMethod="sido" THEN 'sido
			strQue = strQue & " ORDER BY sido"
		ELSEIF FRectSortMethod="idx" THEN 'idx
			strQue = strQue & " ORDER BY idx"
		ELSE
		    strQue = strQue & " ORDER BY $RELEVANCE DESC"  ''���յ���
		END IF
		
		query = strQue
	End Sub

	Function getQrCon(query)
		if Not(query="" or isNull(query)) then
			getQrCon = " and "
		end if
	End Function

	'####### ������ȣ  �˻� - �˻� ���� ######
	PUBLIC SUB getSearchList()

		DIM Scn
		DIM Docruzer,ret

		DIM Logs ,iRows
		DIM arrData,arrSize, retMatchCd, retMatchVal
		Dim iDocErrMsg

		'// �˻� ��� ��� �ó�������
		Scn= "scn_dt_zipcode"		    

		StartNum = (FCurrPage -1)*FPageSize '// �˻����� Row

		CALL getSearchQuery(SearchQuery)	'// �˻� ��������
		CALL getSortQuery(SortQuery)		'// ���� ���� ����
		''Response.Write SearchQuery &"<Br>"
		IF SearchQuery="" THEN
			EXIT SUB
		END IF
		

		IF (FALSE) and (FLogsAccept) and (FRectSearchTxt<>"") and (FCurrPage="1") THEN
		    
            'Logs = "��ǰ+^" & FRectSearchTxt & "]##" & FRectSearchTxt & "||" & FRectPrevSearchTxt  	'// �αװ�
            ''2015 search4
            '�⺻:[����Ʈ@ī�װ���+�����$�����ڵ�|����|�˻���Ÿ��(����)|ù�˻�|��������ȣ|���ļ�^�����˻���##�˻���] ''�⺻
            Dim iLOG_SITE : iLOG_SITE = "ZIP"
            Dim iLOG_CATE : iLOG_CATE = "RECT" 
            Dim iLOG_USER : iLOG_USER = GetUserLevelStr(GetLoginUserLevel) '' ȸ������� ���
            Dim iLOG_SEX  : iLOG_SEX  = "" '' 0��α���,1����,2����
            Dim iLOG_AGE  : iLOG_AGE  = "" '' 0��α���,1:10��,2:20��,3:30��,4:40��,5:50��
            Dim iLOG_STYPE : iLOG_STYPE = "" '' ���� ������ X
            Dim iLOG_FIRST : iLOG_FIRST = "" '' ù�˻�/��˻� ������ X  FCheckResearch
            
            Logs = iLOG_SITE&"@"                ''[ @
            Logs = Logs&iLOG_CATE&"+"           ''@ +
            Logs = Logs&iLOG_USER&"$"           ''+ $
            Logs = Logs&iLOG_SEX&"|"            ''$ |
            Logs = Logs&iLOG_AGE&"|"            ''| | 
            Logs = Logs&iLOG_STYPE&"|"          ''| | 
            Logs = Logs&iLOG_FIRST&"|"          ''| | 
            Logs = Logs&FCurrPage&"|"           ''| | 
            Logs = Logs&FRectSortMethod&"^"     ''| ^ 
            Logs = Logs&FRectPrevSearchTxt&"##" ''^ ##
            Logs = Logs&FRectSearchTxt          ''## ]
		END IF

       
        ''---------------------------------------------------------------------------------------------------------

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT SUB
		END IF
		
		ret = Docruzer.SubmitQuery(SvrAddr, SvrPort, _
						AuthCode, Logs, Scn, _
						SearchQuery,SortQuery, _
						FRectSearchTxt,StartNum, FPageSize, _
						Docruzer.LC_KOREAN, Docruzer.CS_UTF8)

		IF( ret < 0 ) THEN
		    IF (application("Svr_Info")="Dev") THEN
    		    rw "err1"
    		    rw SearchQuery
    		    rw Docruzer.GetErrorMessage()
    		END IF
    		
			'dbget.execute "EXECUTE db_log.dbo.sp_Ten_DocLog @ErrMsg ='"& html2db(SearchQuery) & "[" & html2db(Docruzer.GetErrorMessage()) &"]'"

			dbget.execute "EXECUTE db_log.dbo.sp_Ten_DocLog @ErrMsg ='["&SvrAddr&"]"& html2db(Docruzer.GetErrorMessage())&"["&Request.ServerVariables("REMOTE_ADDR")&"]["&Request.ServerVariables("LOCAL_ADDR")&"]["& html2db(SearchQuery)&"]'"

			iDocErrMsg = Docruzer.GetErrorMessage()
            if (InStr(iDocErrMsg,"recv queue full")>0) or (InStr(iDocErrMsg,"socket time out")>0) or (InStr(iDocErrMsg,"cannot connect to server")>0) or (InStr(iDocErrMsg,"scn_dt_zipcode' doesn't exist")>0) then
                IF (SvrAddr=Application("G_ZIPSCH_ADDR")) then
                    if (Application("G_ZIPSCH_ADDR")= GG_ORGSCH_ADDR) then
                        Application("G_ZIPSCH_ADDR") = GG_2NDSCH_ADDR
                    elseif (Application("G_ZIPSCH_ADDR")=GG_2NDSCH_ADDR) then
                        Application("G_ZIPSCH_ADDR") = GG_4THSCH_ADDR
					elseif (Application("G_ZIPSCH_ADDR")=GG_4THSCH_ADDR) then
                        Application("G_ZIPSCH_ADDR") = GG_ORGSCH_ADDR
                    end if
                end if
            end if

			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING

'				'// 1�� ���� ������ 2������ ����(2���� ������ Skip)
'				if (SvrAddr = G_1STSCH_ADDR) then
'					SvrAddr = G_2NDSCH_ADDR  ''"192.168.0.108"
'					if (G_1STSCH_ADDR<>G_2NDSCH_ADDR) then  ''�߰� 2013/09
'					    call getSearchList()
'				    end if
'				end if

			EXIT SUB
		END IF

		Call Docruzer.GetResult_TotalCount(FTotalCount) '�˻���� �� ��
		Call Docruzer.GetResult_RowSize(FResultcount) '�˻� ��� ��
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		'Response.write "�˻������ : " & FTotalCount & "<br>"
		IF( FResultCount <= 0 ) THEN
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB 'Response.write "GetResult_RowSize: " & Docruzer.GetErrorMessage()
		END IF

		FTotalPage =  Cdbl(FTotalCount\FPageSize)
		IF  (FTotalCount\FPageSize)<>(FTotalCount/FPageSize) THEN
			FtotalPage = FtotalPage +1
		END IF

		REDIM FItemList(FResultCount)

		FOR iRows=0 to FResultCount -1

			ret = Docruzer.GetResult_Row( arrData, arrSize, iRows )

			IF( ret < 0 ) THEN
				'Response.write "GetResult_Row: " & Docruzer.GetErrorMessage()
				EXIT FOR
			END IF

			SET FItemList(iRows) = NEW CZipPrdItem
                
                FItemList(iRows).Fidx           = arrData(0)
                FItemList(iRows).Fzipcode       = arrData(1)
                FItemList(iRows).Fsido          = arrData(2)
                FItemList(iRows).Fgungu         = arrData(3)
                FItemList(iRows).Feupmyun       = arrData(4)
                FItemList(iRows).Froad          = arrData(5)
                FItemList(iRows).Fisunder       = arrData(6)
                FItemList(iRows).Fbuilding_no   = arrData(7)
                FItemList(iRows).Fbuilding_sub  = arrData(8)
                FItemList(iRows).Fofficial_bld  = arrData(9)
                FItemList(iRows).Fdong          = arrData(10)
                FItemList(iRows).Fri            = arrData(11)
                FItemList(iRows).Fdong_admin    = arrData(12)
                FItemList(iRows).Fissan         = arrData(13)
                FItemList(iRows).Fjibun_main    = arrData(14)
                FItemList(iRows).Fjibun_sub     = arrData(15)
                FItemList(iRows).Fold_zipcode   = arrData(16)
                
			SET arrData = NOTHING
			SET arrSize = NOTHING

		NEXT
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB

	'####### �˻� �׷�  ī����  ######

		

	'####### ��ǰ �˻� �귣�庰 ī����  ######
	PUBLIC SUB getGroupbySido()

		'// �˻� ��� ��� �ó�������
		
		Scn= "scn_dt_zipcodeGroup"
		
		dim Logs
		dim Docruzer,ret
		dim iRows
		dim arrData,arrSize

		dim FseekTime

		StartNum = 0 						'// �˻����� Row
		call getSearchQuery(SearchQuery)	'// �˻� ��������

		'//�׷� ������ ����(���� ���� ����)
		
		SortQuery = " GROUP BY zipgroup order by zipgroup asc " ''desc $RELEVANCE
		
		IF SearchQuery="" Then
			EXIT SUB
		End If

		dim Rowids,Scores
		FTotalCount = 0

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT SUB
		END IF
		
	'response.write "group : " & SearchQuery & SortQuery & "<br>"
		ret = Docruzer.SubmitQuery(SvrAddr, SvrPort, _
						AuthCode, Logs, Scn, _
						SearchQuery,SortQuery, _
						FRectSearchTxt,StartNum, FPageSize, _
						Docruzer.LC_KOREAN, Docruzer.CS_UTF8)

		If( ret < 0 ) Then
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT SUB
		END IF

		Call Docruzer.GetResult_RowSize(FResultcount) '�˻� ��� ��
		Call Docruzer.GetResult_Rowid(Rowids,Scores)
CALL debugQuery(Docruzer,Scn,SearchQuery,SortQuery,FTotalCount,FResultcount)

		REDIM FItemList(FResultCount)

		Call Docruzer.GetResult_TotalCount(FTotalCount) '�˻���� �� ��

		FOR iRows = 0 to FResultcount -1

			ret = Docruzer.GetResult_Row( arrData, arrSize, iRows )

			IF( ret < 0 ) THEN
				'Response.write "GetResult_Row: " & Docruzer.msg
				EXIT FOR
			END IF

			SET FItemList(iRows) = new CGroupbySidoItem
				FItemList(iRows).Fzipgroup	= arrData(0)
				FItemList(iRows).Fsido		= arrData(1)
				FItemList(iRows).Fgungu	    = arrData(2)
				FItemList(iRows).FCNT 	    = Scores(iRows)

			SET arrData = NOTHING
			SET arrSize = NOTHING

		NEXT

		SET Rowids= NOTHING
		SET Scores= NOTHING

		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	End SUB


	PUBLIC FUNCTION HasPreScroll()
		HasPreScroll = StartScrollPage > 1
	END FUNCTION

	PUBLIC FUNCTION HasNextScroll()
		HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
	END FUNCTION

	PUBLIC FUNCTION StartScrollPage()
		StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
	END FUNCTION

End Class



'// ������/���Ǿ� ��ȯ ó�� (����� �� ���Ǿ �ȵǴ� ���� ������ ���)
Function chgCoinedKeyword(kwd)
	dim arrChgTxt, arrItm
	arrChgTxt = split("��8||ban8",",")

	for each arrItm in arrChgTxt
		arrItm = split(arrItm,"||")
		if ubound(arrItm)>0 then
			kwd = Replace(kwd,arrItm(0),arrItm(1))
		end if
	next

	chgCoinedKeyword = kwd
end Function


'// �߰� ī�װ��� ��ȣ ���� (�߰�ī�װ������� �ش� ī�װ��� ��ȣ�� ����)
Function getArrayDispCate(vDisp,vArr)
	Dim vRst, i

	if vArr="" or isNull(vArr) or vDisp="" or isNull(vDisp) then Exit Function

	vArr = replace(trim(vArr)," ",",")
	vRst = split(vArr,",")

	if Not(isArray(vRst)) then Exit Function

	for i=0 to ubound(vRst)
		if inStr(vRst(i),vDisp)>0 then
			getArrayDispCate = vRst(i)
			Exit function
		end if
	next
end Function
%>