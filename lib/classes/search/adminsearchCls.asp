<%
DIM G_KSCOLORCD : G_KSCOLORCD = Array("023","001","002","010","021","003","004","024","019","005","016","006","007","020","008","018","017","009","011","012","022","013","014","015","025","026","027","028","029","030","031")
DIM G_KSCOLORNM : G_KSCOLORNM = Array("����","����","��Ȳ","����","ī��","���","������","���̺���","īŰ","�ʷ�","��Ʈ","���Ķ�","�Ķ�","���̺�","����","������","���̺���ũ","��ũ","���","����ȸ��","£��ȸ��","����","����","�ݻ�","üũ","��Ʈ������","��Ʈ","�ö��","�����","�ִϸ�","������")

Dim G_KSSTYLECD : G_KSSTYLECD = Array("010","020","030","040","050","060","070","080","090")
Dim G_KSSTYLENM : G_KSSTYLENM = Array("Ŭ����","ťƼ","���","���","���߷�","������Ż","��","�θ�ƽ","��Ƽ��")

DIM G_ORGSCH_ADDR
DIM G_1STSCH_ADDR
DIM G_2NDSCH_ADDR
DIM G_3RDSCH_ADDR
Dim G_4THSCH_ADDR

IF (application("Svr_Info") = "Dev") THEN
     G_1STSCH_ADDR = "192.168.50.10"  ''"110.93.128.109" ''
     G_2NDSCH_ADDR = "192.168.50.10"
     G_3RDSCH_ADDR = "192.168.50.10"
     G_4THSCH_ADDR = "192.168.50.10"
     G_ORGSCH_ADDR = "192.168.50.10"
ELSE
     G_1STSCH_ADDR = "192.168.0.210"        ''192.168.0.110  :: �˻�������(search.asp)   '
     G_2NDSCH_ADDR = "192.168.0.207"        ''192.168.0.107  :: ī�װ�, ��ǰ, �귣��
     G_3RDSCH_ADDR = "192.168.0.209"        ''192.168.0.109  :: GiftPlus , scn_dt_itemDispColor :: Ȯ��.
     G_4THSCH_ADDR = "192.168.0.208"        ''192.168.0.108  :: mobile 6:10�п� �ε��� ���� ī��
     G_ORGSCH_ADDR = "192.168.0.206"
END IF

Class SearchItemCls
    Private SvrAddr
	Private SvrPort
	Private AuthCode
	Private Logs
	
	dim FItemList
	dim FPageSize
	dim FCurrPage
	dim FScrollCount
	dim FResultCount
	dim FTotalCount
	dim FTotalPage

	dim FRectSearchTxt
	dim FRectColsSize
	dim FLogsAccept
	
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
    
    Private SUB Class_initialize()
        ''�⺻ 1�� ����.------------------------
		SvrAddr = (G_1STSCH_ADDR)
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
	
	'####### �ǽð��α�˻���:K-search  ######
	PUBLIC FUNCTION getRealtimePopularKeyWords(byRef arDt, byRef arTg,byval iSVR,byval iType, byval idomain)

		DIM Docruzer,ret
		DIM iRows
		DIM arrData,arrSize, arrTags
		DIM MaxCnt : MaxCnt =FPageSize

		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT Function
		END IF
		if (iSVR="") then
            SvrAddr = G_ORGSCH_ADDR  '' 106������ �ϴ�
        else
            SvrAddr = iSVR
        end if
        if (iType="") then iType=1
        if (idomain="") then idomain=0
                
		ret = Docruzer.getRealTimePopularKeyword _
						(SvrAddr&":"&SvrPort,_
						arrSize,arrData,arrTags,_
						MaxCnt,iType,idomain)                     ''' 0 file / 1 memory,  ������
		getRealtimePopularKeyWords = ret
		IF( ret < 0 ) THEN
		    'rw "TTT:"&Docruzer.GetErrorMessage()
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT FUNCTION
		END IF
		arDt = arrData
		arTg = arrTags
		SET arrData = NOTHING
		SET arrTags = NOTHING
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	END FUNCTION
	
    '####### ��õ�˻���  ######
	PUBLIC FUNCTION getRecommendKeyWords(idomainno)

		Dim Docruzer,ret
		Dim iRows
		Dim arrData,arrSize
		Dim MaxCnt : MaxCnt = 1000
        
        if (FPageSize<MaxCnt) then MaxCnt=FPageSize
            
		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT Function
		END IF
		
        SvrAddr = G_ORGSCH_ADDR  '' 106������ �ϴ�

		ret = Docruzer.RecommendKeyWord _
						(SvrAddr&":"&SvrPort,_
						arrSize,arrData,_
						MaxCnt,replace(FRectSearchTxt," ",""),idomainno)

		IF( ret < 0 ) THEN
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT FUNCTION

		END IF

		getRecommendKeyWords = arrData
		SET arrData = NOTHING
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	END FUNCTION

	'####### �α�˻���  ######
	PUBLIC FUNCTION getPopularKeyWords(idomainno)

		DIM Docruzer,ret
		DIM iRows
		DIM arrData,arrSize
		DIM MaxCnt : MaxCnt =1000
        
        if (FPageSize<MaxCnt) then MaxCnt=FPageSize
            
		SET Docruzer = Server.CreateObject("ATLKSearch.Client")

		IF NOT InitDocruzer(Docruzer) THEN
		    SET Docruzer = Nothing
			EXIT Function
		END IF
		
        SvrAddr = G_ORGSCH_ADDR  '' 106������ �ϴ�

		ret = Docruzer.getPopularKeyword _
						(SvrAddr&":"&SvrPort,_
						arrSize,arrData,_
						MaxCnt,idomainno)
		IF( ret < 0 ) THEN
			CALL Docruzer.EndSession()
			SET Docruzer = NOTHING
			EXIT FUNCTION
		END IF
		getPopularKeyWords = arrData
		SET arrData = NOTHING
		CALL Docruzer.EndSession()
		SET Docruzer = NOTHING

	END FUNCTION
End Class
%>