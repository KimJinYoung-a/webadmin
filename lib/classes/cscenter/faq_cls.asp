<%
'###########################################################
' Description : 고객센터 faq관리 클래스
' Hieditor : 2009.03.02 이영진 생성
'			 2021.07.30 한용민 수정
'###########################################################

'##### FAQ 레코드셋용 클래스 #####
class CfaqItem

    public FfaqId
    public FcommCd
    public Ftitle
    public Fcontents
    public Fuserid
    public Fusername
    public Fregdate
    public FhitCount
    public Fisusing
    public Flinkname
    public Flinkurl
    public Fdisporder
    public Fcomm_name
    public FlastWorker
    public FlastUpdate
    public FlastWorkerName

	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub

end Class


'##### FAQ 클래스 #####
Class Cfaq

	public FfaqList()
	public FTotalCount
	public FCurrPage
	public FTotalPage
	public FPageSize
	public FResultCount
	public FScrollCount
	public FRectfaqid
	public FRectTopCnt
	public FRectsearchDiv
	public FRectsearchKey
	public FRectsearchString
	public FRectisusing

	'// 기본 변수값 지정
	Private Sub Class_Initialize()
		redim preserve FfaqList(0)

		FCurrPage         = 1
		FPageSize         = 10
		FResultCount      = 0
		FScrollCount      = 10
		FTotalCount       = 0
	End Sub

	Private Sub Class_Terminate()

	End Sub


	'// 공지 목록 출력
	public Sub GetFAQList()

		Dim i, strSql, objRs
		Dim paramInfo
		paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
			,Array("@PageSize"		, adInteger	, adParamInput	,		, FPageSize)	_
			,Array("@CurrPage"		, adInteger	, adParamInput	,		, FCurrPage) _
			,Array("@isUsing"		, adVarchar	, adParamInput	, 1		, FRectisusing) _
			,Array("@commCD"		, adVarchar	, adParamInput	, 4		, FRectsearchDiv) _
			,Array("@title"			, adVarchar	, adParamInput	, 50	, FRectsearchString) _
		)
		strSql = "[db_cs].[dbo].sp_Ten_FaqList_Admin"
		Call fnExecSPReturnRSOutput(strSql, paramInfo)

		FTotalCount = CDbl(GetValue(paramInfo, "@RETURN_VALUE"))	' 토탈카운트
		FtotalPage  = Int ( (FTotalCount - 1) / FPageSize ) + 1
		If FTotalCount = 0 Then	FtotalPage = 1
		

		i=0
		if  not rsget.EOF  then
			do until rsget.eof

				redim preserve FfaqList(i)

				set FfaqList(i) = new CfaqItem

				FfaqList(i).Ffaqid		= rsget("faqid")
				FfaqList(i).Ftitle		= rsget("title")
				FfaqList(i).Fuserid		= rsget("userid")
				FfaqList(i).FcommCd		= rsget("commCd")
				FfaqList(i).Fcomm_name	= rsget("commName")
				FfaqList(i).FhitCount	= rsget("hitcount")
				FfaqList(i).Fisusing	= rsget("isusing")
				FfaqList(i).Fregdate	= rsget("regdate")
				FfaqList(i).Flinkname  = rsget("linkname")
				FfaqList(i).Flinkurl   = rsget("linkurl")
				FfaqList(i).Fdisporder	= rsget("disporder")

				i=i+1
				rsget.moveNext
			loop
		end if

		FResultCount = i

		rsget.Close



	end Sub



	'// FAQ 내용 보기
	public Sub GetFAQRead()

		redim FfaqList(0)
		set FfaqList(0) = new CfaqItem

		If FRectFaqId <> "" Then 

			Dim i, strSql, objRs
			Dim paramInfo
			paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
				,Array("@PKID"			, adInteger	, adParamInput	,		, FRectFaqId)	_
			)
			strSql = "[db_cs].[dbo].sp_Ten_FaqOne"
			Call fnExecSPReturnRSOutput(strSql, paramInfo)


			if Not(rsget.EOF or rsget.BOF) then


				FfaqList(0).Ffaqid		= rsget("faqid")
				FfaqList(0).Ftitle		= rsget("title")
				FfaqList(0).Fcontents	= rsget("contents")
				FfaqList(0).Fuserid		= rsget("userid")
				FfaqList(0).Fusername	= rsget("regusername")
				FfaqList(0).FcommCd		= rsget("commCd")
				FfaqList(0).Fcomm_name	= rsget("commName")
				FfaqList(0).Fisusing	= rsget("isusing")
				FfaqList(0).Fregdate	= rsget("regdate")
				FfaqList(0).Flinkname   = rsget("linkname")
				FfaqList(0).Flinkurl    = rsget("linkurl")
				FfaqList(0).Fdisporder	= rsget("disporder")
				FfaqList(0).FlastWorker	= rsget("lastWorker")
				FfaqList(0).FlastUpdate	= rsget("lastUpdate")
				FfaqList(0).FlastWorkerName	= rsget("lastWorkerName")

			end if
			rsget.close

		End If 

	end sub


	'// 공통코드 옵션 생성 //
	function optCommCd(grpCd, nowCd)
		dim SQL, strOpt

		SQL =	"Select comm_cd, comm_name From db_cs.dbo.tbl_cs_comm_code Where comm_group='" & grpCd & "'"
		rsget.Open sql, dbget, 1

		if Not(rsget.EOF or rsget.BOF) then
			Do Until rsget.EOF
				strOpt = strOpt & "<option value='" & rsget("comm_cd") & "' "

				if nowCd=rsget("comm_cd") then strOpt = strOpt & "selected"

				strOpt = strOpt & " >" & rsget("comm_name") & "</option>"
				rsget.MoveNext
			Loop
		end if

		rsget.Close

		optCommCd = strOpt

	end function

	' 등록,수정,삭제,원복,증가
    Public Function ProcData(ByVal mode)

		'On Error Resume Next 
		dbget.BeginTrans

		Dim ErrCode, ErrMsg
        
		Dim strSql
		Dim paramInfo
		paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
			,Array("@mode"			, adVarchar	, adParamInput	, 10	, mode)	_
			,Array("@faqId"		, adInteger	, adParamInput	, 9	, FfaqList(0).FfaqId) _
			,Array("@commCd"		, adVarchar	, adParamInput	, 4	, FfaqList(0).FcommCd) _
			,Array("@title"		, adVarchar	, adParamInput	, 200	, FfaqList(0).Ftitle) _
			,Array("@contents"	, adVarchar	, adParamInput	, 8000	, FfaqList(0).Fcontents) _
			,Array("@userid"		, adVarchar	, adParamInput	, 32	, session("ssBctId")) _
			,Array("@regusername"	, adVarchar	, adParamInput	, 64	, session("ssBctCname")) _
			,Array("@linkname"	, adVarchar	, adParamInput	, 255	, FfaqList(0).Flinkname) _
			,Array("@linkurl"		, adVarchar	, adParamInput	, 255	, FfaqList(0).Flinkurl) _
			,Array("@disporder"	, adInteger	, adParamInput	, 4	, FfaqList(0).Fdisporder) _
			,Array("@isusing"	, adVarchar	, adParamInput	, 1	, FfaqList(0).fisusing) _
		)

		strSql = "db_cs.dbo.sp_Ten_FaqProc"
		Call fnExecSP(strSql, paramInfo)

		ErrCode  = CInt(GetValue(paramInfo, "@RETURN_VALUE"))			' 에러코드

		If Err Or ErrCode <> 0 Then
			dbget.RollBackTrans
			ErrMsg = "오류발생 : " & Err.Number & " : " & Err.Source & " : " & Replace(Err.Description,"'","") & " : " 
		Else
			dbget.CommitTrans
		End If 
		ProcData = ErrMsg 

	End Function 



end Class

%>