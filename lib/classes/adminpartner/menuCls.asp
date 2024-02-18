<%
'################################################
' 업체어드민 메뉴정보
' 2014-05-08 생성
'################################################

Class CMenuList
public FRectMakerID
public FParentMenu()
public FParentMenuName()
public FChildMenu()
public FParentSize
public FChildSize()

public Fmenuposnotice
public Fmenuposhelp
public FRectID
public FRectUserDiv

	Private Sub Class_Initialize()
		redim FParentMenu(0)
		redim FParentMenuName(0)
		redim FChildSize(0)
	End Sub
	Private Sub Class_Terminate()
	End Sub

	'브랜드가 오푸샵인지 온라인지 확인 - 메뉴 달라짐
	public Function fnChkOffShop(byref isOffUpBeaExists) ''수정
	dim strSql
		strSql ="[db_partner].[dbo].sp_Ten_partnerA_CheckOFFShop('"&FRectMakerID&"')"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnChkOffShop = true
			isOffUpBeaExists = (rsget("defaultbeasongdiv")>0)
	    ELSE
	        fnChkOffShop = false
	        isOffUpBeaExists = false
		END IF
		rsget.close
	End Function

	''기존 개발 소스
	public Function fnChkOffShop_OLD
		Dim objCmd
		Set objCmd = Server.CreateObject("ADODB.COMMAND")
		With objCmd
			.ActiveConnection = dbget
			.CommandType = adCmdText
			.CommandText = "{?= call db_partner.[dbo].[sp_Ten_partnerA_CheckOFFShop]('"&FRectMakerID&"' )}"
			.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
			.Execute, , adExecuteNoRecords
			End With

			IF objCmd(0).Value = 1 THEN
				fnChkOffShop_OLD = True
			ELSE
				fnChkOffShop_OLD = False
			END IF

		Set objCmd = nothing
	End Function

	public Function fnGetParentMenu
	dim strSql
		strSql ="[db_partner].[dbo].sp_Ten_partnerA_getParentMenuList ('"&FRectUserDiv&"' ) "
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetParentMenu = rsget.getRows()
		END IF
		rsget.close
	End Function

	public Function fnGetChileMenu
	dim strSql
		strSql ="[db_partner].[dbo].sp_Ten_partnerA_getChildMenuList ('"&FRectUserDiv&"' )"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			fnGetChileMenu = rsget.getRows()
		END IF
		rsget.close
	End Function


	'업체어드민 메뉴 뿌리기
	'상위메뉴 순번값으로 배열지정 후 FParentMenu(순번)
	'하위메뉴 이차배열로 저장 FChildMenu(FParentMenu(순번), 상위메뉴순번에 해당하는 하위메뉴순번)
	public Sub sbMenuList
		Dim isOffShop, arrList, arrCList, intLoop, intCLoop
		Dim i ,intV ,iniV,iniOldV, intX, intMaxX

		arrList = fnGetParentMenu '상위 메뉴
		arrCList = fnGetChileMenu '하위 메뉴
		FParentSize = 0

		IF isArray(arrList) THEN '---- 상위메뉴 배열 순번으로 저장----------
			FParentSize = Ubound(arrList,2)
			redim preserve FParentMenu(FParentSize)
			redim preserve FParentMenuName(FParentSize)

			For intLoop = 0 To FParentSize
				FParentMenu(intLoop) =  arrList(0,intLoop)
				FParentMenuName(intLoop) = replace(arrList(1,intLoop),"[업체] ","")
			Next
		END IF


			redim FChildMenu(FParentSize,0)
			redim FChildSize(FParentSize)
			 intX = -1
			 iniV = 0
			 intMaxX = 0
			For intLoop = 0 To  FParentSize
			 	intX = -1
			 	iniOldV =""
			 	For  intCLoop=iniV To uBound(arrCList,2)
					 	if (iniOldV <> arrCList(5,intCLoop) and iniOldV <> "") THEN 	'-- 상위메뉴코드 값이 이전 내용과 같은지 비교해서 서브 for 문의 시작값을 늘린다.(루프반복 최소화 위함)
			 			'	FChildSize(intLoop) = intX
			 				exit For											'-- 상위메뉴코드값 변경시 하위메뉴순번 리셋(0부터 시작) :FChildMenu(0,0),FChildMenu(0,1),FChildMenu(1,0).....
		 				end if

			 		 	IF FParentMenu(intLoop) = arrCList(5,intCLoop) THEN
			 		 		 intX = intX  + 1

			 		 		if (intMaxX<intX) then
									intMaxX = intX
									redim preserve FChildMenu(FParentSize,intMaxX)
							end if

			 		 		 set FChildMenu(intLoop,intX) = new CMenuListConts

								FChildMenu(intLoop,intX).Fid  			= arrCList(0,intCLoop)
								FChildMenu(intLoop,intX).Fmenuname 	= arrCList(1,intCLoop)
								FChildMenu(intLoop,intX).Fhaschild 	= arrCList(3,intCLoop)
								FChildMenu(intLoop,intX).Fviewidx  	= arrCList(4,intCLoop)
								FChildMenu(intLoop,intX).Fparentid		= arrCList(5,intCLoop)
								FChildMenu(intLoop,intX).Fdivcd        = arrCList(6,intCLoop)
								FChildMenu(intLoop,intX).Fisusing      = arrCList(7,intCLoop)
								FChildMenu(intLoop,intX).Fmenucolor    = arrCList(8,intCLoop)
								FChildMenu(intLoop,intX).Fmenuposnotice= arrCList(9,intCLoop)
								FChildMenu(intLoop,intX).Fmenuposhelp  = arrCList(10,intCLoop)
								FChildMenu(intLoop,intX).Fmenuname_En  = arrCList(11,intCLoop)
								FChildMenu(intLoop,intX).FuseSslYN     = arrCList(12,intCLoop)

								''2017/07/10 SSL 관련.
								if (FChildMenu(intLoop,intX).FuseSslYN="Y") then
'								    if (application("Svr_Info") = "Dev") then
'								        FChildMenu(intLoop,intX).Flinkurl  	= "https://testwebadmin.10x10.co.kr" & arrCList(2,intCLoop)
'								    else
'								        FChildMenu(intLoop,intX).Flinkurl  	= "https://webadmin.10x10.co.kr" & arrCList(2,intCLoop)
'								    end if
									FChildMenu(intLoop,intX).Flinkurl  	= getSCMSSLURL & arrCList(2,intCLoop)
								else
'								    if (application("Svr_Info") = "Dev") then
'								        FChildMenu(intLoop,intX).Flinkurl  	= "http://testwebadmin.10x10.co.kr" & arrCList(2,intCLoop)
'								    else
'								        FChildMenu(intLoop,intX).Flinkurl  	= "http://webadmin.10x10.co.kr" & arrCList(2,intCLoop)
'								    end if
									FChildMenu(intLoop,intX).Flinkurl  	= getSCMURL & arrCList(2,intCLoop)
							    end if

			 			ELSE

			 				Exit For
			 			END IF
			 			iniV = iniV + 1
			 			iniOldV = arrCList(5,intCLoop)
				Next
			 	FChildSize(intLoop) = intX
		  Next

	End Sub

	public Function fnGetMenuInfo
		Dim strSql
		strSql = "[db_partner].[dbo].[sp_Ten_partnerA_getMenuInfo]("&FRectID&")"
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
		IF Not (rsget.EOF OR rsget.BOF) THEN
			Fmenuposnotice= arrCList(9,intLoop)
			Fmenuposhelp  = arrCList(10,intLoop)
		END IF
		rsget.close
	end Function
End Class

class CMenuListConts
public Fid
public Fmenuname
public Flinkurl
public Fhaschild
public Fviewidx
public Fparentid
public Fdivcd
public Fisusing
public Fmenucolor
public Fmenuposnotice
public Fmenuposhelp
public Fmenuname_En
public FuseSslYN

	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
End Class
%>
