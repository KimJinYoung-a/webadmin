<%
'###########################################################
' Description :  cs 메모
' History : 2007.10.26 한용민 수정
'###########################################################

function DrawMemoDivCombo(selName,selVal)
    dim retVal
    retVal = "<select class='select' name='"&selName&"'>"
    retVal = retVal&"    <option value=''>전체</option>"
    retVal = retVal&"    <option value='00' "& ChkIIF(selVal="00","selected","") &" >배송문의</option>"
    retVal = retVal&"    <option value='01' "& ChkIIF(selVal="01","selected","") &" >주문문의</option>"
    retVal = retVal&"    <option value='02' "& ChkIIF(selVal="02","selected","") &" >상품문의</option>"
    retVal = retVal&"    <option value='03' "& ChkIIF(selVal="03","selected","") &" >재고문의</option>"
    retVal = retVal&"    <option value='04' "& ChkIIF(selVal="04","selected","") &" >취소문의</option>"
    retVal = retVal&"    <option value='05' "& ChkIIF(selVal="05","selected","") &" >환불문의</option>"
    retVal = retVal&"    <option value='06' "& ChkIIF(selVal="06","selected","") &" >교환문의</option>"
    retVal = retVal&"    <option value='07' "& ChkIIF(selVal="07","selected","") &" >AS문의</option>  "
    retVal = retVal&"    <option value='08' "& ChkIIF(selVal="08","selected","") &" >이벤트문의</option>"
    retVal = retVal&"    <option value='09' "& ChkIIF(selVal="09","selected","") &" >증빙서류문의</option> "
    retVal = retVal&"    <option value='10' "& ChkIIF(selVal="10","selected","") &" >시스템문의</option>"
    retVal = retVal&"    <option value='11' "& ChkIIF(selVal="11","selected","") &" >회원제도문의</option>"
    retVal = retVal&"    <option value='12' "& ChkIIF(selVal="12","selected","") &" >회원정보문의</option>"
    retVal = retVal&"    <option value='13' "& ChkIIF(selVal="13","selected","") &" >당첨문의</option>"
    retVal = retVal&"    <option value='14' "& ChkIIF(selVal="14","selected","") &" >반품문의</option>"
    retVal = retVal&"    <option value='15' "& ChkIIF(selVal="15","selected","") &" >입금문의</option>"
    retVal = retVal&"    <option value='16' "& ChkIIF(selVal="16","selected","") &" >오프라인문의</option>"
    retVal = retVal&"    <option value='17' "& ChkIIF(selVal="17","selected","") &" >쿠폰/마일리지문의</option>"
    retVal = retVal&"    <option value='18' "& ChkIIF(selVal="18","selected","") &" >결제방법문의</option>"
    retVal = retVal&"    <option value='20' "& ChkIIF(selVal="20","selected","") &" >기타문의</option>"
    retVal = retVal&"</select>"

    response.write retVal
end function

function getMemoDivName(iqadiv)
    Select Case iqadiv
        Case "00" : getMemoDivName="배송문의"
        Case "01" : getMemoDivName="주문문의"
        Case "02" : getMemoDivName="상품문의"
        Case "03" : getMemoDivName="재고문의"
        Case "04" : getMemoDivName="취소문의"
        Case "05" : getMemoDivName="환불문의"
        Case "06" : getMemoDivName="교환문의"
        Case "07" : getMemoDivName="AS문의"
        Case "08" : getMemoDivName="이벤트문의"
        Case "09" : getMemoDivName="증빙서류문의"
        Case "10" : getMemoDivName="시스템문의"
        Case "11" : getMemoDivName="회원제도문의"
        Case "12" : getMemoDivName="회원정보문의"
        Case "13" : getMemoDivName="당첨문의"
        Case "14" : getMemoDivName="반품문의"
        Case "15" : getMemoDivName="입금문의"
        Case "16" : getMemoDivName="오프라인"
        Case "17" : getMemoDivName="쿠폰/마일리지문의"
        Case "18" : getMemoDivName="결제방법문의"

        Case "20" : getMemoDivName="기타문의"

        Case "50" : getMemoDivName="일반메모"
        Case "51" : getMemoDivName="품절취소요청"
        Case "52" : getMemoDivName="1:1게시판"
        Case "53" : getMemoDivName="D+2업체지연"
        Case "54" : getMemoDivName="D+3업체지연"
        Case "55" : getMemoDivName="업체장기지연"
        Case "56" : getMemoDivName="나의미처리"

        Case "57" : getMemoDivName="업체게시판"
        Case "58" : getMemoDivName="D+4 반품미처리"
        Case "59" : getMemoDivName="제휴몰"
        Case "60" : getMemoDivName="업체문의"

        CASE Else : getMemoDivName=iqadiv
    End Select
end function

Class CCSMemoItem
        public Fid
        public Forderserial
        public Fdivcd
        public FmmGubun
        public Fuserid
        public FphoneNumber

        public Fwriteuser
        public Ffinishuser
        public Fcontents_jupsu
        public Fregdate
        public Ffinishyn
        public Ffinishdate
		public fcontents_div
		public Fqadiv

		public Fretrydate
		public Fspecialmemo
		Public Fsitename

        public function GetDivCDName()
                if Fdivcd="1" then
                        GetDivCDName = "단순"
                elseif Fdivcd="2" then
                        GetDivCDName = "<font color='blue'>요청</font>"
                else
                        GetDivCDName = "-"
                end if
        end function

        public function GetmmGubunName()
                if FmmGubun="0" then
                        GetmmGubunName = "일반"
                elseif FmmGubun="1" then
                        GetmmGubunName = "In"
                elseif FmmGubun="2" then
                        GetmmGubunName = "Out"
                elseif FmmGubun="3" then
                        GetmmGubunName = "업체"
                elseif FmmGubun="4" then
                        GetmmGubunName = "SMS"
                elseif FmmGubun="5" then
                        GetmmGubunName = "EMAIL"
                else
                        GetmmGubunName = FmmGubun
                end if
        end function

		public function GetQaDivName()
            GetQaDivName = getMemoDivName(Fqadiv)
        end function

        public function GetSiteName()
                if Left(Forderserial,1)="A" then
                        GetSiteName = "핑거스"
                else
                        GetSiteName = "텐주문"
                end if
        end function

        Private Sub Class_Initialize()

        End Sub

        Private Sub Class_Terminate()

        End Sub
end Class

Class CCSMemo
        public FItemList()
        public FOneItem

        public FCurrPage
        public FTotalPage
        public FPageSize
        public FResultCount
        public FScrollCount
        public FTotalCount

        public FRectUserID
        public FRectOrderserial
        public FRectId
        public FRectIsFinished
        public FRectSiteGubun
        public FRectPhoneNumber
        public FRectDivCD
        public FRectMMGubun
        public FRectWriteUser
        public FRectQadiv
		Public FRectSiteName

        public FRectContents
		public FRectRegStart
		public FRectRegEnd
		public FRectMMGubunExvlude
		public FRectRetryDateExclude

        public Sub GetCSMemoList()
            dim i,sqlStr

            sqlStr = " select top 100 id, orderserial, divcd, userid, mmgubun, qadiv, phonenumber, qadiv,"
            sqlStr = sqlStr + " writeuser, finishuser, contents_jupsu, regdate, finishyn, finishdate, sitename, retrydate, specialmemo "
            sqlStr = sqlStr + " from " & TABLE_CS_MEMO & " "
            sqlStr = sqlStr + " where 1 = 1 "

            if (FRectUserID <> "") then
                sqlStr = sqlStr + " and userid = '" + CStr(FRectUserID) + "' "
            end if

            if (FRectOrderserial <> "") then
                sqlStr = sqlStr + " and orderserial = '" + CStr(FRectOrderserial) + "' "
            end if

            if (FRectIsFinished <> "") then
                sqlStr = sqlStr + " and finishyn = '" + CStr(FRectIsFinished) + "' "
            end if

		    if (FRectPhoneNumber<> "") then
		         sqlStr = sqlStr + " and phonenumber = '" + CStr(FRectPhoneNumber) + "' "
            end if

            if (FRectWriteUser<>"") then
                sqlStr = sqlStr + " and writeuser = '" + FRectWriteUser + "' "
            end If

			if (FRectSiteName<>"") then
                sqlStr = sqlStr + " and sitename = '" + FRectSiteName + "' "
            end if

            sqlStr = sqlStr + " order by id desc "

            rsACADEMYget.Open sqlStr, dbACADEMYget, 1
			''response.write sqlStr

            FResultCount = rsACADEMYget.RecordCount

            redim preserve FItemList(FResultCount)
            if  not rsACADEMYget.EOF  then
                    i = 0
                    do until rsACADEMYget.eof
                            set FItemList(i) = new CCSMemoItem

                            FItemList(i).Fid                = rsACADEMYget("id")
                            FItemList(i).Forderserial       = rsACADEMYget("orderserial")
                            FItemList(i).Fdivcd             = rsACADEMYget("divcd")
                            FItemList(i).FmmGubun           = rsACADEMYget("mmgubun")
                            FItemList(i).Fuserid            = rsACADEMYget("userid")
                            FItemList(i).FphoneNumber       = rsACADEMYget("phonenumber")
                            FItemList(i).Fwriteuser         = rsACADEMYget("writeuser")
                            FItemList(i).Ffinishuser        = rsACADEMYget("finishuser")
                            FItemList(i).Fcontents_jupsu    = db2html(rsACADEMYget("contents_jupsu"))
                            FItemList(i).Fregdate           = rsACADEMYget("regdate")
                            FItemList(i).Ffinishyn          = rsACADEMYget("finishyn")
                            FItemList(i).Ffinishdate        = rsACADEMYget("finishdate")
                            FItemList(i).Fqadiv             = rsACADEMYget("qadiv")
							FItemList(i).Fsitename        	= rsACADEMYget("sitename")
							FItemList(i).Fretrydate        	= rsACADEMYget("retrydate")
							FItemList(i).Fspecialmemo     	= rsACADEMYget("specialmemo")

                            rsACADEMYget.MoveNext
                            i = i + 1
                    loop
            end if
            rsACADEMYget.close
        end sub

        public Sub GetCSMemoDetail()
            dim i,sqlStr

            sqlStr = " select top 1 id, orderserial, divcd, userid, mmgubun, phonenumber, qadiv, writeuser, finishuser, contents_jupsu,"
            sqlStr = sqlStr + " regdate, finishyn, finishdate, sitename, retrydate, specialmemo "
            sqlStr = sqlStr + " from " & TABLE_CS_MEMO & " "
            sqlStr = sqlStr + " where 1 = 1 "

            if (FRectUserID <> "") then
                sqlStr = sqlStr + " and userid = '" + CStr(FRectUserID) + "' "
            end if

            if (FRectOrderserial <> "") then
                sqlStr = sqlStr + " and orderserial = '" + CStr(FRectOrderserial) + "' "
            end if

			if (FRectId <> "") then
                sqlStr = sqlStr + " and id = '" + CStr(FRectId) + "' "
            end if

            rsACADEMYget.Open sqlStr, dbACADEMYget, 1
			''response.write sqlStr

            if  not rsACADEMYget.EOF  then
                    set FOneItem = new CCSMemoItem

                    FOneItem.Fid                = rsACADEMYget("id")
                    FOneItem.Forderserial       = rsACADEMYget("orderserial")
                    FOneItem.Fdivcd             = rsACADEMYget("divcd")
                    FOneItem.FmmGubun       	= rsACADEMYget("mmgubun")
                    FOneItem.Fuserid        	= rsACADEMYget("userid")
                    FOneItem.FphoneNumber   	= rsACADEMYget("phonenumber")
                    FOneItem.Fwriteuser         = rsACADEMYget("writeuser")
                    FOneItem.Ffinishuser        = rsACADEMYget("finishuser")
                    FOneItem.Fcontents_jupsu    = db2html(rsACADEMYget("contents_jupsu"))
                    FOneItem.Fregdate           = rsACADEMYget("regdate")
                    FOneItem.Ffinishyn          = rsACADEMYget("finishyn")
                    FOneItem.Ffinishdate        = rsACADEMYget("finishdate")
                    FOneItem.Fqadiv        		= rsACADEMYget("qadiv")
					FOneItem.Fsitename        	= rsACADEMYget("sitename")
					FOneItem.Fretrydate        	= rsACADEMYget("retrydate")
					FOneItem.Fspecialmemo     	= rsACADEMYget("specialmemo")
            end if
            rsACADEMYget.close
        end sub

        public Sub GetCSMemoBlankDetail()
            dim i,sqlStr

            set FOneItem = new CCSMemoItem
        end sub

        Private Sub Class_Initialize()
                FCurrPage       = 1
                FPageSize       = 20
                FResultCount    = 0
                FScrollCount    = 10
                FTotalCount     = 0
        End Sub

        Private Sub Class_Terminate()

        End Sub

        public Function HasPreScroll()
                HasPreScroll = StartScrollPage > 1
        end Function

        public Function HasNextScroll()
                HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
        end Function

        public Function StartScrollPage()
                StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
        end Function

end Class

%>
