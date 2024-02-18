<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

'핑거스 웹페이지에서 가져온다.
'server.mappath("/lib/email") 를 server.mappath("/academy/lib/email") 로 수정한다.

sub sendmail(mailfrom, mailto, mailtitle, mailcontent)
        dim mailobject
        dim cdoMessage,cdoConfig
        
        
        Set cdoConfig = CreateObject("CDO.Configuration")

		'-> 서버 접근방법을 설정합니다
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '1 - (cdoSendUsingPickUp)  2 - (cdoSendUsingPort)

		'-> 서버 주소를 설정합니다
    	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "110.93.128.94"

		'-> 접근할 포트번호를 설정합니다
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

		'-> 접속시도할 제한시간을 설정합니다
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10

		'-> SMTP 접속 인증방법을 설정합니다
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1

		'-> SMTP 서버에 인증할 ID를 입력합니다
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "MailSendUser"

		'-> SMTP 서버에 인증할 암호를 입력합니다
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "wjddlswjddls"

		cdoConfig.Fields.Update

		Set cdoMessage = CreateObject("CDO.Message")

		Set cdoMessage.Configuration = cdoConfig

		cdoMessage.To 				= mailto
		cdoMessage.From 			= mailfrom
		cdoMessage.SubJect 	= mailtitle
		'메일 내용이 텍스트일 경우 cdoMessage.TextBody, html일 경우 cdoMessage.HTMLBody
		cdoMessage.HTMLBody	= mailcontent

		cdoMessage.BodyPart.Charset="ks_c_5601-1987"         '/// 한글을 위해선 꼭 넣어 주어야 합니다.
        cdoMessage.HTMLBodyPart.Charset="ks_c_5601-1987"     '/// 한글을 위해선 꼭 넣어 주어야 합니다.
        
        if (application("Svr_Info")	= "Dev") then
            ''테스트 환경
    		if ((InStr(mailto,"10x10.co.kr")>0) or (mailto="archilee@shinbiro.com")) then
    		    cdoMessage.Send
            end if
        else
		    cdoMessage.Send
		end if

		Set cdoMessage = nothing
		Set cdoConfig = nothing
end sub

sub dsendmail(mailfrom, mailto, mailtitle, mailcontent)
        dim mailobject
        dim cdoMessage,cdoConfig
        
        
        Set cdoConfig = CreateObject("CDO.Configuration")

		'-> 서버 접근방법을 설정합니다
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 '1 - (cdoSendUsingPickUp)  2 - (cdoSendUsingPort)

		'-> 서버 주소를 설정합니다
    	cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "110.93.128.94"

		'-> 접근할 포트번호를 설정합니다
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

		'-> 접속시도할 제한시간을 설정합니다
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10

		'-> SMTP 접속 인증방법을 설정합니다
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1

		'-> SMTP 서버에 인증할 ID를 입력합니다
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "MailSendUser"

		'-> SMTP 서버에 인증할 암호를 입력합니다
		cdoConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "wjddlswjddls"

		cdoConfig.Fields.Update

		Set cdoMessage = CreateObject("CDO.Message")

		Set cdoMessage.Configuration = cdoConfig

		cdoMessage.To 				= mailto
		cdoMessage.From 			= mailfrom
		cdoMessage.SubJect 	= mailtitle
		'메일 내용이 텍스트일 경우 cdoMessage.TextBody, html일 경우 cdoMessage.HTMLBody
		cdoMessage.HTMLBody	= mailcontent

		cdoMessage.BodyPart.Charset="ks_c_5601-1987"         '/// 한글을 위해선 꼭 넣어 주어야 합니다.
        cdoMessage.HTMLBodyPart.Charset="ks_c_5601-1987"     '/// 한글을 위해선 꼭 넣어 주어야 합니다.
        
        if (application("Svr_Info")	= "Dev") then
            ''테스트 환경
    		if ((InStr(mailto,"10x10.co.kr")>0) or (mailto="archilee@shinbiro.com")) then
    		    cdoMessage.Send
            end if
        else
		    cdoMessage.Send
		end if

		Set cdoMessage = nothing
		Set cdoConfig = nothing
end sub

function SendmailFingersNewuser(mailto,userName) ' 가입메일파일을 읽어들이는 방식으로 전환
        dim mailfrom, mailtitle, mailcontent,dirPath,fileName
        dim fs,objFile

        mailfrom = "customer@thefingers.co.kr"
        mailtitle = "텐바이텐 핑거스 사이트 가입을 축하 드립니다."

        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/academy/lib/email")
        fileName = dirPath&"\\email_join.htm"
        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
        mailcontent = replace(mailcontent,":USERNAME:",userName)

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        SendmailFingersNewuser = mailcontent
end function

'강좌알림 메일용
'TODO : 상품결제에 대한 발송메일은 별도 제작필요합니다.(시간상 skip)
function SendmailLectureOrder(orderserial,mailfrom)
        dim sql,discountrate,paymethod
        dim mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal


        mailtitle = "주문이 정상적으로 접수되었습니다!"

        '주문자 메일주소 확인,주문거래종류 선택---------------------------------------------------------------------------
        sql = "select buyemail,accountdiv from [db_academy].[dbo].tbl_academy_order_master where orderserial = '" + orderserial + "'"
        rsAcademyget.Open sql,dbAcademyget,1
        if  not rsAcademyget.EOF  then
                rsAcademyget.Movefirst
                mailto = rsAcademyget("buyemail")
                paymethod = trim(rsAcademyget("accountdiv"))
        else
                exit function
        end if
        rsAcademyget.close

        ' 파일을 불러와서 ---------------------------------------------------------------------------
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/academy/lib/email")
        if paymethod = "7" then    ' 무통장
            fileName = dirPath&"\\email_bank1.htm"
        elseif paymethod = "100" then   ' 신용카드
            fileName = dirPath&"\\email_card1.htm"
        else
        	fileName = dirPath&"\\email_default.htm"
        end if

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
'        mailcontent = replace(mailcontent,":USERNAME:",userName)


		dim SpendMile, tencardspend
        '주문정보 확인.---------------------------------------------------------------------------
        sql = "select top 1 l.lec_title, l.lecturer_name, l.lec_startday1, totalitemno, buyname, reqname, "
        sql = sql + " reqzipcode, (a.reqzipaddr + ' ' + a.reqaddress) as reqalladdress, a.reqphone, a.totalsum,"
        sql = sql + " a.totalmileage, c.itemcost,a.discountrate,a.subtotalprice, a.miletotalprice ,a.tencardspend, a.comment"
        sql = sql + " from [db_academy].[dbo].tbl_academy_order_master a, [db_academy].[dbo].tbl_academy_order_detail c, "
        sql = sql + " [db_academy].[dbo].tbl_lec_item l "
        sql = sql + " where a.orderserial = '" + orderserial + "' and c.orderserial = a.orderserial and c.itemid = l.idx "

		rsAcademyget.Open sql,dbAcademyget,1
        if  not rsAcademyget.EOF  then
                discountrate = rsAcademyget("discountrate")
                tencardspend = rsAcademyget("tencardspend")
                rsAcademyget.Movefirst
                mailcontent = replace(mailcontent,":SUBTOTALPRICE:", formatNumber(FormatCurrency(cstr(rsAcademyget("subtotalprice"))),0) ) ' 주문총액
                'mailcontent = replace(mailcontent,":TOTALITEMPRICE:",  FormatCurrency(cstr(Titemcost - rsAcademyget("itemcost"))) ) ' 주문한 총item  가격
                'mailcontent = replace(mailcontent,":DELIVERYFEE:",  formatNumber(FormatCurrency(cstr(rsAcademyget("itemcost"))),0) ) ' 배송금액
                mailcontent = replace(mailcontent,":BUYNAME:", rsAcademyget("buyname")) ' 주문자 이름
                mailcontent = replace(mailcontent,":USERNAME:", rsAcademyget("buyname"))

                if (rsAcademyget("totalitemno") > 1) then
                    mailcontent = replace(mailcontent,":REQNAME:", CStr(rsAcademyget("buyname")) + " 외 " + CStr(rsAcademyget("totalitemno")-1) + " 명") ' 수강자
                else
                    mailcontent = replace(mailcontent,":REQNAME:", CStr(rsAcademyget("buyname"))) ' 수강자
                end if

                mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' 주문번호
                mailcontent = replace(mailcontent,":REQPHONE:", rsAcademyget("reqphone")) ' 주문자 전화번호
                mailcontent = replace(mailcontent,":BEASONGMEMO:", db2html(rsAcademyget("comment"))) ' 배송메모

                mailcontent = replace(mailcontent,":LECTITLE:", db2html(rsAcademyget("lec_title"))) ' 강좌명
                mailcontent = replace(mailcontent,":LECTURERNAME:", db2html(rsAcademyget("lecturer_name"))) ' 강사명
                mailcontent = replace(mailcontent,":STARTDAY1:", Left(rsAcademyget("lec_startday1"),10)) ' 시작일
                if IsNull(rsAcademyget("miletotalprice")) then
                	SpendMile =""
                else
                	SpendMile = rsAcademyget("miletotalprice") + tencardspend
                	SpendMile = formatNumber(FormatCurrency(SpendMile),0)
            	end if
            	mailcontent = replace(mailcontent,":SPENDMILEAGE:", SpendMile) ' 마일리지
            	mailcontent = replace(mailcontent,":TOTALITEMPRICE:",  formatNumber(rsAcademyget("totalsum"),0) ) ' 주문한 총item  가격
		else
				rsAcademyget.close
                exit function
        end if
        rsAcademyget.close





        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        SendmailLectureOrder = mailcontent
end function

function ReSendmailLectureOrder(orderserial,mailfrom)
        SendmailLectureOrder orderserial,mailfrom
end function

function sendmailbankok(mailto,userName,orderserial) ' 입금완료
        dim sql,discountrate
        dim mailfrom, mailtitle, mailcontent
        dim fs,objFile,dirPath,fileName

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "무통장 입금이 정상적으로 처리 되었습니다!"

        ' 파일을 불러와서
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/academy/lib/email")
        fileName = dirPath&"\\email_bank2.htm"
        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall



		dim SpendMile, tencardspend
        '주문정보 확인.---------------------------------------------------------------------------
        sql = "select top 1 l.lec_title, l.lecturer_name, l.lec_startday1, totalitemno, buyname, reqname, reqzipcode, (a.reqzipaddr + ' ' + a.reqaddress) as reqalladdress, a.reqphone, a.totalsum, a.totalmileage, c.itemcost,a.discountrate,a.subtotalprice, a.miletotalprice ,a.tencardspend, a.comment from [db_academy].[dbo].tbl_academy_order_master a, [db_academy].[dbo].tbl_academy_order_detail c, [db_academy].[dbo].tbl_lec_item l "
        sql = sql + " where a.orderserial = '" + orderserial + "' and c.orderserial = a.orderserial and c.itemid = l.idx "

		rsAcademyget.Open sql,dbAcademyget,1
        if  not rsAcademyget.EOF  then
                discountrate = rsAcademyget("discountrate")
                tencardspend = rsAcademyget("tencardspend")
                rsAcademyget.Movefirst
                mailcontent = replace(mailcontent,":SUBTOTALPRICE:", formatNumber(FormatCurrency(cstr(rsAcademyget("subtotalprice"))),0) ) ' 주문총액
                'mailcontent = replace(mailcontent,":TOTALITEMPRICE:",  FormatCurrency(cstr(Titemcost - rsAcademyget("itemcost"))) ) ' 주문한 총item  가격
                'mailcontent = replace(mailcontent,":DELIVERYFEE:",  formatNumber(FormatCurrency(cstr(rsAcademyget("itemcost"))),0) ) ' 배송금액
                mailcontent = replace(mailcontent,":BUYNAME:", rsAcademyget("buyname")) ' 주문자 이름
                mailcontent = replace(mailcontent,":USERNAME:", rsAcademyget("buyname"))

                if (rsAcademyget("totalitemno") > 1) then
                    mailcontent = replace(mailcontent,":REQNAME:", CStr(rsAcademyget("buyname")) + " 외 " + CStr(rsAcademyget("totalitemno")) + " 명") ' 수강자
                else
                    mailcontent = replace(mailcontent,":REQNAME:", CStr(rsAcademyget("buyname"))) ' 수강자
                end if

                mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' 주문번호
                mailcontent = replace(mailcontent,":REQPHONE:", rsAcademyget("reqphone")) ' 주문자 전화번호
                mailcontent = replace(mailcontent,":BEASONGMEMO:", db2html(rsAcademyget("comment"))) ' 배송메모

                mailcontent = replace(mailcontent,":LECTITLE:", db2html(rsAcademyget("lec_title"))) ' 강좌명
                mailcontent = replace(mailcontent,":LECTURERNAME:", db2html(rsAcademyget("lecturer_name"))) ' 강사명
                mailcontent = replace(mailcontent,":STARTDAY1:", Left(rsAcademyget("lec_startday1"),10)) ' 시작일
                if IsNull(rsAcademyget("miletotalprice")) then
                	SpendMile =""
                else
                	SpendMile = rsAcademyget("miletotalprice") + tencardspend
                	SpendMile = formatNumber(FormatCurrency(SpendMile),0)
            	end if
            	mailcontent = replace(mailcontent,":SPENDMILEAGE:", SpendMile) ' 마일리지
            	mailcontent = replace(mailcontent,":TOTALITEMPRICE:",  formatNumber(rsAcademyget("totalsum"),0) ) ' 주문한 총item  가격
		else
                exit function
        end if
        rsAcademyget.close



        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
'        sendmailbankok = mailcontent
end function

function sendmailfinish(orderserial,deliverno)
        dim sql,discountrate,paymethod
        dim mailfrom, mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal
        dim subtotalprice
        mailfrom = "customer@10x10.co.kr"
        mailtitle = "주문하신 상품에 대한 텐바이텐 배송안내입니다!"


        ' 파일을 불러와서
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/academy/lib/email")
        fileName = dirPath&"\\email_finish.htm"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall

		sql = "select top 1 buyname,buyemail,subtotalprice from [db_academy].[dbo].tbl_academy_order_master"
		sql = sql + " where orderserial = '" + orderserial + "'"
		rsAcademyget.Open sql,dbAcademyget,1
		if  not rsAcademyget.EOF  then
			mailto = rsAcademyget("buyemail")
			subtotalprice = rsAcademyget("subtotalprice")
			mailcontent = replace(mailcontent,":BUYNAME:", db2html(rsAcademyget("buyname"))) ' 주문자 이름
			'if Left(deliverno,1)="6" then
			'	mailcontent = replace(mailcontent,":DELIVERNOWITHSRC:",  "http://www.cjgls.co.kr/contents/gls/gls004/gls004_06_01.asp?slipno=" + CStr(deliverno) ) ' 운송장번호
			'else
				mailcontent = replace(mailcontent,":DELIVERNOWITHSRC:",  "http://www.hydex.net/ehydex/jsp/home/distribution/tracking/trackingViewCus.jsp?InvNo=" + CStr(deliverno) ) ' 운송장번호
			'end if

			mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' 주문번호
		else
			exit function
		end if
		rsAcademyget.close



        'item 루프 앞뒤부분 짜르기
'        beforeItemHtml = Left(mailcontent,InStr(mailcontent,":ITEMSTART:")-1)
'        afterItemHtml = Mid(mailcontent,InStr(mailcontent,":ITEMEND:")+11)

        'item 루프를 돌릴부분 자르기
'        itemHtmlOri = Left(mailcontent,InStr(mailcontent,":ITEMEND:")-1)
'        itemHtmlOri = Mid(itemHtmlOri,InStr(itemHtmlOri,":ITEMSTART:")+11)

        '주문아이템 정보 확인.
        dim itemserial,inx,sinx,einx
		  dim BaesongState
		  dim transco,transurl,songjangstr
'        sql = " select d.itemid, d.itemoptionname, m.imglist, d.itemname,"
'		   sql = sql + " d.itemcost, d.makerid, d.itemno"
'		   sql = sql + " from [db_academy].[dbo].tbl_academy_order_detail d"
'		   sql = sql + " left join [db_item].[dbo].tbl_item_image m on d.itemid=m.itemid"
'        sql = sql + " where d.orderserial = '" + orderserial + "'"
'        sql = sql + " and d.itemid <>0"
'        sql = sql + " and d.cancelyn<>'Y'"

        sql = " select a.itemid, a.itemoptionname, c.smallimage, c.itemname," + vbcrlf
        sql = sql + " (c.itemserial_large + c.itemserial_mid + c.itemserial_small) as itemserial," + vbcrlf
        sql = sql + " a.itemcost as sellcash, a.itemno, a.isupchebeasong, a.songjangdiv, isnull(a.songjangno,'') as songjangno, a.currstate" + vbcrlf
        sql = sql + " from [db_academy].[dbo].tbl_academy_order_detail a," + vbcrlf
        sql = sql + " [db_item].[dbo].tbl_item c" + vbcrlf
        sql = sql + " where a.orderserial = '" + orderserial + "'" + vbcrlf
        sql = sql + " and a.itemid <> '0'" + vbcrlf
        sql = sql + " and c.itemid = a.itemid" + vbcrlf
        sql = sql + " and (a.cancelyn<>'Y')" + vbcrlf
        sql = sql + " order by a.isupchebeasong asc" + vbcrlf

        inx = 0
		  sinx = 1
		  einx = 0
itemHtml = "<table border='0' cellpadding='0' cellspacing='0'>"

        rsAcademyget.Open sql,dbAcademyget,1
        if  not rsAcademyget.EOF  then
                rsAcademyget.Movefirst
                do until rsAcademyget.eof

						  if inx = 0 then
								if rsAcademyget("isupchebeasong") = "N" then
									sinx = 0' 텐바이텐배송이 처음실행될때
									einx = 1
								elseif rsAcademyget("isupchebeasong") = "Y" then
									sinx = 0'업체배송이 처음실행될때
								end if
						  elseif einx = 1 and (rsAcademyget("isupchebeasong") = "Y") then
									einx = 0
									sinx = 0'텐바이텐배송 뿌려준후 업체배송 처음 뿌려줄때
						  end if
'response.write sinx & "<br>"
'response.write einx
'dbget.close()	:	response.End
if sinx = 0 then
itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td>"
itemHtml = itemHtml + "<table width='550' border='0' cellspacing='1' cellpadding='0'>"
itemHtml = itemHtml + "<tr>"
if rsAcademyget("isupchebeasong") = "N" then
itemHtml = itemHtml + "<td align='left' valign='top'><img src='http://www.10x10.co.kr/lib/email/images/deliver_ten_t.gif' width='121' height='30'></td>"
else
itemHtml = itemHtml + "<td align='left' valign='top'><img src='http://www.10x10.co.kr/lib/email/images/deliver_upche_t.gif' width='121' height='30'></td>"
end if
itemHtml = itemHtml + "<td>&nbsp;</td>"
itemHtml = itemHtml + "</tr>"
itemHtml = itemHtml + "</table>"
itemHtml = itemHtml + "</td>"
itemHtml = itemHtml + "</tr>"
itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td>"
itemHtml = itemHtml + "<table style='border-top: 1px solid #aaaaaa' border='0' cellpadding='0' cellspacing='0' height='4' bgcolor='ECECEC'width='550'>"
itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td><img src='http://www.10x10.co.kr/lib/email/images/spacer.gif' width='550' height='4' align='center'></td>"
itemHtml = itemHtml + "</tr>"
itemHtml = itemHtml + "</table>"
itemHtml = itemHtml + "<table style='border-bottom: 1px solid #555555;'width='550' border='0' height='23' cellpadding='0' cellspacing='0'>"
itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td width='50' class='p11' align='center'>상품</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' class='p11' align='center'>상품명</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='80' class='p11' align='center'>옵션</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='30' class='p11' align='center'>수량</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='80' class='p11' align='center'>배송현황</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='100' class='p11' align='center'>택배/송장</td>"
itemHtml = itemHtml + "</tr>"
itemHtml = itemHtml + "</table>"
itemHtml = itemHtml + "</td>"
itemHtml = itemHtml + "</tr>"
end if

'배송상태 지정
if rsAcademyget("isupchebeasong") = "N" then
		 BaesongState = "<font color='red'>출고완료</font>"
else
	 if rsAcademyget("currstate") = 7 then
		 BaesongState = "<font color='red'>출고완료</font>"
	 else
		 BaesongState = "<font color='#004080'>상품준비중</font>"
	 end if
end if

'택배사 지정
if rsAcademyget("songjangdiv") = "1" then
transco = "한진택배"
transurl = "http://www.hanjin.co.kr/transmission/main.htm"
elseif rsAcademyget("songjangdiv") = "2" then
transco = "현대택배"
transurl = "http://www.hydex.net/ehydex/jsp/home/distribution/tracking/trackingViewCus.jsp?InvNo="
elseif rsAcademyget("songjangdiv") = "3" then
transco = "대한통운"
transurl = "http://doortodoor.korex.co.kr/jsp/cmn/index.jsp"
elseif rsAcademyget("songjangdiv") = "4" then
transco = "CJ GLS"
transurl = "http://www.cjgls.co.kr"
elseif rsAcademyget("songjangdiv") = "5" then
transco = "이클라인"
transurl = "http://www.ecline.net/tracking/customer02.html#t01"
elseif rsAcademyget("songjangdiv") = "6" then
transco = "HTH"
transurl = "https://samsunghth.com/homepage/searchTraceGoods/SearchTraceResult.jhtml?dtdShtno="
elseif rsAcademyget("songjangdiv") = "7" then
transco = "훼미리택배"
transurl = "http://www.e-family.co.kr/"
elseif rsAcademyget("songjangdiv") = "8" then
transco = "우체국"
transurl = "http://parcel.epost.go.kr"
elseif rsAcademyget("songjangdiv") = "9" then
transco = "KGB"
transurl = "http://www.kgbl.co.kr/"
elseif rsAcademyget("songjangdiv") = "10" then
transco = "아주택배"
transurl = "http://www.ajulogis.co.kr/"
elseif rsAcademyget("songjangdiv") = "11" then
transco = "오렌지택배"
transurl = "http://www.10x10.co.kr/cscenter/csmain.asp"
elseif rsAcademyget("songjangdiv") = "12" then
transco = "한국택배"
transurl = "http://www.kls.co.kr/"
elseif rsAcademyget("songjangdiv") = "13" then
transco = "옐로우캡"
transurl = "http://www.yellowcap.co.kr/"
elseif rsAcademyget("songjangdiv") = "14" then
transco = "나이스택배"
transurl = "http://www.10x10.co.kr/cscenter/csmain.asp"
elseif rsAcademyget("songjangdiv") = "15" then
transco = "중앙택배"
transurl = "http://www.10x10.co.kr/cscenter/csmain.asp"
elseif rsAcademyget("songjangdiv") = "16" then
transco = "주코택배"
transurl = "http://www.10x10.co.kr/cscenter/csmain.asp"
elseif rsAcademyget("songjangdiv") = "17" then
transco = "트라넷택배"
transurl = "http://www.etranet.co.kr/"
elseif rsAcademyget("songjangdiv") = "18" then
transco = "로젠택배"
transurl = "http://www.ilogen.com/"
elseif rsAcademyget("songjangdiv") = "19" then
transco = "KGB특급택배"
transurl = "http://www.ikgb.co.kr/"
elseif rsAcademyget("songjangdiv") = "20" then
transco = "KT로지스"
transurl = "http://www.kls.co.kr/customer/cus_trace_01.asp"
elseif rsAcademyget("songjangdiv") = "21" then
transco = "경동택배"
transurl = "http://www.kdexp.com"
else
transco = "기타"
transurl = "http://www.10x10.co.kr/cscenter/csmain.asp"
end if

'택배/송장 설정
if rsAcademyget("isupchebeasong") = "N" then
	songjangstr =  "현대택배<br>(<a href='http://www.hydex.net/ehydex/jsp/home/distribution/tracking/trackingViewCus.jsp?InvNo=" + Cstr(deliverno) + "' target='_blank'>" + Cstr(deliverno) + "</a>)"
else
	 If rsAcademyget("songjangdiv") = "2" Then
		  if rsAcademyget("songjangno")<>"" or isnull(rsAcademyget("songjangno")) then
			  songjangstr =  "현대택배<br>(<a href='http://www.hydex.net/ehydex/jsp/home/distribution/tracking/trackingViewCus.jsp?InvNo=" + Cstr(rsAcademyget("songjangno")) + "' target='_blank'>" + rsAcademyget("songjangno") + "</a>)"
		  else
			  songjangstr="-"
		  end if
	 Else
		  if rsAcademyget("songjangno")<>"" or isnull(rsAcademyget("songjangno")) then
			  songjangstr = transco + "<br>(<a href='" + transurl + "' target='_blank'>" + rsAcademyget("songjangno") + "</a>)"
		  else
			  songjangstr="-"
		  end If
	 End If
end if

itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td>"
itemHtml = itemHtml + "<table style='border-bottom: 1px solid #c8c8c8' width='550' border='0' height='57' cellpadding='0' cellspacing='0'>"
itemHtml = itemHtml + "<tr>"
itemHtml = itemHtml + "<td width='50'><img src='http://webimage.10x10.co.kr/image/small/" + cstr( "0" + CStr(Clng(rsAcademyget("itemid")\10000)) + "/" + rsAcademyget("smallimage")) + "' width='50' height='50'></td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6'>" + db2html(rsAcademyget("itemname")) + "</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='80' align='center'>" + rsAcademyget("itemoptionname") + "</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='30' align='center'>" + Cstr(rsAcademyget("itemno")) + "ea</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='80' align='center'>" + BaesongState + "</td>"
itemHtml = itemHtml + "<td style='padding:3 6 3 6' width='100' align='center'>" + songjangstr + "</td>"
itemHtml = itemHtml + "</tr>"
itemHtml = itemHtml + "</table>"
itemHtml = itemHtml + "</td>"
itemHtml = itemHtml + "</tr>"


                inx = inx + 1
                sinx = sinx + 1
                rsAcademyget.movenext
                loop
        else
                exit function
        end if
        rsAcademyget.close

		itemHtml = itemHtml + "</table>"

		itemHtmlTotal = replace(mailcontent,":INNERORDERTABLE:", itemHtml) ' 주문정보테이블 넣기

      mailcontent = itemHtmlTotal

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        sendmailfinish = mailcontent
end function


function sendmailfinish_old(orderserial,deliverno)
        dim sql,discountrate,paymethod
        dim mailfrom, mailto, mailtitle, mailcontent,itemHtml,itemHtmlOri
        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal
        dim subtotalprice
        mailfrom = "customer@10x10.co.kr"
        mailtitle = "주문하신 상품에 대한 텐바이텐 배송안내입니다!"
        '주문자 메일주소 확인,주문거래종류 선택
        sql = "select buyemail,discountrate,subtotalprice from [db_academy].[dbo].tbl_academy_order_master where orderserial = '" + orderserial + "'"
        rsAcademyget.Open sql,dbAcademyget,1
        if  not rsAcademyget.EOF  then
                rsAcademyget.Movefirst
                mailto = rsAcademyget("buyemail")
                discountrate = rsAcademyget("discountrate")
                subtotalprice = rsAcademyget("subtotalprice")
        else
                exit function
        end if
        rsAcademyget.close

        ' 파일을 불러와서
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/academy/lib/email")
        fileName = dirPath&"\\email_finish.htm"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall


        '주문정보 확인.
        sql = "select buyname, reqzipcode, "
        sql = sql + " (a.reqzipaddr + ' ' + a.reqaddress) as reqalladdress, a.totalsum, a.totalmileage, c.itemcost,a.discountrate,a.subtotalprice "
        sql = sql + " from [db_academy].[dbo].tbl_academy_order_master a,  [db_academy].[dbo].tbl_academy_order_detail c"
        sql = sql + " where a.orderserial = '" + orderserial + "' and c.orderserial = '" + orderserial + "' and c.itemid = '0'"
        rsAcademyget.Open sql,dbAcademyget,1
        if  not rsAcademyget.EOF  then
                discountrate = rsAcademyget("discountrate")
                rsAcademyget.Movefirst
                mailcontent = replace(mailcontent,":SUBTOTALPRICE:", FormatCurrency(cstr(rsAcademyget("subtotalprice")))) ' 주문총액
                mailcontent = replace(mailcontent,":TOTALITEMPRICE:",  FormatCurrency(cstr(rsAcademyget("subtotalprice") - rsAcademyget("itemcost"))) ) ' 주문한 총item  가격
                mailcontent = replace(mailcontent,":DELIVERYFEE:",  FormatCurrency(cstr(rsAcademyget("itemcost"))) ) ' 배송금액

                'if (Left(deliverno,1)="6") then
                	mailcontent = replace(mailcontent,":DELIVERNOWITHSRC:",  "http://www.hydex.net/ehydex/jsp/home/distribution/tracking/trackingViewCus.jsp?InvNo=" + CStr(deliverno) ) ' 운송장번호
                'else
                '	mailcontent = replace(mailcontent,":DELIVERNOWITHSRC:",  "http://www.doortodoor.co.kr/html/parcels/Tracking/TrackingResult.asp?TDNUM=" + CStr(deliverno) ) ' 운송장번호
                'end if

                mailcontent = replace(mailcontent,":DELIVERNO:",  deliverno ) ' 운송장번호
                mailcontent = replace(mailcontent,":BUYNAME:", rsAcademyget("buyname")) ' 주문자 이름
                mailcontent = replace(mailcontent,":ORDERSERIAL:", orderserial) ' 주문번호
                mailcontent = replace(mailcontent,":REQZIPCODE:", rsAcademyget("reqzipcode")) ' 배송우편번호
                mailcontent = replace(mailcontent,":REQALLADDRESS:", rsAcademyget("reqalladdress")) ' 배송주소


        else
                exit function
        end if
        rsAcademyget.close

        'item 루프 앞뒤부분 짜르기
        beforeItemHtml = Left(mailcontent,InStr(mailcontent,":ITEMSTART:")-1)
        afterItemHtml = Mid(mailcontent,InStr(mailcontent,":ITEMEND:")+11)

        'item 루프를 돌릴부분 자르기
        itemHtmlOri = Left(mailcontent,InStr(mailcontent,":ITEMEND:")-1)
        itemHtmlOri = Mid(itemHtmlOri,InStr(itemHtmlOri,":ITEMSTART:")+11)

        '주문아이템 정보 확인.
        dim itemserial,inx
        sql = " select a.itemid, a.itemoptionname, c.listimage, c.itemname,"
		sql = sql + " (c.itemserial_large + c.itemserial_mid + c.itemserial_small) as itemserial, c.sellcash, c.makerid, a.itemno"
		sql = sql + " from [db_academy].[dbo].tbl_academy_order_detail a, [db_item].[dbo].tbl_item c"
        sql = sql + " where a.orderserial = '" + orderserial + "' and a.itemid <> '0' and c.itemid = a.itemid"
        sql = sql + " and (a.cancelyn='N' or a.cancelyn='A')"
        inx = 1
        rsAcademyget.Open sql,dbAcademyget,1
        if  not rsAcademyget.EOF  then
                rsAcademyget.Movefirst
                do until rsAcademyget.eof
                        itemserial = rsAcademyget("itemserial") + "-" + FormatCode(rsAcademyget("itemid"))
                        itemHtml = replace(itemHtmlOri,":ITEMSERIAL:", itemserial) ' 상품코드
                        itemHtml = replace(itemHtml,":ITEMNAME:", rsAcademyget("itemname")) ' 상품이름
                        itemHtml = replace(itemHtml,":ITEMPRICE:",  cstr(rsAcademyget("sellcash")*cdbl(discountrate)) ) ' 상품가격
                        itemHtml = replace(itemHtml,":ITEMNO:", cstr(rsAcademyget("itemno"))) ' 수량

						if rsAcademyget("itemoptionname") <> "" then
                        itemHtml = replace(itemHtml,":ITEMOPTIONNAME:", rsAcademyget("itemoptionname")) ' 옵션명
						else
                        itemHtml = replace(itemHtml,":ITEMOPTIONNAME:", "-") ' 옵션명
						end if

                        itemHtml = replace(itemHtml,":IMGLIST:", cstr( "0" + CStr(Clng(rsAcademyget("itemid")\10000)) + "/" + rsAcademyget("listimage"))) ' 상품이미지
                        itemHtml = replace(itemHtml,":MAKERID:", cstr(rsAcademyget("makerid"))) ' 상품이미지

                        itemHtmlTotal = itemHtmlTotal & itemHtml

                inx = inx + 1
                rsAcademyget.movenext
                loop
        else
                exit function
        end if
        rsAcademyget.close

        mailcontent = beforeItemHtml & itemHtmlTotal & afterItemHtml

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        sendmailfinish_old = mailcontent
end function

function sendmailfinish_ting(orderserial,deliverno)
		on error resume next

        dim sql
        dim mailfrom, mailto, mailtitle, mailcontent
        dim itemid, imglist
        dim buyname,itemname, itemoption

        dim fs,objFile,dirPath,fileName,beforeItemHtml,afterItemHtml,itemHtmlTotal

        mailfrom = "tingmart@011ting.com"
        mailtitle = "주문하신 상품에 대한 팅마트 배송안내입니다!"

        '주문자 메일주소 확인,주문거래종류 선택
        sql = "select * from [db_ting].[dbo].tbl_new_ting_orderhistory where orderserial = '" + orderserial + "'"
        rsAcademyget.Open sql,dbAcademyget,1
        if  not rsAcademyget.EOF  then
        		rsAcademyget.movefirst

                mailto = rsAcademyget("buyemail")
                buyname = rsAcademyget("buyname")
                itemname = rsAcademyget("itemname")
                itemid = rsAcademyget("itemid")
                itemoption = rsAcademyget("itemoption")

                rsAcademyget.Close

                sql = "select m.imglist , IsNull(o.codeview,'-') as optname, m.itemid"
                sql = sql + " from [db_item].[dbo].tbl_item_image m"
                sql = sql + " left join [db_item].[dbo].vw_all_option o on o.optioncode='" + CStr(itemoption) + "'"
                sql = sql + " where m.itemid=" + CStr(itemid)

                rsAcademyget.Open sql,dbAcademyget,1
                if Not rsAcademyget.Eof then
                	imglist = "http://image.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(itemid) + "/" + rsAcademyget("listimage")
                	itemoption = rsAcademyget("optname")
                end if
                rsAcademyget.Close
        else
        	rsAcademyget.Close
                exit function
        end if


        ' 파일을 불러와서
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/ext/ting/mail")
        fileName = dirPath & "\\email_finish_ting.htm"

        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall

        mailcontent = replace(mailcontent,"[IBUYNAME]",buyname)
        mailcontent = replace(mailcontent,"[ILISTIMAGE]",imglist)
        mailcontent = replace(mailcontent,"[IITEMNAME]",itemname)
        mailcontent = replace(mailcontent,"[IOPTION]",itemoption)
        mailcontent = replace(mailcontent,"[ISONGJANG]",deliverno)

		if (Left(deliverno,1)="6") then
			mailcontent = replace(mailcontent,"[ISONGJANGWITSRC]","http://www.cjgls.co.kr/contents/gls/gls004/gls004_06_01.asp?slipno=" + deliverno)
		else
			mailcontent = replace(mailcontent,"[ISONGJANGWITSRC]","http://www.doortodoor.co.kr/html/parcels/Tracking/TrackingResult.asp?TDNUM=" + deliverno)
		end if

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        sendmailfinish_ting = mailcontent

        if err then
        	response.write err.description
        end if
end function

function sendmailsearchpass(mailto,userName,imsipass)
        dim mailfrom, mailtitle, mailcontent,dirPath,fileName
        dim fs,objFile

        mailfrom = "customer@10x10.co.kr"
        mailtitle = "[10x10] " + userName + "님의 임시비밀번호 입니다."

        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        dirPath = server.mappath("/academy/lib/email")
        fileName = dirPath&"\\email_searchpass.htm"
        Set objFile = fs.OpenTextFile(fileName,1)
        mailcontent = objFile.readall
        mailcontent = replace(mailcontent,":USERNAME:",userName)
        mailcontent = replace(mailcontent,":IMSIPASS:",imsipass)

        call sendmail(mailfrom, mailto, mailtitle, mailcontent)
        sendmailsearchpass = mailcontent
end function

%>
