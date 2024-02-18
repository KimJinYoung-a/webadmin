<%@Language="VBScript" CODEPAGE="65001" %>
<%
Response.CharSet="utf-8"
''Session.codepage="65001"
Response.codepage="65001"
Response.ContentType="text/html;charset=utf-8"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/popheaderUTF8.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/admin/ordermaster/lib/KISA_SEED_ECB.asp" -->
<!-- #include virtual="/admin/ordermaster/lib/KISA_SEED_ECB_Support.asp" -->
<!-- #include virtual="/lib/classes/order/baljucls.asp"-->
<!-- #include virtual="lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="lib/classes/cscenter/cs_aslistcls.asp"-->
<%

dim mode, idx, obalju, i, k
dim objXML, xmlURL, objData, xmlDOM, obj
dim key, regData, securityKey
dim custno, orderno
dim reqno, receiveseq, regino, sender, receivename, countrycd, songjangno
dim emsGubun, apprno, memberID, premiumcd, em_ee
dim gubun
dim sqlStr

mode = RequestCheckVar(request("mode"),32)
idx = RequestCheckVar(request("idx"),32)
emsGubun = RequestCheckVar(request("emsGubun"),32)
gubun = RequestCheckVar(request("gubun"),32)

function removeSpecialChar(str)
	removeSpecialChar = Replace(Replace(str, "=", ""), "&", "")
end function


if (emsGubun = "KPT") then
	'// K-Packet
	key = "4e161da3b4b3c7fae1521606363094"
	securityKey = "4e161da3b4b3c7f9"
	memberID = "tenbyten10"
	custno = "0004759367"
	apprno = "10828J0004"
	premiumcd = "14"
	em_ee = "rl"
else
	'// EMS
	key = "de1581d62ed767fca1478667852662"
	securityKey = "4e161da3b4b3c7f9151623"
	custno = "0000024936"
	apprno = "10042H0468"
	memberID = "+10x10"
	premiumcd = "31"
	em_ee = "em"
end if

select case mode
	case "getapprno":
		'// 계약승인번호 조회
		''http://eship.epost.go.kr/api.EmsPrcPayMethodList.ems?regkey=test& regData=d30c973e3e4b5eb75462423c0aa1
		regData = "custno=" & custno
		response.write "key : " & key & "<br />"
		response.write "securityKey : " & securityKey & "<br />"
		response.write "custno : " & custno & "<br />"
		response.write regData & "<br />"
		regData = SeedECBEncrypt(securityKey, regData)
		xmlURL = "http://eship.epost.go.kr/api.EmsPrcPayMethodList.ems?key=" & key & "&regData=" & regData
		response.write xmlURL & "<br />"

		Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

		objXML.setTimeouts 5 * 000, 5 * 000, 15 * 000, 25 * 000
		objXML.Open "GET", xmlURL, false
		objXML.setRequestHeader "Connection", "keep-alive"
		objXML.setRequestHeader "Host", "eship.epost.go.kr"
		objXML.setRequestHeader "User-Agent", "Apache-HttpClient/4.5.1 (Java/1.8.0_91)"

		objXML.Send()

		if objXML.Status = "200" then
			objData = objXML.ResponseText
		else
			response.write "ERROR : 통신오류"
			dbget.close() : response.end
		end if

		'// XML DOM 생성
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		xmlDOM.LoadXML objData

		Set obj = xmlDOM.selectSingleNode("/xsync/error_code/text()")
		if Not obj is Nothing then
			Set obj = xmlDOM.selectSingleNode("/xsync/message/text()")
			response.write "ERROR : " & obj.nodeValue & "<br />"
			''dbget.close() : response.end
		end if

		''response.write objXML.Status
		response.write "a" & objData & "a"

		Set obj = Nothing
		Set xmlDOM = Nothing
		Set objXML  = Nothing
	case "getcustno":
		''고객번호 조회
		''http://eship.epost.go.kr/api.EmsIdCustnoInfo.ems?regkey=test& regData=980e605c405a3264e00acbb0a341902c
		regData = "memberID=" & memberID
		response.write "key : " & key & "<br />"
		response.write "securityKey : " & securityKey & "<br />"
		response.write "memberID : " & memberID & "<br />"
		response.write regData & "<br />"
		regData = SeedECBEncrypt(securityKey, regData)
		xmlURL = "http://eship.epost.go.kr/api.EmsIdCustnoInfo.ems?key=" & key & "&regData=" & regData
		response.write xmlURL & "<br />"

		Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

		objXML.setTimeouts 5 * 000, 5 * 000, 15 * 000, 25 * 000
		objXML.Open "GET", xmlURL, false
		objXML.setRequestHeader "Connection", "keep-alive"
		objXML.setRequestHeader "Host", "eship.epost.go.kr"
		objXML.setRequestHeader "User-Agent", "Apache-HttpClient/4.5.1 (Java/1.8.0_91)"

		objXML.Send()

		if objXML.Status = "200" then
			objData = objXML.ResponseText
		else
			response.write "ERROR : 통신오류"
			dbget.close() : response.end
		end if

		'// XML DOM 생성
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		xmlDOM.LoadXML objData

		Set obj = xmlDOM.selectSingleNode("/xsync/error_code/text()")
		if Not obj is Nothing then
			Set obj = xmlDOM.selectSingleNode("/xsync/message/text()")
			response.write "ERROR : " & obj.nodeValue & "<br />"
			''dbget.close() : response.end
		end if

		''response.write objXML.Status
		response.write "a" & objData & "a"

		Set obj = Nothing
		Set xmlDOM = Nothing
		Set objXML  = Nothing
	case "recvStat":
		'// 접수확인

		''18031841307
		''18031844496
		orderno = "EG221396857KR"
		regData = "custno=" & custno & "&orderno=" & orderno
		regData = SeedECBEncrypt(securityKey, regData)
		xmlURL = "http://eship.epost.go.kr/api.RetrieveEMSResDset.ems?key=" & key & "&regData=" & regData
		''response.write xmlURL

		Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

		objXML.setTimeouts 5 * 000, 5 * 000, 15 * 000, 25 * 000
		objXML.Open "GET", xmlURL, false
		objXML.setRequestHeader "Connection", "keep-alive"
		objXML.setRequestHeader "Host", "eship.epost.go.kr"
		objXML.setRequestHeader "User-Agent", "Apache-HttpClient/4.5.1 (Java/1.8.0_91)"

		objXML.Send()

		if objXML.Status = "200" then
			objData = objXML.ResponseText
		else
			response.write "ERROR : 통신오류"
			dbget.close() : response.end
		end if

		'// XML DOM 생성
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		xmlDOM.LoadXML objData

		Set obj = xmlDOM.selectSingleNode("/xsync/error/error_code/text()")
		if Not obj is Nothing then
			''ERR-111	요청 데이터(regData)을(를) 입력하여 주세요.	요청데이터(regData) 값이 없는 경우
			''ERR-111	고객번호(custno)를 입력해 주세요.	고객번호가 없을 경우
			''ERR-111	주문번호(orderno)를 입력해 주세요.	주문번호가 없을 경우
			''ERR-123	고객번호(custno) 10자리 초과	고객번호 자리수가 10자리보다 클 경우
			''ERR-125	조회 결과가 없습니다.	조회된 결과가 없을시
			Set obj = xmlDOM.selectSingleNode("/xsync/error/message/text()")
			response.write "ERROR : " & obj.nodeValue & "<br />"
			''dbget.close() : response.end
		end if

		''response.write objXML.Status
		response.write "a" & objData & "a"

		Set obj = Nothing
		Set xmlDOM = Nothing
		Set objXML  = Nothing
	case "sendReq":
		'// http://eship.epost.go.kr/api.EmsApplyInsertReceiveTempCmdNew.ems?key=test& regData=d30c973e3e4b5eb75462423c0aa100bc0ea16fc031987
		''api.EmsApplyInsertReceiveTempCmdNew.ems
		''api.EmsApplyInsertReceiveTempCmdNewDEV.ems

		set obalju = New CBalju
		obalju.FRechWeightGubun = gubun
		obalju.getBaljuDetailListEMS idx
		for i = 0 to Ubound(obalju.FBaljuDetailList) - 1
			regData = "custno=" & custno
			regData = regData & "&apprno=" & apprno

			regData = regData & "&sender=" & removeSpecialChar(obalju.FBaljuDetailList(i).FReqName)
			regData = regData & "&senderzipcode=132010"
			regData = regData & "&senderaddr1=" & removeSpecialChar("Dobong-dong, Seoul")
			regData = regData & "&senderaddr2=" & removeSpecialChar("Dobong 63 Street 3rd Floor women com")
			regData = regData & "&sendertelno1=82"
			regData = regData & "&sendertelno2=" & SplitValue(obalju.FBaljuDetailList(i).FBuyPhone,"-",0)
			regData = regData & "&sendertelno3=" & SplitValue(obalju.FBaljuDetailList(i).FBuyPhone,"-",1)
			regData = regData & "&sendertelno4=" & SplitValue(obalju.FBaljuDetailList(i).FBuyPhone,"-",2)
			regData = regData & "&sendermobile1=82"
			regData = regData & "&sendermobile2=" & SplitValue(obalju.FBaljuDetailList(i).FBuyHp,"-",0)
			regData = regData & "&sendermobile3=" & SplitValue(obalju.FBaljuDetailList(i).FBuyHp,"-",1)
			regData = regData & "&sendermobile4=" & SplitValue(obalju.FBaljuDetailList(i).FBuyHp,"-",2)
			regData = regData & "&senderemail=" & removeSpecialChar(obalju.FBaljuDetailList(i).FBuyEmail)

			''premiumcd
			''31 : EMS / 32 : EMS프리미엄 / 14 : K패킷
			regData = regData & "&premiumcd=" & premiumcd

			regData = regData & "&receivename=" & removeSpecialChar(obalju.FBaljuDetailList(i).FReqName)
			regData = regData & "&receivezipcode=" & removeSpecialChar(obalju.FBaljuDetailList(i).Femszipcode)
			regData = regData & "&receiveaddr1=" & removeSpecialChar(obalju.FBaljuDetailList(i).FReqAddr1)
			regData = regData & "&receiveaddr2=" & removeSpecialChar(obalju.FBaljuDetailList(i).FReqAddr2)
			regData = regData & "&receiveaddr3=" & removeSpecialChar("-")
			if (obalju.FBaljuDetailList(i).FSitename="cnglob10x10") then
				regData = regData & "&receivetelno1=" & SplitValue(obalju.FBaljuDetailList(i).FReqHp,"-",0)
				regData = regData & "&receivetelno2=" & SplitValue(obalju.FBaljuDetailList(i).FReqHp,"-",1)
				regData = regData & "&receivetelno3=" & SplitValue(obalju.FBaljuDetailList(i).FReqHp,"-",2)
				regData = regData & "&receivetelno4=" & SplitValue(obalju.FBaljuDetailList(i).FReqHp,"-",3)
				regData = regData & "&receivetelno=" & removeSpecialChar(obalju.FBaljuDetailList(i).FReqHp)
			else
				regData = regData & "&receivetelno1=" & SplitValue(obalju.FBaljuDetailList(i).FreqPhone,"-",0)
				regData = regData & "&receivetelno2=" & SplitValue(obalju.FBaljuDetailList(i).FreqPhone,"-",1)
				regData = regData & "&receivetelno3=" & SplitValue(obalju.FBaljuDetailList(i).FreqPhone,"-",2)
				regData = regData & "&receivetelno4=" & SplitValue(obalju.FBaljuDetailList(i).FreqPhone,"-",3)
				regData = regData & "&receivetelno=" & removeSpecialChar(obalju.FBaljuDetailList(i).FreqPhone)
			end if
			regData = regData & "&receivemail=" & removeSpecialChar(obalju.FBaljuDetailList(i).FBuyEmail)
			regData = regData & "&countrycd=" & removeSpecialChar(obalju.FBaljuDetailList(i).Fdlvcountrycode)
			regData = regData & "&orderno=" & removeSpecialChar(obalju.FBaljuDetailList(i).FOrderserial)

			''em_ee
			''K-Packet : re(K-Packet), rl(K-Packet Light)
			regData = regData & "&em_ee=" & em_ee		'// 비서류

			regData = regData & "&totweight=" & removeSpecialChar(obalju.FBaljuDetailList(i).FrealWeight)
			regData = regData & "&boyn=" & removeSpecialChar(obalju.FBaljuDetailList(i).FInsureYn)
			if obalju.FBaljuDetailList(i).FInsureYn="Y" then
				regData = regData & "&boprc=" & removeSpecialChar(obalju.FBaljuDetailList(i).FItemTotalSum)
			else
				regData = regData & "&boprc=0"
			end if
			regData = regData & "&contents=" & removeSpecialChar(obalju.FBaljuDetailList(i).FgoodNames)
			regData = regData & "&number=1"
			regData = regData & "&weight=" & removeSpecialChar(obalju.FBaljuDetailList(i).FitemWeigth)
			regData = regData & "&value=" & removeSpecialChar(obalju.FBaljuDetailList(i).FitemUsDollar)
			regData = regData & "&hs_code="
			regData = regData & "&origin=KR"
			regData = regData & "&EM_gubun=" & removeSpecialChar(obalju.FBaljuDetailList(i).FitemGubunName)
			regData = regData & "&bizregno=2118700620"
			regData = regData & "&exportsendprsnnm=TENBYTEN"
			regData = regData & "&exportsendprsnaddr=Dobong-dong, Seoul Dobong 63 Street 3rd Floor women com"
			''response.write regData

			'// 어딘가에서 i 값을 덮어쓰는듯...
			k = i
			regData = SeedECBEncrypt(securityKey, regData)
			i = k

			xmlURL = "http://eship.epost.go.kr/api.EmsApplyInsertReceiveTempCmdNew.ems?key=" & key & "&regData=" & regData
			''xmlURL = "http://eship.epost.go.kr/api.EmsApplyInsertReceiveTempCmdNewDEV.ems?key=" & key & "&regData=" & regData
			''response.write xmlURL

			Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")

			objXML.setTimeouts 5 * 000, 5 * 000, 15 * 000, 25 * 000
			objXML.Open "GET", xmlURL, false
			objXML.setRequestHeader "Connection", "keep-alive"
			objXML.setRequestHeader "Host", "eship.epost.go.kr"
			objXML.setRequestHeader "User-Agent", "Apache-HttpClient/4.5.1 (Java/1.8.0_91)"

			objXML.Send()

			if objXML.Status = "200" then
				objData = objXML.ResponseText
			else
				response.write "ERROR : 통신오류"
				dbget.close() : response.end
			end if

			'// XML DOM 생성
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML objData

			Set obj = xmlDOM.selectSingleNode("/xsync/error_code/text()")
			if Not obj is Nothing then
				Set obj = xmlDOM.selectSingleNode("/xsync/message/text()")
				response.write "ERROR : " & obj.nodeValue & "<br />"
				''dbget.close() : response.end
			else
				Set obj = xmlDOM.selectSingleNode("/xsync/regino/text()")
				songjangno = obj.nodeValue
				'// 송장번호 입력
				sqlStr = " db_order.dbo.sp_Ten_EmsSongjangInput '" & obalju.FBaljuDetailList(i).FOrderserial & "', '" & songjangno & "' "
				dbget.Execute sqlStr
				response.write songjangno & "<br />"
			end if

			Set obj = Nothing
			Set xmlDOM = Nothing
			Set objXML  = Nothing
		next
	case else
		response.write "ERROR : unknown mode " & mode
		response.end
end select

%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
