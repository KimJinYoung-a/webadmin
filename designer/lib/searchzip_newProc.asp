<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
'// UTF-8 변환
session.codePage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description :  SCM 우편번호 찾기
' History : 2016.07.01 한용민 프론트 이전 생성
' 아작스에서는 utf-8이 기본이다. 앞단에서는 포기하고 뒷단에서 utf-8로 받고 써야함. 차후 문제시 form 으로 변경해야 한다.
'###########################################################
%>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/search/Zipsearchcls.asp" -->
<!-- #include virtual="/lib/util/pageformlib.asp" -->
<%
'	Response.write "OK|<li class='nodata'>aaa.</li>"
'	session.codePage = 949 : dbAnalget.close() : Response.End

	Dim i '// for문 변수
	Dim refer '// 리퍼러
	Dim strsql '// 쿼리문
	Dim sGubun '// 주소구분(지번, 도로명+건물번호, 동+지번, 건물명)
	Dim tmpconfirmVal '// 리스트 리턴값 저장
	Dim tmppagingVal '// 페이징값 저장
	Dim tmpsReturnCnt '// 리턴값 검색갯수 카운트
	Dim sSidoGubun '// 시군구 구분을 위한 시도값
	Dim tmpReturngungu '// 시군구 리턴값
	Dim sSido '// 시도값
	Dim sGungu '// 시군구값
	Dim sRoadName '// 도로명값
	Dim sRoadBno '// 빌딩번호값
	Dim sRoaddong '// 도로명 동 검색값
	Dim sRoadjibun '// 도로명 지번 검색값
	Dim sRoadBname '// 도로명 건물명 검색값
	Dim sJibundong '// 지번주소의 검색어
	Dim tmpOfficial_bld '// 건물명 임시저장값
	Dim tmpJibun '// 지번 합친값

	Dim tmpsRoadBno
	Dim tmpsJibundong
	Dim tmpsJibundongjgubun
	Dim qrysJibundong

	dim PageSize	: PageSize = getNumeric(requestCheckVar(request("psz"),5))
	dim CurrPage : CurrPage = getNumeric(requestCheckVar(request("cpg"),8))
	if CurrPage="" then CurrPage=1
	if PageSize="" then PageSize=10

	tmpconfirmVal = ""
	tmpReturngungu = ""
	qrysJibundong = ""

	refer = request.ServerVariables("HTTP_REFERER")
	sGubun = requestCheckVar(Request("sGubun"),32)
	sJibundong = requestCheckVar(Request("sJibundong"),512)
	sSidoGubun = requestCheckVar(Request("sSidoGubun"),128)
	sSido = requestCheckVar(Request("sSido"),128)
	sGungu = requestCheckVar(Request("sGungu"),128)
	sRoadName = requestCheckVar(Request("sRoadName"),256)
	sRoadBno = requestCheckVar(Request("sRoadBno"),128)
	sRoaddong = requestCheckVar(Request("sRoaddong"),512)
	sRoadjibun = requestCheckVar(Request("sRoadjibun"),128)
	sRoadBname = requestCheckVar(Request("sRoadBname"),256)


	'// 바로 접속시엔 오류 표시
	If InStr(refer, "10x10.co.kr") < 1 Then
		Response.Write "Err|잘못된 접속입니다[99]."
		session.codePage = 949 : dbAnalget.close() : Response.End
	End If

	If Trim(sRoadBno)<>"" Then
		'// 건물번호는 "-"값이 입력 될 수 있으므로 체크해서 걸러준다.
		If InStr(Trim(sRoadBno),"-")>0 Then
			tmpsRoadBno = Split(sRoadBno, "-")
			sRoadBno = tmpsRoadBno(0)
		End If
		'// "-" 체크를 하였는데도 문자가 있을경우가 있으니 문자가 있으면 튕겨낸다.
		If Not(IsNumeric(sRoadBno)) Then
			Response.Write "Err|건물번호엔 숫자만 입력해주세요."
			session.codePage = 949 : dbAnalget.close() : Response.End
		End If
	End If


	Select Case Trim(sGubun)

		Case "jibun" '// 지번 주소로 검색했을때
			sJibundong = RepWord(sJibundong,"[^가-힣a-zA-Z0-9.&%\-\_\s]","")


			'// 상품검색
			dim oDoc,iLp
			set oDoc = new SearchItemCls
			oDoc.FRectSearchTxt = sJibundong        '' search field allwords
			oDoc.FCurrPage = CurrPage
			oDoc.FPageSize = PageSize
			oDoc.getSearchList


			if oDoc.FTotalCount>0 Then
				Dim ii
				IF oDoc.FResultCount >0 then
				    For ii=0 To oDoc.FResultCount -1 
						If IsNull(tmpOfficial_bld)="" Then
							tmpOfficial_bld = ""
						Else
							tmpOfficial_bld = " "&oDoc.FItemList(ii).Fofficial_bld
						End If

						If Trim(oDoc.FItemList(ii).Fjibun_sub)>0 Then
							tmpJibun = oDoc.FItemList(ii).Fjibun_main&"-"&oDoc.FItemList(ii).Fjibun_sub
						Else
							tmpJibun = oDoc.FItemList(ii).Fjibun_main
						End If

						tmpconfirmVal = tmpconfirmVal&"<li><a href="""" onclick=""setAddr('"&Trim(oDoc.FItemList(ii).Fzipcode)&"','"&Trim(oDoc.FItemList(ii).Fsido)&"','"&Trim(oDoc.FItemList(ii).Fgungu)&"','"&Trim(oDoc.FItemList(ii).Fdong)&"','"&Trim(oDoc.FItemList(ii).Feupmyun)&"','"&Trim(oDoc.FItemList(ii).Fri)&"','"&Trim(tmpOfficial_bld)&"','"&Trim(tmpJibun)&"', '', '', 'jibun', 'jibunDetailtxt','jibunDetailAddr2');return false;"";>"&oDoc.FItemList(ii).Fsido&" "&oDoc.FItemList(ii).Fgungu
						If Trim(oDoc.FItemList(ii).Fdong) = "" Then
							tmpconfirmVal = tmpconfirmVal&" "&oDoc.FItemList(ii).Feupmyun
						Else
							tmpconfirmVal = tmpconfirmVal&" "&oDoc.FItemList(ii).Fdong
						End If

						If Trim(oDoc.FItemList(ii).Fri) <> "" Then
							tmpconfirmVal = tmpconfirmVal&" "&oDoc.FItemList(ii).Fri
						End If
						tmpconfirmVal = tmpconfirmVal&" "&Trim(tmpOfficial_bld)&" "&tmpJibun&" </a></li>"
				    Next
					tmppagingVal = fnDisplayPaging_New_nottextboxdirect(CurrPage,oDoc.FTotalCount,PageSize,5,"jsPageGo")
			    end If
				Response.write "OK|"&tmpconfirmVal&"|"&oDoc.FTotalCount&"|"&tmppagingVal
				session.codePage = 949 : dbAnalget.close() : Response.End
			Else
				Response.write "OK|<li class='nodata'>검색된 주소가 없습니다.</li>"
				session.codePage = 949 : dbAnalget.close() : Response.End
			End If
			oDoc.close

		Case "RoadBnumber" '// 도로명 주소에 도로명 + 건물번호로 검색했을때
			strsql = " Select count(idx) From db_zipcode.dbo.new_zipcode Where sido='"&sSido&"' And gungu='"&sGungu&"' And road='"&sRoadName&"' And building_no='"&sRoadBno&"' "
			rsAnalget.Open strsql, dbAnalget, adOpenForwardOnly, adLockReadOnly
			tmpsReturnCnt = rsAnalget(0)

			rsAnalget.close

			strsql = " Select top 100 zipcode, sido, gungu, dong, eupmyun, ri, road "
			strsql = strsql & ", case when isnull(official_bld,'')='' then '' else ' '+official_bld end as official_bld "
			strsql = strsql & ", convert(varchar(10), jibun_main)+case when jibun_sub>0 then '-'+convert(varchar(10), jibun_sub) else '' end as jibun "
			strsql = strsql & ", convert(varchar(10), building_no)+case when building_sub>0 then '-'+convert(varchar(10), building_sub) else '' end as building_no "
			strsql = strsql & " From db_zipcode.dbo.new_zipcode Where sido='"&sSido&"' And gungu='"&sGungu&"' And road='"&sRoadName&"' And building_no='"&sRoadBno&"' "
			rsAnalget.Open strsql, dbAnalget, adOpenForwardOnly, adLockReadOnly
			If Not(rsAnalget.bof Or rsAnalget.eof) Then
				Do Until rsAnalget.eof
					tmpconfirmVal = tmpconfirmVal&"<li><a href="""" onclick=""setAddr('"&Trim(rsAnalget("zipcode"))&"','"&Trim(rsAnalget("sido"))&"','"&Trim(rsAnalget("gungu"))&"','"&Trim(rsAnalget("dong"))&"','"&Trim(rsAnalget("eupmyun"))&"','"&Trim(rsAnalget("ri"))&"','"&Trim(rsAnalget("official_bld"))&"','"&Trim(rsAnalget("jibun"))&"','"&rsAnalget("road")&"','"&rsAnalget("building_no")&"', 'RoadBnumber', 'RoadBnumberDetailTxt','RoadBnumberDetailAddr2');return false;"";>"&rsAnalget("sido")&" "&rsAnalget("gungu")
					If Trim(rsAnalget("eupmyun")) <> "" Then
						tmpconfirmVal = tmpconfirmVal&" "&rsAnalget("eupmyun")
					End If
					tmpconfirmVal = tmpconfirmVal&" "&rsAnalget("road")&" "&rsAnalget("building_no")

					If Trim(rsAnalget("official_bld")) <> "" Then
						tmpconfirmVal = tmpconfirmVal&" "&Trim(rsAnalget("official_bld"))
					End If

					tmpconfirmVal = tmpconfirmVal&" <span>지번주소 : "&rsAnalget("sido")&" "&rsAnalget("gungu")
					If Trim(rsAnalget("dong")) = "" Then
						tmpconfirmVal = tmpconfirmVal&" "&rsAnalget("eupmyun")
					Else
						tmpconfirmVal = tmpconfirmVal&" "&rsAnalget("dong")
					End If
					If Trim(rsAnalget("ri")) <> "" Then
						tmpconfirmVal = tmpconfirmVal&" "&rsAnalget("ri")
					End If
					tmpconfirmVal = tmpconfirmVal&" "&Trim(rsAnalget("official_bld"))&" "&rsAnalget("jibun")&"</span></a></li>"

				rsAnalget.movenext
				Loop
				Response.write "OK|"&tmpconfirmVal&"|"&tmpsReturnCnt
				session.codePage = 949 : dbAnalget.close() : Response.End
			Else
				Response.write "OK|<li class='nodata'>검색된 주소가 없습니다.</li>"
				session.codePage = 949 : dbAnalget.close() : Response.End
			End If
			rsAnalget.close

		Case "RoadBjibun" '// 도로명 주소에 동 + 지번으로 검색했을때
			
			'// 지번을 쪼개서 각각 검색
			If InStr(sRoadjibun,"-")>0 Then
				tmpsJibundongjgubun = Split(sRoadjibun, "-")
				If IsNumeric(tmpsJibundongjgubun(0)) Or IsNumeric(tmpsJibundongjgubun(1)) Then
					qrysJibundong = qrysJibundong & " And jibun_main='"&tmpsJibundongjgubun(0)&"' And jibun_sub='"&tmpsJibundongjgubun(1)&"' "
				End If
			Else
				If IsNumeric(sRoadjibun) Then
					qrysJibundong = qrysJibundong & " And jibun_main='"&sRoadjibun&"' "
				End If
			End If

			strsql = " Select count(idx) From db_zipcode.dbo.new_zipcode Where sido='"&sSido&"' And gungu='"&sGungu&"' And dong='"&sRoaddong&"' "&qrysJibundong
			rsAnalget.Open strsql, dbAnalget, adOpenForwardOnly, adLockReadOnly
			tmpsReturnCnt = rsAnalget(0)

			rsAnalget.close

			strsql = " Select top 100 zipcode, sido, gungu, dong, eupmyun, ri, road "
			strsql = strsql & ", case when isnull(official_bld,'')='' then '' else ' '+official_bld end as official_bld "
			strsql = strsql & ", convert(varchar(10), jibun_main)+case when jibun_sub>0 then '-'+convert(varchar(10), jibun_sub) else '' end as jibun "
			strsql = strsql & ", convert(varchar(10), building_no)+case when building_sub>0 then '-'+convert(varchar(10), building_sub) else '' end as building_no "
			strsql = strsql & " From db_zipcode.dbo.new_zipcode Where sido='"&sSido&"' And gungu='"&sGungu&"' And dong='"&sRoaddong&"' "&qrysJibundong
			rsAnalget.Open strsql, dbAnalget, adOpenForwardOnly, adLockReadOnly
			If Not(rsAnalget.bof Or rsAnalget.eof) Then
				Do Until rsAnalget.eof
					tmpconfirmVal = tmpconfirmVal&"<li><a href="""" onclick=""setAddr('"&Trim(rsAnalget("zipcode"))&"','"&Trim(rsAnalget("sido"))&"','"&Trim(rsAnalget("gungu"))&"','"&Trim(rsAnalget("dong"))&"','"&Trim(rsAnalget("eupmyun"))&"','"&Trim(rsAnalget("ri"))&"','"&Trim(rsAnalget("official_bld"))&"','"&Trim(rsAnalget("jibun"))&"','"&rsAnalget("road")&"','"&rsAnalget("building_no")&"', 'RoadBjibun', 'RoadBjibunDetailTxt','RoadBjibunDetailAddr2');return false;"";>"&rsAnalget("sido")&" "&rsAnalget("gungu")
					If Trim(rsAnalget("eupmyun")) <> "" Then
						tmpconfirmVal = tmpconfirmVal&" "&rsAnalget("eupmyun")
					End If
					tmpconfirmVal = tmpconfirmVal&" "&rsAnalget("road")&" "&rsAnalget("building_no")

					If Trim(rsAnalget("official_bld")) <> "" Then
						tmpconfirmVal = tmpconfirmVal&" "&Trim(rsAnalget("official_bld"))
					End If

					tmpconfirmVal = tmpconfirmVal&" <span>지번주소 : "&rsAnalget("sido")&" "&rsAnalget("gungu")
					If Trim(rsAnalget("dong")) = "" Then
						tmpconfirmVal = tmpconfirmVal&" "&rsAnalget("eupmyun")
					Else
						tmpconfirmVal = tmpconfirmVal&" "&rsAnalget("dong")
					End If
					If Trim(rsAnalget("ri")) <> "" Then
						tmpconfirmVal = tmpconfirmVal&" "&rsAnalget("ri")
					End If
					tmpconfirmVal = tmpconfirmVal&" "&Trim(rsAnalget("official_bld"))&" "&rsAnalget("jibun")&"</span></a></li>"
				rsAnalget.movenext
				Loop

				Response.write "OK|"&tmpconfirmVal&"|"&tmpsReturnCnt
				session.codePage = 949 : dbAnalget.close() : Response.End
			Else
				Response.write "OK|<li class='nodata'>검색된 주소가 없습니다.</li>"
				session.codePage = 949 : dbAnalget.close() : Response.End
			End If
			rsAnalget.close

		Case "RoadBname" '// 도로명 주소에 건물명으로 검색했을때
			strsql = " Select count(idx) From db_zipcode.dbo.new_zipcode Where sido='"&sSido&"' And gungu='"&sGungu&"' And official_bld='"&sRoadBname&"' "
			rsAnalget.Open strsql, dbAnalget, adOpenForwardOnly, adLockReadOnly
			tmpsReturnCnt = rsAnalget(0)

			rsAnalget.close

			strsql = " Select top 100 zipcode, sido, gungu, dong, eupmyun, ri, road "
			strsql = strsql & ", case when isnull(official_bld,'')='' then '' else ' '+official_bld end as official_bld "
			strsql = strsql & ", convert(varchar(10), jibun_main)+case when jibun_sub>0 then '-'+convert(varchar(10), jibun_sub) else '' end as jibun "
			strsql = strsql & ", convert(varchar(10), building_no)+case when building_sub>0 then '-'+convert(varchar(10), building_sub) else '' end as building_no "
			strsql = strsql & " From db_zipcode.dbo.new_zipcode Where sido='"&sSido&"' And gungu='"&sGungu&"' And official_bld='"&sRoadBname&"' "
			rsAnalget.Open strsql, dbAnalget, adOpenForwardOnly, adLockReadOnly
			If Not(rsAnalget.bof Or rsAnalget.eof) Then
				Do Until rsAnalget.eof
					tmpconfirmVal = tmpconfirmVal&"<li><a href="""" onclick=""setAddr('"&Trim(rsAnalget("zipcode"))&"','"&Trim(rsAnalget("sido"))&"','"&Trim(rsAnalget("gungu"))&"','"&Trim(rsAnalget("dong"))&"','"&Trim(rsAnalget("eupmyun"))&"','"&Trim(rsAnalget("ri"))&"','"&Trim(rsAnalget("official_bld"))&"','"&Trim(rsAnalget("jibun"))&"','"&rsAnalget("road")&"','"&rsAnalget("building_no")&"', 'RoadBname', 'RoadBnameDetailTxt','RoadBnameDetailAddr2');return false;"";>"&rsAnalget("sido")&" "&rsAnalget("gungu")
					If Trim(rsAnalget("eupmyun")) <> "" Then
						tmpconfirmVal = tmpconfirmVal&" "&rsAnalget("eupmyun")
					End If
					tmpconfirmVal = tmpconfirmVal&" "&rsAnalget("road")&" "&rsAnalget("building_no")

					If Trim(rsAnalget("official_bld")) <> "" Then
						tmpconfirmVal = tmpconfirmVal&" "&Trim(rsAnalget("official_bld"))
					End If

					tmpconfirmVal = tmpconfirmVal&" <span>지번주소 : "&rsAnalget("sido")&" "&rsAnalget("gungu")
					If Trim(rsAnalget("dong")) = "" Then
						tmpconfirmVal = tmpconfirmVal&" "&rsAnalget("eupmyun")
					Else
						tmpconfirmVal = tmpconfirmVal&" "&rsAnalget("dong")
					End If
					If Trim(rsAnalget("ri")) <> "" Then
						tmpconfirmVal = tmpconfirmVal&" "&rsAnalget("ri")
					End If
					tmpconfirmVal = tmpconfirmVal&" "&Trim(rsAnalget("official_bld"))&" "&rsAnalget("jibun")&"</span></a></li>"
				rsAnalget.movenext
				Loop

				Response.write "OK|"&tmpconfirmVal&"|"&tmpsReturnCnt
				session.codePage = 949 : dbAnalget.close() : Response.End
			Else
				Response.write "OK|<li class='nodata'>검색된 주소가 없습니다.</li>"
				session.codePage = 949 : dbAnalget.close() : Response.End
			End If
			rsAnalget.close

		Case "gungureturn" '// 시군구 리스트 보냄
			strsql = " Select gungu From db_zipcode.[dbo].[new_zipCode_Gungu] Where sido='"&sSidoGubun&"' order by gungu "
			rsAnalget.Open strsql, dbAnalget, adOpenForwardOnly, adLockReadOnly
			If Not(rsAnalget.bof Or rsAnalget.eof) Then
				Do Until rsAnalget.eof
					tmpReturngungu = tmpReturngungu&"<option value='"&rsAnalget("gungu")&"'>"&rsAnalget("gungu")&"</option>"
				rsAnalget.movenext
				Loop

				Response.write "OK|"&tmpReturngungu
				session.codePage = 949 : dbAnalget.close() : Response.End
			Else
				Response.write "Err|검색된 값이 없습니다."
				session.codePage = 949 : dbAnalget.close() : Response.End
			End If

			rsAnalget.close
		
	End Select

	'// EUC-KR로 재변환
	session.codePage = 949
%>

<!-- #include virtual="/lib/db/dbAnalclose.asp" -->