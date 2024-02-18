<%@  codepage="65001" language="VBScript" %>
<% option Explicit %>
<%Response.Addheader "P3P","policyref='/w3c/p3p.xml', CP='NOI DSP LAW NID PSA ADM OUR IND NAV COM'"%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
response.charset = "utf-8"
%>
<%
'###########################################################
' Description : 주소검색 처리 페이지
' History : 2016.06.16 원승현
'###########################################################
%>
<!-- #include virtual="/lib/db/dbEVTopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp" -->
<!-- #include virtual="/apps/academy/lib/commlib.asp" -->
<!-- #include virtual="/apps/academy/lib/inc_const.asp" -->
<!-- #include virtual="/lib/util/tenEncUtil.asp" -->
<!-- #include virtual="/lib/classes/search/Zipsearchcls.asp" -->
<!-- #include virtual="/apps/academy/lib/pageformlib.asp" -->
<%
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
	Dim zipcodeTableVal '// 우편번호 테이블
	Dim zipcodeGugunVal '// 우편번호 구군

	Dim tmpsRoadBno
	Dim tmpsJibundong
	Dim tmpsJibundongjgubun
	Dim qrysJibundong

	dim PageSize	: PageSize = getNumeric(requestCheckVar(request("psz"),5))
	dim CurrPage : CurrPage = getNumeric(requestCheckVar(request("cpg"),8))
	if CurrPage="" then CurrPage=1
	if PageSize="" then PageSize=5

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


	zipcodeTableVal = "new_zipcode_160823"
	zipcodeGugunVal = "new_zipCode_Gungu160823"

	If sGubun="RoadBnumber" Then
		If Trim(sRoadBno)<>"" Then
			'// 건물번호는 "-"값이 입력 될 수 있으므로 체크해서 걸러준다.
			If InStr(Trim(sRoadBno),"-")>0 Then
				tmpsRoadBno = Split(sRoadBno, "-")
				sRoadBno = tmpsRoadBno(0)
			End If
			'// "-" 체크를 하였는데도 문자가 있을경우가 있으니 문자가 있으면 튕겨낸다.
			If Not(IsNumeric(sRoadBno)) Then
				Response.Write "Err|건물번호엔 숫자만 입력해주세요."
				Response.End
			End If
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
						tmpconfirmVal = tmpconfirmVal&" "&Trim(tmpOfficial_bld)&" "&tmpJibun
						tmpconfirmVal = tmpconfirmVal&"<span>도로명 주소 : "&oDoc.FItemList(ii).Fsido&" "&oDoc.FItemList(ii).Fgungu
						tmpconfirmVal = tmpconfirmVal&" "&oDoc.FItemList(ii).Froad
						If Trim(oDoc.FItemList(ii).Fbuilding_no)<>"" Then
							tmpconfirmVal = tmpconfirmVal&" "&Trim(oDoc.FItemList(ii).Fbuilding_no)
						End If
						If Trim(oDoc.FItemList(ii).Fofficial_bld) <> "" Then
							tmpconfirmVal = tmpconfirmVal&" "&Trim(oDoc.FItemList(ii).Fofficial_bld)
						End If
						If Trim(oDoc.FItemList(ii).Feupmyun) <> "" Then
							tmpconfirmVal = tmpconfirmVal&"("&oDoc.FItemList(ii).Feupmyun&")</span>"
						End If
						tmpconfirmVal = tmpconfirmVal&"</a></li>"
				    Next
					tmppagingVal = fnDisplayPaging_NewMobile(CurrPage,oDoc.FTotalCount,PageSize,"jsPageGo")
			    end If
				Response.write "OK|"&tmpconfirmVal&"|"&oDoc.FTotalCount&"|"&tmppagingVal
				Response.End
			Else
				Response.write "OK|<li class='nodata'>검색된 주소가 없습니다.</li>"
				Response.End
			End If
			oDoc.close

		Case "RoadBnumber" '// 도로명 주소에 도로명 + 건물번호로 검색했을때
			strsql = " Select count(idx) From db_zipcode.dbo."&zipcodeTableVal&" Where sido='"&sSido&"' And gungu='"&sGungu&"' And road='"&sRoadName&"' And building_no='"&sRoadBno&"' "
			rsEVTget.Open strsql, dbEVTget, adOpenForwardOnly, adLockReadOnly
			tmpsReturnCnt = rsEVTget(0)

			rsEVTget.close

			strsql = " Select top 100 zipcode, sido, gungu, dong, eupmyun, ri, road "
			strsql = strsql & ", case when isnull(official_bld,'')='' then '' else ' '+official_bld end as official_bld "
			strsql = strsql & ", convert(varchar(10), jibun_main)+case when jibun_sub>0 then '-'+convert(varchar(10), jibun_sub) else '' end as jibun "
			strsql = strsql & ", convert(varchar(10), building_no)+case when building_sub>0 then '-'+convert(varchar(10), building_sub) else '' end as building_no "
			strsql = strsql & " From db_zipcode.dbo."&zipcodeTableVal&" Where sido='"&sSido&"' And gungu='"&sGungu&"' And road='"&sRoadName&"' And building_no='"&sRoadBno&"' "
			rsEVTget.Open strsql, dbEVTget, adOpenForwardOnly, adLockReadOnly
			If Not(rsEVTget.bof Or rsEVTget.eof) Then
				Do Until rsEVTget.eof
					tmpconfirmVal = tmpconfirmVal&"<li><a href="""" onclick=""setAddr('"&Trim(rsEVTget("zipcode"))&"','"&Trim(rsEVTget("sido"))&"','"&Trim(rsEVTget("gungu"))&"','"&Trim(rsEVTget("dong"))&"','"&Trim(rsEVTget("eupmyun"))&"','"&Trim(rsEVTget("ri"))&"','"&Trim(rsEVTget("official_bld"))&"','"&Trim(rsEVTget("jibun"))&"','"&rsEVTget("road")&"','"&rsEVTget("building_no")&"', 'RoadBnumber', 'RoadBnumberDetailTxt','RoadBnumberDetailAddr2');return false;"";>"&rsEVTget("sido")&" "&rsEVTget("gungu")
					If Trim(rsEVTget("eupmyun")) <> "" Then
						tmpconfirmVal = tmpconfirmVal&" "&rsEVTget("eupmyun")
					End If
					tmpconfirmVal = tmpconfirmVal&" "&rsEVTget("road")&" "&rsEVTget("building_no")

					If Trim(rsEVTget("official_bld")) <> "" Then
						tmpconfirmVal = tmpconfirmVal&" "&Trim(rsEVTget("official_bld"))
					End If

					tmpconfirmVal = tmpconfirmVal&" <span>지번주소 : "&rsEVTget("sido")&" "&rsEVTget("gungu")
					If Trim(rsEVTget("dong")) = "" Then
						tmpconfirmVal = tmpconfirmVal&" "&rsEVTget("eupmyun")
					Else
						tmpconfirmVal = tmpconfirmVal&" "&rsEVTget("dong")
					End If
					If Trim(rsEVTget("ri")) <> "" Then
						tmpconfirmVal = tmpconfirmVal&" "&rsEVTget("ri")
					End If
					tmpconfirmVal = tmpconfirmVal&" "&Trim(rsEVTget("official_bld"))&" "&rsEVTget("jibun")&"</span></a></li>"

				rsEVTget.movenext
				Loop
				Response.write "OK|"&tmpconfirmVal&"|"&tmpsReturnCnt
				Response.End
			Else
				Response.write "OK|<li class='nodata'>검색된 주소가 없습니다.</li>"
				Response.End
			End If
			rsEVTget.close

		Case "RoadBjibun" '// 도로명 주소에 동 + 지번으로 검색했을때
			
			'// 지번을 쪼개서 각각 검색
			If InStr(sRoadjibun,"-")>0 Then
				tmpsJibundongjgubun = Split(sRoadjibun, "-")
				If IsNumeric(tmpsJibundongjgubun(0)) Then
					qrysJibundong = qrysJibundong & " And jibun_main='"&tmpsJibundongjgubun(0)&"'  "
				End If

				If IsNumeric(tmpsJibundongjgubun(1)) Then
					qrysJibundong = qrysJibundong & " And jibun_sub='"&tmpsJibundongjgubun(1)&"' "
				End If
			Else
				If IsNumeric(sRoadjibun) Then
					qrysJibundong = qrysJibundong & " And contains(FtextJibun, '"""&sRoadjibun&"""') "
				End If
			End If

			strsql = " Select count(idx) From db_zipcode.dbo."&zipcodeTableVal&" Where sido='"&sSido&"' And gungu='"&sGungu&"' And dong='"&sRoaddong&"' "&qrysJibundong
			rsEVTget.Open strsql, dbEVTget, adOpenForwardOnly, adLockReadOnly
			tmpsReturnCnt = rsEVTget(0)

			rsEVTget.close

			strsql = " Select top 100 zipcode, sido, gungu, dong, eupmyun, ri, road "
			strsql = strsql & ", case when isnull(official_bld,'')='' then '' else ' '+official_bld end as official_bld "
			strsql = strsql & ", convert(varchar(10), jibun_main)+case when jibun_sub>0 then '-'+convert(varchar(10), jibun_sub) else '' end as jibun "
			strsql = strsql & ", convert(varchar(10), building_no)+case when building_sub>0 then '-'+convert(varchar(10), building_sub) else '' end as building_no "
			strsql = strsql & " From db_zipcode.dbo."&zipcodeTableVal&" Where sido='"&sSido&"' And gungu='"&sGungu&"' And dong='"&sRoaddong&"' "&qrysJibundong
			rsEVTget.Open strsql, dbEVTget, adOpenForwardOnly, adLockReadOnly
			If Not(rsEVTget.bof Or rsEVTget.eof) Then
				Do Until rsEVTget.eof
					tmpconfirmVal = tmpconfirmVal&"<li><a href="""" onclick=""setAddr('"&Trim(rsEVTget("zipcode"))&"','"&Trim(rsEVTget("sido"))&"','"&Trim(rsEVTget("gungu"))&"','"&Trim(rsEVTget("dong"))&"','"&Trim(rsEVTget("eupmyun"))&"','"&Trim(rsEVTget("ri"))&"','"&Trim(rsEVTget("official_bld"))&"','"&Trim(rsEVTget("jibun"))&"','"&rsEVTget("road")&"','"&rsEVTget("building_no")&"', 'RoadBjibun', 'RoadBjibunDetailTxt','RoadBjibunDetailAddr2');return false;"";>"&rsEVTget("sido")&" "&rsEVTget("gungu")
					If Trim(rsEVTget("eupmyun")) <> "" Then
						tmpconfirmVal = tmpconfirmVal&" "&rsEVTget("eupmyun")
					End If
					tmpconfirmVal = tmpconfirmVal&" "&rsEVTget("road")&" "&rsEVTget("building_no")

					If Trim(rsEVTget("official_bld")) <> "" Then
						tmpconfirmVal = tmpconfirmVal&" "&Trim(rsEVTget("official_bld"))
					End If

					tmpconfirmVal = tmpconfirmVal&" <span>지번주소 : "&rsEVTget("sido")&" "&rsEVTget("gungu")
					If Trim(rsEVTget("dong")) = "" Then
						tmpconfirmVal = tmpconfirmVal&" "&rsEVTget("eupmyun")
					Else
						tmpconfirmVal = tmpconfirmVal&" "&rsEVTget("dong")
					End If
					If Trim(rsEVTget("ri")) <> "" Then
						tmpconfirmVal = tmpconfirmVal&" "&rsEVTget("ri")
					End If
					tmpconfirmVal = tmpconfirmVal&" "&Trim(rsEVTget("official_bld"))&" "&rsEVTget("jibun")&"</span></a></li>"
				rsEVTget.movenext
				Loop

				Response.write "OK|"&tmpconfirmVal&"|"&tmpsReturnCnt
				Response.End
			Else
				Response.write "OK|<li class='nodata'>검색된 주소가 없습니다.</li>"
				Response.End
			End If
			rsEVTget.close

		Case "RoadBname" '// 도로명 주소에 건물명으로 검색했을때
			strsql = " Select count(idx) From db_zipcode.dbo."&zipcodeTableVal&" Where sido='"&sSido&"' And gungu='"&sGungu&"' And official_bld='"&sRoadBname&"' "
			rsEVTget.Open strsql, dbEVTget, adOpenForwardOnly, adLockReadOnly
			tmpsReturnCnt = rsEVTget(0)

			rsEVTget.close

			strsql = " Select top 100 zipcode, sido, gungu, dong, eupmyun, ri, road "
			strsql = strsql & ", case when isnull(official_bld,'')='' then '' else ' '+official_bld end as official_bld "
			strsql = strsql & ", convert(varchar(10), jibun_main)+case when jibun_sub>0 then '-'+convert(varchar(10), jibun_sub) else '' end as jibun "
			strsql = strsql & ", convert(varchar(10), building_no)+case when building_sub>0 then '-'+convert(varchar(10), building_sub) else '' end as building_no "
			strsql = strsql & " From db_zipcode.dbo."&zipcodeTableVal&" Where sido='"&sSido&"' And gungu='"&sGungu&"' And official_bld='"&sRoadBname&"' "
			rsEVTget.Open strsql, dbEVTget, adOpenForwardOnly, adLockReadOnly
			If Not(rsEVTget.bof Or rsEVTget.eof) Then
				Do Until rsEVTget.eof
					tmpconfirmVal = tmpconfirmVal&"<li><a href="""" onclick=""setAddr('"&Trim(rsEVTget("zipcode"))&"','"&Trim(rsEVTget("sido"))&"','"&Trim(rsEVTget("gungu"))&"','"&Trim(rsEVTget("dong"))&"','"&Trim(rsEVTget("eupmyun"))&"','"&Trim(rsEVTget("ri"))&"','"&Trim(rsEVTget("official_bld"))&"','"&Trim(rsEVTget("jibun"))&"','"&rsEVTget("road")&"','"&rsEVTget("building_no")&"', 'RoadBname', 'RoadBnameDetailTxt','RoadBnameDetailAddr2');return false;"";>"&rsEVTget("sido")&" "&rsEVTget("gungu")
					If Trim(rsEVTget("eupmyun")) <> "" Then
						tmpconfirmVal = tmpconfirmVal&" "&rsEVTget("eupmyun")
					End If
					tmpconfirmVal = tmpconfirmVal&" "&rsEVTget("road")&" "&rsEVTget("building_no")

					If Trim(rsEVTget("official_bld")) <> "" Then
						tmpconfirmVal = tmpconfirmVal&" "&Trim(rsEVTget("official_bld"))
					End If

					tmpconfirmVal = tmpconfirmVal&" <span>지번주소 : "&rsEVTget("sido")&" "&rsEVTget("gungu")
					If Trim(rsEVTget("dong")) = "" Then
						tmpconfirmVal = tmpconfirmVal&" "&rsEVTget("eupmyun")
					Else
						tmpconfirmVal = tmpconfirmVal&" "&rsEVTget("dong")
					End If
					If Trim(rsEVTget("ri")) <> "" Then
						tmpconfirmVal = tmpconfirmVal&" "&rsEVTget("ri")
					End If
					tmpconfirmVal = tmpconfirmVal&" "&Trim(rsEVTget("official_bld"))&" "&rsEVTget("jibun")&"</span></a></li>"
				rsEVTget.movenext
				Loop

				Response.write "OK|"&tmpconfirmVal&"|"&tmpsReturnCnt
				Response.End
			Else
				Response.write "OK|<li class='nodata'>검색된 주소가 없습니다.</li>"
				Response.End
			End If
			rsEVTget.close

		Case "gungureturn" '// 시군구 리스트 보냄
			strsql = " Select gungu From db_zipcode.[dbo].["&zipcodeGugunVal&"] Where sido='"&sSidoGubun&"' order by gungu "
			rsEVTget.Open strsql, dbEVTget, adOpenForwardOnly, adLockReadOnly
			If Not(rsEVTget.bof Or rsEVTget.eof) Then
				Do Until rsEVTget.eof
					tmpReturngungu = tmpReturngungu&"<option value='"&rsEVTget("gungu")&"'>"&rsEVTget("gungu")&"</option>"
				rsEVTget.movenext
				Loop

				Response.write "OK|"&tmpReturngungu
				Response.End
			Else
				Response.write "Err|검색된 값이 없습니다."
				Response.End
			End If

			rsEVTget.close
		
	End Select

%>
<!-- #include virtual="/lib/db/dbEVTclose.asp" -->