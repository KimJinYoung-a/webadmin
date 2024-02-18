<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
'// UTF-8 º¯È¯
session.codePage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description :  SCM ¿ìÆí¹øÈ£ Ã£±â
' History : 2016.07.01 ÇÑ¿ë¹Î ÇÁ·ÐÆ® ÀÌÀü »ý¼º
' ¾ÆÀÛ½º¿¡¼­´Â utf-8ÀÌ ±âº»ÀÌ´Ù. ¾Õ´Ü¿¡¼­´Â Æ÷±âÇÏ°í µÞ´Ü¿¡¼­ utf-8·Î ¹Þ°í ½á¾ßÇÔ. Â÷ÈÄ ¹®Á¦½Ã form À¸·Î º¯°æÇØ¾ß ÇÑ´Ù.
'###########################################################
%>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/search/Zipsearchcls.asp" -->
<!-- #include virtual="/lib/util/pageformlib.asp" -->
<%
'	Response.write "OK|<li class='nodata'>aaa.</li>"
'	session.codePage = 949 : dbAnalget.close() : Response.End

	Dim i '// for¹® º¯¼ö
	Dim refer '// ¸®ÆÛ·¯
	Dim strsql '// Äõ¸®¹®
	Dim sGubun '// ÁÖ¼Ò±¸ºÐ(Áö¹ø, µµ·Î¸í+°Ç¹°¹øÈ£, µ¿+Áö¹ø, °Ç¹°¸í)
	Dim tmpconfirmVal '// ¸®½ºÆ® ¸®ÅÏ°ª ÀúÀå
	Dim tmppagingVal '// ÆäÀÌÂ¡°ª ÀúÀå
	Dim tmpsReturnCnt '// ¸®ÅÏ°ª °Ë»ö°¹¼ö Ä«¿îÆ®
	Dim sSidoGubun '// ½Ã±º±¸ ±¸ºÐÀ» À§ÇÑ ½Ãµµ°ª
	Dim tmpReturngungu '// ½Ã±º±¸ ¸®ÅÏ°ª
	Dim sSido '// ½Ãµµ°ª
	Dim sGungu '// ½Ã±º±¸°ª
	Dim sRoadName '// µµ·Î¸í°ª
	Dim sRoadBno '// ºôµù¹øÈ£°ª
	Dim sRoaddong '// µµ·Î¸í µ¿ °Ë»ö°ª
	Dim sRoadjibun '// µµ·Î¸í Áö¹ø °Ë»ö°ª
	Dim sRoadBname '// µµ·Î¸í °Ç¹°¸í °Ë»ö°ª
	Dim sJibundong '// Áö¹øÁÖ¼ÒÀÇ °Ë»ö¾î
	Dim tmpOfficial_bld '// °Ç¹°¸í ÀÓ½ÃÀúÀå°ª
	Dim tmpJibun '// Áö¹ø ÇÕÄ£°ª

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


	'// ¹Ù·Î Á¢¼Ó½Ã¿£ ¿À·ù Ç¥½Ã
	If InStr(refer, "10x10.co.kr") < 1 Then
		Response.Write "Err|Àß¸øµÈ Á¢¼ÓÀÔ´Ï´Ù[99]."
		session.codePage = 949 : dbAnalget.close() : Response.End
	End If

	If Trim(sRoadBno)<>"" Then
		'// °Ç¹°¹øÈ£´Â "-"°ªÀÌ ÀÔ·Â µÉ ¼ö ÀÖÀ¸¹Ç·Î Ã¼Å©ÇØ¼­ °É·¯ÁØ´Ù.
		If InStr(Trim(sRoadBno),"-")>0 Then
			tmpsRoadBno = Split(sRoadBno, "-")
			sRoadBno = tmpsRoadBno(0)
		End If
		'// "-" Ã¼Å©¸¦ ÇÏ¿´´Âµ¥µµ ¹®ÀÚ°¡ ÀÖÀ»°æ¿ì°¡ ÀÖÀ¸´Ï ¹®ÀÚ°¡ ÀÖÀ¸¸é Æ¨°Ü³½´Ù.
		If Not(IsNumeric(sRoadBno)) Then
			Response.Write "Err|°Ç¹°¹øÈ£¿£ ¼ýÀÚ¸¸ ÀÔ·ÂÇØÁÖ¼¼¿ä."
			session.codePage = 949 : dbAnalget.close() : Response.End
		End If
	End If


	Select Case Trim(sGubun)

		Case "jibun" '// Áö¹ø ÁÖ¼Ò·Î °Ë»öÇßÀ»¶§
			sJibundong = RepWord(sJibundong,"[^°¡-ÆRa-zA-Z0-9.&%\-\_\s]","")


			'// »óÇ°°Ë»ö
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
				Response.write "OK|<li class='nodata'>°Ë»öµÈ ÁÖ¼Ò°¡ ¾ø½À´Ï´Ù.</li>"
				session.codePage = 949 : dbAnalget.close() : Response.End
			End If
			oDoc.close

		Case "RoadBnumber" '// µµ·Î¸í ÁÖ¼Ò¿¡ µµ·Î¸í + °Ç¹°¹øÈ£·Î °Ë»öÇßÀ»¶§
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

					tmpconfirmVal = tmpconfirmVal&" <span>Áö¹øÁÖ¼Ò : "&rsAnalget("sido")&" "&rsAnalget("gungu")
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
				Response.write "OK|<li class='nodata'>°Ë»öµÈ ÁÖ¼Ò°¡ ¾ø½À´Ï´Ù.</li>"
				session.codePage = 949 : dbAnalget.close() : Response.End
			End If
			rsAnalget.close

		Case "RoadBjibun" '// µµ·Î¸í ÁÖ¼Ò¿¡ µ¿ + Áö¹øÀ¸·Î °Ë»öÇßÀ»¶§
			
			'// Áö¹øÀ» ÂÉ°³¼­ °¢°¢ °Ë»ö
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

					tmpconfirmVal = tmpconfirmVal&" <span>Áö¹øÁÖ¼Ò : "&rsAnalget("sido")&" "&rsAnalget("gungu")
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
				Response.write "OK|<li class='nodata'>°Ë»öµÈ ÁÖ¼Ò°¡ ¾ø½À´Ï´Ù.</li>"
				session.codePage = 949 : dbAnalget.close() : Response.End
			End If
			rsAnalget.close

		Case "RoadBname" '// µµ·Î¸í ÁÖ¼Ò¿¡ °Ç¹°¸íÀ¸·Î °Ë»öÇßÀ»¶§
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

					tmpconfirmVal = tmpconfirmVal&" <span>Áö¹øÁÖ¼Ò : "&rsAnalget("sido")&" "&rsAnalget("gungu")
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
				Response.write "OK|<li class='nodata'>°Ë»öµÈ ÁÖ¼Ò°¡ ¾ø½À´Ï´Ù.</li>"
				session.codePage = 949 : dbAnalget.close() : Response.End
			End If
			rsAnalget.close

		Case "gungureturn" '// ½Ã±º±¸ ¸®½ºÆ® º¸³¿
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
				Response.write "Err|°Ë»öµÈ °ªÀÌ ¾ø½À´Ï´Ù."
				session.codePage = 949 : dbAnalget.close() : Response.End
			End If

			rsAnalget.close
		
	End Select

	'// EUC-KR·Î Àçº¯È¯
	session.codePage = 949
%>

<!-- #include virtual="/lib/db/dbAnalclose.asp" -->