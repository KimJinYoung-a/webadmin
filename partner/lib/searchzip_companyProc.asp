<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<%
'// UTF-8 ��ȯ
session.codePage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description :  SCM �����ȣ ã��
' History : 2016.07.01 �ѿ�� ����Ʈ ���� ����
' ���۽������� utf-8�� �⺻�̴�. �մܿ����� �����ϰ� �޴ܿ��� utf-8�� �ް� �����. ���� ������ form ���� �����ؾ� �Ѵ�.
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

	Dim i '// for�� ����
	Dim refer '// ���۷�
	Dim strsql '// ������
	Dim sGubun '// �ּұ���(����, ���θ�+�ǹ���ȣ, ��+����, �ǹ���)
	Dim tmpconfirmVal '// ����Ʈ ���ϰ� ����
	Dim tmppagingVal '// ����¡�� ����
	Dim tmpsReturnCnt '// ���ϰ� �˻����� ī��Ʈ
	Dim sSidoGubun '// �ñ��� ������ ���� �õ���
	Dim tmpReturngungu '// �ñ��� ���ϰ�
	Dim sSido '// �õ���
	Dim sGungu '// �ñ�����
	Dim sRoadName '// ���θ�
	Dim sRoadBno '// ������ȣ��
	Dim sRoaddong '// ���θ� �� �˻���
	Dim sRoadjibun '// ���θ� ���� �˻���
	Dim sRoadBname '// ���θ� �ǹ��� �˻���
	Dim sJibundong '// �����ּ��� �˻���
	Dim tmpOfficial_bld '// �ǹ��� �ӽ����尪
	Dim tmpJibun '// ���� ��ģ��

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


	'// �ٷ� ���ӽÿ� ���� ǥ��
	If InStr(refer, "10x10.co.kr") < 1 Then
		Response.Write "Err|�߸��� �����Դϴ�[99]."
		session.codePage = 949 : dbAnalget.close() : Response.End
	End If

	If Trim(sRoadBno)<>"" Then
		'// �ǹ���ȣ�� "-"���� �Է� �� �� �����Ƿ� üũ�ؼ� �ɷ��ش�.
		If InStr(Trim(sRoadBno),"-")>0 Then
			tmpsRoadBno = Split(sRoadBno, "-")
			sRoadBno = tmpsRoadBno(0)
		End If
		'// "-" üũ�� �Ͽ��µ��� ���ڰ� ������찡 ������ ���ڰ� ������ ƨ�ܳ���.
		If Not(IsNumeric(sRoadBno)) Then
			Response.Write "Err|�ǹ���ȣ�� ���ڸ� �Է����ּ���."
			session.codePage = 949 : dbAnalget.close() : Response.End
		End If
	End If


	Select Case Trim(sGubun)

		Case "jibun" '// ���� �ּҷ� �˻�������
			sJibundong = RepWord(sJibundong,"[^��-�Ra-zA-Z0-9.&%\-\_\s]","")


			'// ��ǰ�˻�
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
				Response.write "OK|<li class='nodata'>�˻��� �ּҰ� �����ϴ�.</li>"
				session.codePage = 949 : dbAnalget.close() : Response.End
			End If
			oDoc.close

		Case "RoadBnumber" '// ���θ� �ּҿ� ���θ� + �ǹ���ȣ�� �˻�������
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

					tmpconfirmVal = tmpconfirmVal&" <span>�����ּ� : "&rsAnalget("sido")&" "&rsAnalget("gungu")
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
				Response.write "OK|<li class='nodata'>�˻��� �ּҰ� �����ϴ�.</li>"
				session.codePage = 949 : dbAnalget.close() : Response.End
			End If
			rsAnalget.close

		Case "RoadBjibun" '// ���θ� �ּҿ� �� + �������� �˻�������
			
			'// ������ �ɰ��� ���� �˻�
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

					tmpconfirmVal = tmpconfirmVal&" <span>�����ּ� : "&rsAnalget("sido")&" "&rsAnalget("gungu")
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
				Response.write "OK|<li class='nodata'>�˻��� �ּҰ� �����ϴ�.</li>"
				session.codePage = 949 : dbAnalget.close() : Response.End
			End If
			rsAnalget.close

		Case "RoadBname" '// ���θ� �ּҿ� �ǹ������� �˻�������
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

					tmpconfirmVal = tmpconfirmVal&" <span>�����ּ� : "&rsAnalget("sido")&" "&rsAnalget("gungu")
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
				Response.write "OK|<li class='nodata'>�˻��� �ּҰ� �����ϴ�.</li>"
				session.codePage = 949 : dbAnalget.close() : Response.End
			End If
			rsAnalget.close

		Case "gungureturn" '// �ñ��� ����Ʈ ����
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
				Response.write "Err|�˻��� ���� �����ϴ�."
				session.codePage = 949 : dbAnalget.close() : Response.End
			End If

			rsAnalget.close
		
	End Select

	'// EUC-KR�� �纯ȯ
	session.codePage = 949
%>

<!-- #include virtual="/lib/db/dbAnalclose.asp" -->