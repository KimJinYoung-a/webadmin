<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
	dim menupos, iLp, sYYYY, sMM, sCDL, sLevel, arrLecIdx, arrSn, mode
	dim strSql, rstMsg

	menupos = RequestCheckvar(request("menupos"),10)
	mode = RequestCheckvar(request("mode"),16)
	sYYYY = RequestCheckvar(request("yyyy"),4)
	sMM = RequestCheckvar(request("mm"),2)
	sCDL = RequestCheckvar(request("cdl"),3)
	sLevel = RequestCheckvar(request("level"),1)
	arrLecIdx = request("arrLecIdx")
	arrSn = request("arrSn")
  	if arrLecIdx <> "" then
		if checkNotValidHTML(arrLecIdx) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end If
  	if arrSn <> "" then
		if checkNotValidHTML(arrSn) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end if	
'ȸ���� �ʿ���� �ϼż� ���� �ϴ� 2�� �߰�(2012-09-17)
	if sYYYY="" then sYYYY=year(date)
	if sMM="" then sMM=Num2Str(Month(date),2,"0","R")

	if mode="" then
		Call Alert_return("�߸��� �����Դϴ�.")
		response.End
	end if

	Select Case mode
		Case "add"
			if arrLecIdx="" then 
				Call Alert_return("����� ���¹�ȣ�� �����ϴ�.")
				response.End
			end if
			if sYYYY="" or sMM="" then 
				Call Alert_return("����� ȸ���� �������� �ʾҽ��ϴ�.")
				response.End
			end if
			if sCDL="" then 
				Call Alert_return("����� ī�װ��� �������� �ʾҽ��ϴ�.")
				response.End
			end if
			if sLevel="" then 
				Call Alert_return("����� ������ ���̵��� �������� �ʾҽ��ϴ�.")
				response.End
			end if

			'�ߺ� üũ
			strSql = "Select lecIdx From [db_academy].[dbo].tbl_lec_pickInfo " &_
					" Where YYYYMM='" & sYYYY & sMM & "'" &_
					"	and lecIdx in (" & arrLecIdx & ")"
			rsACADEMYget.Open strSql, dbACADEMYget, 1
			if Not(rsACADEMYget.EOF or rsACADEMYget.BOF) then
				rstMsg = "�����Ͻ� ȸ���� �̹� ��ϵ� ["
				do until rsACADEMYget.EOF
					if rstMsg="�����Ͻ� ȸ���� �̹� ��ϵ� [" then
						rstMsg = rstMsg & rsACADEMYget(0)
					else
						rstMsg = rstMsg & "," & rsACADEMYget(0)
					end if
					rsACADEMYget.MoveNext
				loop
				rstMsg = rstMsg & "]�� �����ϰ�\n"
			end if
			rsACADEMYget.Close

			'���� ó��
			strSql = "Insert Into [db_academy].[dbo].tbl_lec_pickInfo (YYYYMM,lecLevel,code_large,lecIdx) " &_
					" Select '" & sYYYY & sMM & "', '" & sLevel & "', '" & sCDL & "', idx " &_
					" From [db_academy].[dbo].tbl_lec_item " &_
					" Where idx in (" & arrLecIdx & ")" &_
					"	and idx not in (" &_
					"		Select lecIdx From [db_academy].[dbo].tbl_lec_pickInfo " &_
					" 		Where YYYYMM='" & sYYYY & sMM & "'" &_
					"			and lecIdx in (" & arrLecIdx & "))"
			dbACADEMYget.execute(strSql)

			rstMsg = rstMsg & "���°� ��ϵǾ����ϴ�."
		Case "del"
			if arrSn="" then 
				Call Alert_return("������ ���¹�ȣ�� �����ϴ�.")
				response.End
			end if

			'���� ����
			strSql = "delete from [db_academy].[dbo].tbl_lec_pickInfo " &_
					" Where pickSn in (" & arrSn & ")"
			dbACADEMYget.execute(strSql)

			rstMsg = "�����Ͻ� ���°� �����Ǿ����ϴ�."
	end Select

	Call Alert_Move(rstMsg,"lec_pickInfo.asp?menupos="&menupos&"&yyyy="&sYYYY&"&mm="&sMM&"&cdl="&sCDL&"&level="&sLevel)
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyClose.asp" -->