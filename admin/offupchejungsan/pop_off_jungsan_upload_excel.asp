<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offjungsancls.asp"-->
<%
If (session("ssBctId") = "") or (session("ssBctDiv") <> "9999" and session("ssBctDiv") > "9") then
    response.write "<script language='javascript'>alert('������ ����Ǿ����ϴ�.');</script>"
    dbget.close()	:	response.End
end if

dim searchtype, thismonth, filetitle

searchtype =  request("searchtype")

thismonth = Left(CStr(DateSerial(year(now()),month(now())-1,1)),7)

if (searchtype = "") then
	searchtype = "thismonth"
end if



'==============================================================================
dim ojungsan

set ojungsan = new COffJungsan

if (searchtype = "thismonth") then

	filetitle = "�ݿ�(" & thismonth & ") ���ݰ�꼭"
	ojungsan.FRectNotIncludeWonChon = "on"
	ojungsan.FRectYYYYMM = thismonth
	ojungsan.FRectbankingupflag = "Y"

elseif (searchtype = "prevmonth") then

	filetitle = "���� ���ݰ�꼭"
	ojungsan.FRectYYYYMM = ""
	ojungsan.FRectNotIncludeWonChon = "on"
	ojungsan.FRectNotYYYYMM = thismonth
	ojungsan.FRectbankingupflag = "Y"

elseif (searchtype = "withholding") then

	filetitle = "��õ¡�� �����"
	ojungsan.FRectYYYYMM = ""
	ojungsan.FRectNotYYYYMM = ""
	ojungsan.FRectNotIncludeWonChon = ""
	ojungsan.FRectOnlyIncludeWonChon = "on"
	ojungsan.FRectbankingupflag = "Y"

else

    response.write "<script language='javascript'>alert('�߸��� �����Դϴ�.');</script>"
    dbget.close()	:	response.End

end if


ojungsan.JungsanFixedList

dim ipsum,i
ipsum =0



'Response.ContentType = "application/unknown"
'Response.Write("<meta http-equiv='Content-Type' content='text/html; charset=EUC-KR'>")

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=" & filetitle & ".xls"
Response.CacheControl = "public"



dim bufStr, tmpS
bufStr = ""


response.write "<style type='text/css'>" & VbCrlf
response.write ".txt {mso-number-format:'\@'}" & VbCrlf
response.write "</style>" & VbCrlf

bufStr = "<table><tr><td>����</td><td>����</td><td>����ݾ�</td><td>��ü��</td><td>����ڵ�Ϲ�ȣ</td><td>(��)�ٹ�����</td></tr>"
response.write bufStr & VbCrlf

for i=0 to ojungsan.FresultCount-1
	bufStr = "<tr><td class='txt'>"

	if ojungsan.FItemList(i).Fipkum_bank = "ȫ�ἧ����" then
		bufStr = bufStr & "HSBC"
	elseif ojungsan.FItemList(i).Fipkum_bank = "��������" then
		bufStr = bufStr & "����"
	elseif ojungsan.FItemList(i).Fipkum_bank = "����" then
		bufStr = bufStr & "SC����"
	elseif ojungsan.FItemList(i).Fipkum_bank = "��Ƽ" then
		bufStr = bufStr & "�ѱ���Ƽ"
	else
		bufStr = bufStr & Left(ojungsan.FItemList(i).Fipkum_bank,9)
	end if

	bufStr = bufStr & "</td><td class='txt'>" & Trim(ojungsan.FItemList(i).Fipkum_acctno)

	if (searchtype = "withholding") then
		bufStr = bufStr & "</td><td class='txt'>" & fix(ojungsan.FItemList(i).Ftot_jungsanprice*0.967)
	else
		bufStr = bufStr & "</td><td class='txt'>" & ojungsan.FItemList(i).Ftot_jungsanprice
	end if
	bufStr = bufStr & "</td><td class='txt'>" & Replace(Left(Trim(Replace(ojungsan.FItemList(i).Fcompany_name, " ", "")),9), ",", " ")
	bufStr = bufStr & "</td><td class='txt'>" & ojungsan.FItemList(i).Fcompany_no
	bufStr = bufStr & "</td><td class='txt'>(��)�ٹ�����</td></tr>"

	response.write bufStr & VbCrlf
next
response.write "<table>" & VbCrlf
%>
<%
set ojungsan = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->