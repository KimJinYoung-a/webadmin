<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 60*10		' 10��
%>
<%
'####################################################
' Description : ���޸� ����������
' History : 2017.04.06 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/commissionjungsan_cls.asp"-->
<%
dim refer
	refer = request.ServerVariables("HTTP_REFERER")

if InStr(refer,"10x10.co.kr")<1 and session("ssBctId") <> "tozzinet" then
	Response.Write "�߸��� �����Դϴ�."
	dbget.close() : response.end
end if

dim yyyy, mm, mode, sql, arrlist, bufStr, i, tendb, orderserial, itemnoptionname, cjungsan, rdsite
dim ismobile, csum, arrsum, bufsumStr, jungsandategubun, jungsanCount
	yyyy = requestcheckvar(getNumeric(request("yyyy")),4)
	mm = requestcheckvar(getNumeric(request("mm")),2)
	orderserial = requestcheckvar(getNumeric(request("orderserial")),11)
	itemnoptionname = requestcheckvar(request("itemnoptionname"),10)
	mode = requestcheckvar(request("mode"),32)
	rdsite = requestcheckvar(request("rdsite"),32)
	ismobile = requestcheckvar(getNumeric(request("ismobile")),1)

jungsanCount="N"
if yyyy="" then
	stdate = CStr(Now)
	stdate = DateSerial(Left(stdate,4), CLng(Mid(stdate,6,2)),1)
	yyyy = Left(stdate,4)
	mm = Mid(stdate,6,2)
end if

IF application("Svr_Info")="Dev" THEN
	tendb = "tendb."
end IF

'��Ʈ�� csv �ٿ�ε�
if mode="csvbetween" then
	if trim(yyyy="") or trim(mm="") then
		response.write "<script type='text/javascript'>"
		response.write "	alert('�Ⱓ�� �����ϴ�.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if
	yyyy = trim(yyyy)
	mm = trim(mm)

	Set cjungsan = New Ccommission
		cjungsan.FRectyyyymm = yyyy + "-" + mm
		cjungsan.FPageSize = 100000
		cjungsan.FCurrPage = 1
		cjungsan.frectorderserial = orderserial
		cjungsan.frectitemname = itemnoptionname
		cjungsan.Getcommissionjungsan_between_notpaging()

		if cjungsan.FTotalCount < 1 then
			response.write "<script type='text/javascript'>"
			response.write "	alert('�����Ͱ� �����ϴ�.');"
			response.write "	parent.location.reload()"
			response.write "</script>"
			dbget.close() : response.end
		end if

		arrlist = cjungsan.farrlist

'	Response.Buffer=False
	Response.Expires=0
	response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=����������between_"& yyyy &"��"& mm &"��.csv"
	Response.CacheControl = "public"

	response.write "�ֹ�����,����Ȯ������,�ֹ���ȣ,��ǰ��,�ֹ�����,�ֹ��ݾ�(V.A.T����),��������,������,�ֹ�����,��ҳ�¥" & vbcrlf

	if isarray(arrlist) then
		For i = 0 To ubound(arrlist,2)
		bufStr = ""
		bufStr = bufStr & arrlist(0,i)
		bufStr = bufStr & "," & arrlist(1,i)
		bufStr = bufStr & "," & arrlist(2,i)
		bufStr = bufStr & "," & escapedstring(arrlist(3,i))
		bufStr = bufStr & "," & arrlist(4,i)
		bufStr = bufStr & "," & arrlist(5,i)
		bufStr = bufStr & "," & arrlist(6,i)
		bufStr = bufStr & "," & arrlist(7,i)
		bufStr = bufStr & "," & arrlist(8,i)
		bufStr = bufStr & "," & arrlist(9,i)

		response.write bufStr & VbCrlf

		if i>0 and (i mod 10000)=0 then response.flush	'���� �ʰ��� �������� �߰� �÷���
		next
	end if

	set cjungsan = nothing

'��Ʈ�� �����ۼ�
elseif mode="regbetween" then
	if trim(yyyy="") or trim(mm="") then
		response.write "<script type='text/javascript'>"
		response.write "	alert('�Ⱓ�� �����ϴ�.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if
	yyyy = trim(yyyy)
	mm = trim(mm)

	jungsanCount="0"
	sql = "SELECT count(jd.orderserial) as jungsanCount"
	sql = sql & " from db_jungsan.dbo.tbl_nvshop_jungsan_detail jd with (nolock)" + vbcrlf
	sql = sql & " where jd.rdsite='betweenshop' and jd.jmonth='"& yyyy + "-" + mm &"'"

	'response.write sql & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

	If Not rsget.Eof Then
		jungsanCount = rsget("jungsancount")
	End If

	rsget.close

	if jungsanCount>0 then
		response.write "<script type='text/javascript'>"
		response.write "	alert('�̹� �ۼ��� ���� �����Ͱ� �ֽ��ϴ�.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if

	' ���� ���� �ٲ�.
	if DateSerial(yyyy, mm, "01") >= "2021-03-01" then
		jungsandategubun="d.jungsanfixdate"
	else
		jungsandategubun="m.beadaldate"
	end if

	'/��� �� ���(���� �ۼ� �� ���) ������ ��ü�� ������ �߸�����ٴ���. ����� ��� ��쵵 ����.
	sql = "DECLARE @STDT varchar(10)" & vbcrlf
	sql = sql & " DECLARE @EDDT varchar(10)" & vbcrlf
	sql = sql & " SET @STDT='2014-07-01'" & vbcrlf		'������(�ٲ�������)
	sql = sql & " SET @EDDT='"& dateadd("m", +1, DateSerial(yyyy, mm, "01")) &"'" & vbcrlf 
	sql = sql & " insert into db_jungsan.dbo.tbl_nvshop_jungsan_detail" & vbcrlf
	sql = sql & " (orderserial,itemid,itemoption,cancelDT,jMonth,ismobile,orgorderserial,jumundiv,rdsite" & vbcrlf
	sql = sql & " ,rDate,fixedDate,itemNOptionName,itemno,suppPrc,commpro,commissoin,ordStatName)" & vbcrlf
	sql = sql & " 	select" & vbcrlf
	sql = sql & " 	d.orderserial,d.itemid,d.itemoption,isnull(convert(varchar(10),isNULL(M.canceldate,D.canceldate),21),'')" & vbcrlf
	sql = sql & " 	,convert(varchar(7),"& jungsandategubun &",21) as jMonth" & vbcrlf
	sql = sql & " 	,1 as ismobile,isNULL(m.linkorderserial,m.orderserial),m.jumundiv,isNULL(m.rdsite,'')" & vbcrlf
	sql = sql & " 	,convert(varchar(10),m.regdate,21) as rDate,convert(varchar(10),"& jungsandategubun &",21) as fixedDate" & vbcrlf
	sql = sql & " 	,d.itemname+(CASE WHEN d.itemoptionname<>'' THEN ' ['+d.itemoptionname+']' ELSE '' END)" & vbcrlf
	sql = sql & " 	,d.itemno*-1" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0)*-1 as SuppPrc" & vbcrlf
	sql = sql & " 	,0.06" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0)*(0.06)*-1" & vbcrlf
	sql = sql & " 	,CASE WHEN m.jumundiv='9' THEN '��ǰ�����' " & vbcrlf
	sql = sql & " 		WHEN m.jumundiv='6' THEN '��ȯ�Ϸ�'  ELSE '��������' END" & vbcrlf
	sql = sql & " 	from db_order.dbo.tbl_order_master m" & vbcrlf
	sql = sql & " 	Join db_order.dbo.tbl_order_detail d" & vbcrlf
	sql = sql & " 		on m.orderserial=d.orderserial" & vbcrlf
	sql = sql & " 	join db_jungsan.dbo.tbl_nvshop_jungsan_detail J" & vbcrlf
	sql = sql & " 		on d.orderserial=J.orderserial" & vbcrlf
	sql = sql & " 		and d.itemid=j.itemid" & vbcrlf
	sql = sql & " 		and d.itemoption=j.itemoption" & vbcrlf 	'and ordStatName<>'��������'
	sql = sql & " 	left join db_item.[dbo].[tbl_item_standing_item] si" & vbcrlf
	sql = sql & " 		on d.itemid = si.orgitemid" & vbcrlf
	sql = sql & " 		and d.itemoption = si.orgitemoption" & vbcrlf
	sql = sql & " 	left join db_item.[dbo].[tbl_item_standing_order] so" & vbcrlf		' ���ⱸ�� ������ȸ�� �Ϸ� �Ǹ� ���� ��ǰ�� ��ҷ� ���ư��� ���� ����. ����
	sql = sql & " 		on d.itemid = so.orgitemid" & vbcrlf
	sql = sql & " 		and d.itemoption = so.orgitemoption" & vbcrlf
	sql = sql & " 		and si.endreserveidx = so.reserveidx" & vbcrlf
	sql = sql & " 	where 1=1 and m.rdsite='betweenshop'" & vbcrlf
	sql = sql & " 	and m.ipkumdiv>3" & vbcrlf
	'sql = sql & " 	and m.ipkumdiv=8" & vbcrlf
	sql = sql & " 	and (m.cancelyn<>'N' or  d.cancelyn='Y')" & vbcrlf	'����� CASE
	sql = sql & " 	and "& jungsandategubun &">@STDT and "& jungsandategubun &"<@EDDT" & vbcrlf
	sql = sql & " 	and d.itemid not in (0,100)" & vbcrlf
	sql = sql & " 	and so.orgitemid is null" & vbcrlf

	'response.write sql & "<br>"
	dbget.CommandTimeout = 60*5   ' 5��
	dbget.execute sql

	sql = "DECLARE @STDT varchar(10)" & vbcrlf
	sql = sql & " DECLARE @EDDT varchar(10)" & vbcrlf
	sql = sql & " SET @STDT='2014-07-01'" & vbcrlf		'������(�ٲ�������)
	sql = sql & " SET @EDDT='"& dateadd("m", +1, DateSerial(yyyy, mm, "01")) &"'" & vbcrlf
	sql = sql & " insert into db_jungsan.dbo.tbl_nvshop_jungsan_detail" & vbcrlf
	sql = sql & " (orderserial,itemid,itemoption,cancelDT,jMonth,ismobile,orgorderserial,jumundiv,rdsite" & vbcrlf
	sql = sql & " ,rDate,fixedDate,itemNOptionName,itemno,suppPrc,commpro,commissoin,ordStatName)" & vbcrlf
	sql = sql & " 	select" & vbcrlf
	sql = sql & " 	d.orderserial,d.itemid,d.itemoption,'',convert(varchar(7),"& jungsandategubun &",21) as jMonth" & vbcrlf
	sql = sql & " 	,1 as ismobile,isNULL(m.linkorderserial,m.orderserial),m.jumundiv,isNULL(m.rdsite,'')" & vbcrlf
	sql = sql & " 	,convert(varchar(10),m.regdate,21) as rDate" & vbcrlf
	sql = sql & " 	,convert(varchar(10),"& jungsandategubun &",21) as fixedDate" & vbcrlf
	sql = sql & " 	,d.itemname+(CASE WHEN d.itemoptionname<>'' THEN ' ['+d.itemoptionname+']' ELSE '' END)" & vbcrlf
	sql = sql & " 	,d.itemno" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0) as SuppPrc" & vbcrlf
	sql = sql & " 	,(0.06)" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0)*(0.06)" & vbcrlf
	sql = sql & " 	,CASE WHEN m.jumundiv='9' THEN '��ǰ�Ϸ�'" & vbcrlf 
	sql = sql & " 			WHEN	m.jumundiv='6' THEN '��ȯ�Ϸ�' ELSE '���Ϸ�' END" & vbcrlf
	sql = sql & " 	from db_order.dbo.tbl_order_master m" & vbcrlf
	sql = sql & " 	Join db_order.dbo.tbl_order_detail d" & vbcrlf
	sql = sql & " 		on m.orderserial=d.orderserial" & vbcrlf
	sql = sql & " 	left join db_jungsan.dbo.tbl_nvshop_jungsan_detail J" & vbcrlf
	sql = sql & " 		on d.orderserial=J.orderserial" & vbcrlf
	sql = sql & " 		and d.itemid=j.itemid" & vbcrlf
	sql = sql & " 		and d.itemoption=j.itemoption" & vbcrlf
	sql = sql & " 	where 1=1 and m.rdsite='betweenshop'" & vbcrlf
	sql = sql & " 	and J.itemoption is NULL" & vbcrlf
	'sql = sql & " 	and R.ismobile=0" & vbcrlf 	'������ΰ��
	sql = sql & " 	and m.ipkumdiv>3" & vbcrlf
	sql = sql & " 	and m.ipkumdiv=8" & vbcrlf
	sql = sql & " 	and m.cancelyn='N'" & vbcrlf
	sql = sql & " 	and "& jungsandategubun &">@STDT and "& jungsandategubun &"<@EDDT" & vbcrlf
	sql = sql & " 	and d.itemid not in (0,100)" & vbcrlf
	sql = sql & " 	and d.cancelyn<>'Y'" & vbcrlf   '��Ҵ� ���� ����.
	'sql = sql & " 	order by rDate,d.orderserial" & vbcrlf

	'response.write sql & "<br>"
	dbget.CommandTimeout = 60*5   ' 5��
	dbget.execute sql

	response.write "<script type='text/javascript'>"
	response.write "	alert('"& yyyy &"�� "& mm&"�� between ���������� �ۼ� �Ϸ�.');"
	response.write "	parent.location.reload()"
	response.write "</script>"
	dbget.close() : response.end

'����, ����Ʈ csv �ٿ�ε�
elseif mode="csvdaum" or mode="csvnate" then
	if rdsite="" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('���޸� ������ �����ϴ�.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if

	if trim(yyyy="") or trim(mm="") then
		response.write "<script type='text/javascript'>"
		response.write "	alert('�Ⱓ�� �����ϴ�.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if
	yyyy = trim(yyyy)
	mm = trim(mm)

	Set cjungsan = New Ccommission
		cjungsan.FRectyyyymm = yyyy + "-" + mm
		cjungsan.FPageSize = 100000
		cjungsan.FCurrPage = 1
		cjungsan.frectorderserial = orderserial
		cjungsan.frectitemname = itemnoptionname
		cjungsan.frectrdsite = rdsite
		cjungsan.Getcommissionjungsan_daum_notpaging()

		if cjungsan.FTotalCount < 1 then
			response.write "<script type='text/javascript'>"
			response.write "	alert('�����Ͱ� �����ϴ�.');"
			response.write "	parent.location.reload()"
			response.write "</script>"
			dbget.close() : response.end
		end if

		arrlist = cjungsan.farrlist

'	Response.Buffer=False
	Response.Expires=0
	response.ContentType = "application/vnd.ms-excel"

	if mode="csvdaum" then
		Response.AddHeader "Content-Disposition", "attachment; filename=����������daum_"& yyyy &"��"& mm &"��.csv"
	ELSE
		Response.AddHeader "Content-Disposition", "attachment; filename=����������nate_"& yyyy &"��"& mm &"��.csv"
	end if

	Response.CacheControl = "public"

	response.write "�ֹ�����,�������/Ȯ������,�����ڵ�,����ϱ���,�ֹ���ȣ,��ǰ��,�ֹ�����,�ֹ��ݾ�(V.A.T����),��������,������,�ֹ�����,��ҳ�¥" & vbcrlf

	if isarray(arrlist) then
		For i = 0 To ubound(arrlist,2)
		bufStr = ""
		bufStr = bufStr & arrlist(0,i)
		bufStr = bufStr & "," & arrlist(1,i)
		bufStr = bufStr & "," & arrlist(2,i)
		bufStr = bufStr & "," & arrlist(3,i)
		bufStr = bufStr & "," & arrlist(4,i)
		bufStr = bufStr & "," & escapedstring(arrlist(5,i))
		bufStr = bufStr & "," & arrlist(6,i)
		bufStr = bufStr & "," & arrlist(7,i)
		bufStr = bufStr & "," & arrlist(8,i)
		bufStr = bufStr & "," & arrlist(9,i)
		bufStr = bufStr & "," & arrlist(10,i)
		bufStr = bufStr & "," & arrlist(11,i)

		response.write bufStr & VbCrlf

		if i>0 and (i mod 10000)=0 then response.flush	'���� �ʰ��� �������� �߰� �÷���
		next
	end if

	set cjungsan = nothing

'����, ����Ʈ �����ۼ�
elseif mode="regdaum" or mode="regnate" then
	if rdsite="" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('���޸� ������ �����ϴ�.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if

	if trim(yyyy="") or trim(mm="") then
		response.write "<script type='text/javascript'>"
		response.write "	alert('�Ⱓ�� �����ϴ�.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if
	yyyy = trim(yyyy)
	mm = trim(mm)

	jungsanCount="0"
	sql = "SELECT count(jd.orderserial) as jungsanCount"
	sql = sql & " from db_jungsan.dbo.tbl_nvshop_jungsan_detail jd with (nolock)" + vbcrlf
	sql = sql & " Join db_item.dbo.tbl_Outmall_RdsiteGubun R with (nolock)"
	sql = sql & " 	on jd.rdsite=R.rdsite"
	sql = sql & " 	and R.gubun in ('"& rdsite &"')" + vbcrlf
	sql = sql & " where jd.jmonth='"& yyyy + "-" + mm &"'"

	'response.write sql & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

	If Not rsget.Eof Then
		jungsanCount = rsget("jungsancount")
	End If

	rsget.close

	if jungsanCount>0 then
		response.write "<script type='text/javascript'>"
		response.write "	alert('�̹� �ۼ��� ���� �����Ͱ� �ֽ��ϴ�.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if

	' ���� ���� �ٲ�.
	if DateSerial(yyyy, mm, "01") >= "2021-03-01" then
		jungsandategubun="d.jungsanfixdate"
	else
		jungsandategubun="m.beadaldate"
	end if

	'/��� �� ���(���� �ۼ� �� ���) ������ ��ü�� ������ �߸�����ٴ���. ����� ��� ��쵵 ����.
	sql = "DECLARE @STDT varchar(10)" & vbcrlf
	sql = sql & " DECLARE @EDDT varchar(10)" & vbcrlf
	sql = sql & " SET @STDT='2014-08-01'" & vbcrlf 		'������(�ٲ�������)
	sql = sql & " SET @EDDT='"& dateadd("m", +1, DateSerial(yyyy, mm, "01")) &"'" & vbcrlf
	sql = sql & " insert into db_jungsan.dbo.tbl_nvshop_jungsan_detail" & vbcrlf
	sql = sql & " (orderserial,itemid,itemoption,cancelDT,jMonth,ismobile,orgorderserial,jumundiv,rdsite" & vbcrlf
	sql = sql & " ,rDate,fixedDate,itemNOptionName,itemno,suppPrc,commpro,commissoin,ordStatName)" & vbcrlf
	sql = sql & " 	select" & vbcrlf
	sql = sql & " 	d.orderserial,d.itemid,d.itemoption,isnull(convert(varchar(10),isNULL(M.canceldate,D.canceldate),21),'')" & vbcrlf
	sql = sql & " 	,convert(varchar(7),"& jungsandategubun &",21) as jMonth" & vbcrlf
	sql = sql & " 	,R.ismobile,isNULL(m.linkorderserial,m.orderserial),m.jumundiv,isNULL(m.rdsite,'')" & vbcrlf
	sql = sql & " 	,convert(varchar(10),m.regdate,21) as rDate" & vbcrlf
	sql = sql & " 	,convert(varchar(10),"& jungsandategubun &",21) as fixedDate" & vbcrlf
	sql = sql & " 	,d.itemname+(CASE WHEN d.itemoptionname<>'' THEN ' ['+d.itemoptionname+']' ELSE '' END)" & vbcrlf
	sql = sql & " 	,d.itemno*-1" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0)*-1 as SuppPrc" & vbcrlf
	sql = sql & " 	,(CASE WHEN R.isCharge='Y' THEN 0.03 ELSE 0 END)" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0)*(CASE WHEN R.isCharge='Y' THEN 0.03 ELSE 0 END)*-1" & vbcrlf
	sql = sql & " 	,CASE WHEN m.jumundiv='9' THEN '��ǰ�����'" & vbcrlf
	sql = sql & " 		WHEN m.jumundiv='6' THEN '��ȯ�Ϸ�'  ELSE '��������' END" & vbcrlf
	sql = sql & " 	from db_order.dbo.tbl_order_master m with (nolock)" & vbcrlf
	sql = sql & " 	Join db_order.dbo.tbl_order_detail d with (nolock)" & vbcrlf
	sql = sql & " 		on m.orderserial=d.orderserial" & vbcrlf
	sql = sql & " 	Join db_item.dbo.tbl_Outmall_RdsiteGubun R with (nolock)" & vbcrlf
	sql = sql & " 		on m.rdsite=R.rdsite" & vbcrlf
	sql = sql & " 		and R.gubun in ('"& rdsite &"')" & vbcrlf
	sql = sql & " 	join db_jungsan.dbo.tbl_nvshop_jungsan_detail J with (nolock)" & vbcrlf
	sql = sql & " 		on d.orderserial=J.orderserial" & vbcrlf
	sql = sql & " 		and d.itemid=j.itemid" & vbcrlf
	sql = sql & " 		and d.itemoption=j.itemoption" & vbcrlf 	'and ordStatName<>'��������'
	sql = sql & " 	left join db_item.[dbo].[tbl_item_standing_item] si with (nolock)" & vbcrlf
	sql = sql & " 		on d.itemid = si.orgitemid" & vbcrlf
	sql = sql & " 		and d.itemoption = si.orgitemoption" & vbcrlf
	sql = sql & " 	left join db_item.[dbo].[tbl_item_standing_order] so with (nolock)" & vbcrlf		' ���ⱸ�� ������ȸ�� �Ϸ� �Ǹ� ���� ��ǰ�� ��ҷ� ���ư��� ���� ����. ����
	sql = sql & " 		on d.itemid = so.orgitemid" & vbcrlf
	sql = sql & " 		and d.itemoption = so.orgitemoption" & vbcrlf
	sql = sql & " 		and si.endreserveidx = so.reserveidx" & vbcrlf
	sql = sql & " 	left join db_jungsan.dbo.tbl_nvshop_jungsan_detail JJ with (nolock)" & vbcrlf
	sql = sql & " 		on d.orderserial=JJ.orderserial" & vbcrlf
	sql = sql & " 		and d.itemid=JJ.itemid" & vbcrlf
	sql = sql & " 		and isnull(convert(varchar(10),isNULL(M.canceldate,D.canceldate),21),'')=JJ.cancelDT" & vbcrlf
	sql = sql & " 	where 1=1" & vbcrlf
	sql = sql & " 	and m.ipkumdiv>3" & vbcrlf
	sql = sql & " 	and (m.cancelyn<>'N' or  d.cancelyn='Y')" & vbcrlf	'����� CASE
	sql = sql & " 	and "& jungsandategubun &">@STDT and "& jungsandategubun &"<@EDDT" & vbcrlf
	sql = sql & " 	and d.itemid not in (0,100)" & vbcrlf
	sql = sql & " 	and so.orgitemid is null" & vbcrlf
	sql = sql & " 	and JJ.orderserial is null" & vbcrlf	' �̹� ����Ȱ� ����
	sql = sql & " 	and m.orderserial not in ("
	sql = sql & " 		'21022272847'"
	sql = sql & " 	)"

	'response.write sql & "<br>"
	dbget.CommandTimeout = 60*5   ' 5��
	dbget.execute sql

	sql = "DECLARE @STDT varchar(10)" & vbcrlf
	sql = sql & " DECLARE @EDDT varchar(10)" & vbcrlf
	sql = sql & " SET @STDT='2014-08-01'" & vbcrlf 		'������(�ٲ�������)
	sql = sql & " SET @EDDT='"& dateadd("m", +1, DateSerial(yyyy, mm, "01")) &"'" & vbcrlf
	sql = sql & " insert into db_jungsan.dbo.tbl_nvshop_jungsan_detail" & vbcrlf
	sql = sql & " (orderserial,itemid,itemoption,cancelDT,jMonth,ismobile,orgorderserial,jumundiv,rdsite" & vbcrlf
	sql = sql & " ,rDate,fixedDate,itemNOptionName,itemno,suppPrc,commpro,commissoin,ordStatName)" & vbcrlf
	sql = sql & " 	select" & vbcrlf
	sql = sql & " 	d.orderserial,d.itemid,d.itemoption,''" & vbcrlf
	sql = sql & " 	,convert(varchar(7),"& jungsandategubun &",21) as jMonth" & vbcrlf
	sql = sql & " 	,R.ismobile,isNULL(m.linkorderserial,m.orderserial),m.jumundiv,isNULL(m.rdsite,'')" & vbcrlf
	sql = sql & " 	,convert(varchar(10),m.regdate,21) as rDate" & vbcrlf
	sql = sql & " 	,convert(varchar(10),"& jungsandategubun &",21) as fixedDate" & vbcrlf
	sql = sql & " 	,d.itemname+(CASE WHEN d.itemoptionname<>'' THEN ' ['+d.itemoptionname+']' ELSE '' END)" & vbcrlf
	sql = sql & " 	,d.itemno" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0) as SuppPrc" & vbcrlf
	sql = sql & " 	,(CASE WHEN R.isCharge='Y' THEN 0.03 ELSE 0 END)" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0)*(CASE WHEN R.isCharge='Y' THEN 0.03 ELSE 0 END)" & vbcrlf
	sql = sql & " 	,CASE WHEN m.jumundiv='9' THEN '��ǰ�Ϸ�' WHEN	m.jumundiv='6' THEN '��ȯ�Ϸ�' ELSE '���Ϸ�' END" & vbcrlf
	sql = sql & " 	from db_order.dbo.tbl_order_master m with (nolock)" & vbcrlf
	sql = sql & " 	Join db_order.dbo.tbl_order_detail d with (nolock)" & vbcrlf
	sql = sql & " 		on m.orderserial=d.orderserial" & vbcrlf
	sql = sql & " 	Join db_item.dbo.tbl_Outmall_RdsiteGubun R with (nolock)" & vbcrlf
	sql = sql & " 		on m.rdsite=R.rdsite" & vbcrlf
	sql = sql & " 		and R.gubun in ('"& rdsite &"')" & vbcrlf
	sql = sql & " 	left join db_jungsan.dbo.tbl_nvshop_jungsan_detail J with (nolock)" & vbcrlf
	sql = sql & " 		on d.orderserial=J.orderserial" & vbcrlf
	sql = sql & " 		and d.itemid=j.itemid" & vbcrlf
	sql = sql & " 		and d.itemoption=j.itemoption" & vbcrlf
	sql = sql & " 	where 1=1" & vbcrlf
	sql = sql & " 	and (J.itemoption is NULL" & vbcrlf
	sql = sql & " 	and m.ipkumdiv>3" & vbcrlf
	sql = sql & " 	and m.ipkumdiv=8" & vbcrlf
	sql = sql & " 	and m.cancelyn='N'" & vbcrlf
	sql = sql & " 	and "& jungsandategubun &">@STDT and "& jungsandategubun &"<@EDDT" & vbcrlf
	sql = sql & " 	and d.itemid not in (0,100)" & vbcrlf
	sql = sql & " 	and d.cancelyn<>'Y')" & vbcrlf   '��Ҵ� ���� ����.
	sql = sql & " 	or (m.orderserial = '21022272847' and J.itemoption is NULL)" & vbcrlf   '�ӽ�

	'response.write sql & "<br>"
	dbget.CommandTimeout = 60*5   ' 5��
	dbget.execute sql

	response.write "<script type='text/javascript'>"
	
	if mode="regdaum" then
		response.write "	alert('"& yyyy &"�� "& mm&"�� daum ���������� �ۼ� �Ϸ�.');"
	ELSE
		response.write "	alert('"& yyyy &"�� "& mm&"�� nate ���������� �ۼ� �Ϸ�.');"
	end if

	response.write "	parent.location.reload()"
	response.write "</script>"
	dbget.close() : response.end

'���̹� csv �ٿ�ε�
elseif mode="csvnaver" then
	if ismobile="" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('������ �����ϴ�.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if

	if trim(yyyy="") or trim(mm="") then
		response.write "<script type='text/javascript'>"
		response.write "	alert('�Ⱓ�� �����ϴ�.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if
	yyyy = trim(yyyy)
	mm = trim(mm)

	Set cjungsan = New Ccommission
		cjungsan.FRectyyyymm = yyyy + "-" + mm
		cjungsan.FPageSize = 100000
		cjungsan.FCurrPage = 1
		cjungsan.frectorderserial = orderserial
		cjungsan.frectitemname = itemnoptionname
		cjungsan.frectismobile = ismobile
		cjungsan.Getcommissionjungsan_naver_notpaging()

		if cjungsan.FTotalCount < 1 then
			response.write "<script type='text/javascript'>"
			response.write "	alert('�����Ͱ� �����ϴ�.');"
			response.write "	parent.location.reload()"
			response.write "</script>"
			dbget.close() : response.end
		end if

		arrlist = cjungsan.farrlist

	Set csum = New Ccommission
		csum.FRectyyyymm = yyyy + "-" + mm
		csum.FPageSize = 500
		csum.FCurrPage = 1
'		csum.frectorderserial = orderserial
'		csum.frectitemname = itemnoptionname
		csum.frectismobile = ismobile
		csum.Getcommissionjungsan_naver_sum_notpaging()

		if csum.FTotalCount < 1 then
			response.write "<script type='text/javascript'>"
			response.write "	alert('���� �� �����Ͱ� �����ϴ�.');"
			response.write "	parent.location.reload()"
			response.write "</script>"
			dbget.close() : response.end
		end if

		arrsum = csum.farrlist

'	Response.Buffer=False
	Response.Expires=0
	response.ContentType = "application/vnd.ms-excel"

	Response.AddHeader "Content-Disposition", "attachment; filename=����������naver_"& yyyy &"��"& mm &"��.csv"
	Response.CacheControl = "public"
	response.write "����Ʈ,���Ǹż���,���ֹ��ݾ�,������,���������꿩��,������" & vbcrlf

	if isarray(arrsum) then
		For i = 0 To ubound(arrsum,2)
		bufsumStr = ""
		bufsumStr = bufsumStr & arrsum(0,i)
		bufsumStr = bufsumStr & "," & arrsum(1,i)
		bufsumStr = bufsumStr & "," & arrsum(2,i)
		bufsumStr = bufsumStr & "," & arrsum(3,i)
		bufsumStr = bufsumStr & "," & escapedstring(arrsum(4,i))

		response.write bufsumStr & VbCrlf
		next
	end if

	response.write VbCrlf

	response.write "�ֹ�����,����Ȯ������(�����Ϸ�����),�ֹ���ȣ,��ǰ��,�ֹ�����,�ֹ��ݾ�(V.A.T����),��������,������,�ֹ�����,��ҳ�¥" & vbcrlf

	if isarray(arrlist) then
		For i = 0 To ubound(arrlist,2)
			bufStr = ""
			bufStr = bufStr & arrlist(0,i)
			bufStr = bufStr & "," & arrlist(1,i)
			bufStr = bufStr & "," & arrlist(2,i)
			bufStr = bufStr & "," & escapedstring(arrlist(3,i))
			bufStr = bufStr & "," & arrlist(4,i)
			bufStr = bufStr & "," & arrlist(5,i)
			bufStr = bufStr & "," & arrlist(6,i)
			bufStr = bufStr & "," & arrlist(7,i)
			bufStr = bufStr & "," & arrlist(8,i)
			bufStr = bufStr & "," & arrlist(9,i)

			response.write bufStr & VbCrlf

			if i>0 and (i mod 10000)=0 then response.flush	'���� �ʰ��� �������� �߰� �÷���
		next
	end if

	set cjungsan = nothing

'���̹� �����ۼ�
elseif mode="regnaver" then
	if trim(yyyy="") or trim(mm="") then
		response.write "<script type='text/javascript'>"
		response.write "	alert('�Ⱓ�� �����ϴ�.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if
	yyyy = trim(yyyy)
	mm = trim(mm)

	jungsanCount="0"
	sql = "SELECT count(jd.orderserial) as jungsanCount"
	sql = sql & " from db_jungsan.dbo.tbl_nvshop_jungsan_detail jd with (nolock)" + vbcrlf
	sql = sql & " join db_item.dbo.tbl_Outmall_RdsiteGubun R with (nolock)" + vbcrlf
	sql = sql & " 	on jd.rdsite=R.rdsite" + vbcrlf
	sql = sql & " 	and R.gubun='nvshop'" + vbcrlf
	sql = sql & " where jd.jmonth='"& yyyy + "-" + mm &"'"

	'response.write sql & "<Br>"
	rsget.CursorLocation = adUseClient
	rsget.Open sql, dbget, adOpenForwardOnly, adLockReadOnly

	If Not rsget.Eof Then
		jungsanCount = rsget("jungsancount")
	End If

	rsget.close

	if jungsanCount>0 then
		response.write "<script type='text/javascript'>"
		response.write "	alert('�̹� �ۼ��� ���� �����Ͱ� �ֽ��ϴ�.');"
		response.write "	parent.location.reload()"
		response.write "</script>"
		dbget.close() : response.end
	end if

	' ���� ���� �ٲ�.
	if DateSerial(yyyy, mm, "01") >= "2021-03-01" then
		jungsandategubun="d.jungsanfixdate"
	else
		jungsandategubun="m.beadaldate"
	end if

	'/��� �� ���(���� �ۼ� �� ���) ������ ��ü�� ������ �߸�����ٴ���. ����� ��� ��쵵 ����.
	sql = "DECLARE @STDT varchar(10)" & vbcrlf
	sql = sql & " DECLARE @EDDT varchar(10)" & vbcrlf
	sql = sql & " DECLARE @jMonth varchar(7)" & vbcrlf
	sql = sql & " DECLARE @STORDERSERIAL varchar(11)" & vbcrlf
	sql = sql & " SET @STDT='2014-07-26'" & vbcrlf '���� �ſ�26~ 25�ϱ���
	sql = sql & " SET @EDDT='"& DateSerial(yyyy, mm, "26") &"'" & vbcrlf
	sql = sql & " SET @STORDERSERIAL=RIGHT(REPLACE(@STDT,'-',''),6)+'00000'" & vbcrlf
	sql = sql & " SET @jMonth=LEFT(@EDDT,7)" & vbcrlf
	sql = sql & " insert into db_jungsan.dbo.tbl_nvshop_jungsan_detail" & vbcrlf
	sql = sql & " (orderserial,itemid,itemoption,cancelDT,jMonth,ismobile,orgorderserial,jumundiv,rdsite" & vbcrlf
	sql = sql & " ,rDate,fixedDate,itemNOptionName,itemno,suppPrc,commpro,commissoin,ordStatName)" & vbcrlf
	sql = sql & " 	select" & vbcrlf
	sql = sql & " 	d.orderserial,d.itemid,d.itemoption,isnull(convert(varchar(10),isNULL(M.canceldate,D.canceldate),21),'')" & vbcrlf
	'sql = sql & " 	,convert(varchar(7),"& jungsandategubun &",21) as jMonth" & vbcrlf
	sql = sql & " 	,(case when day("& jungsandategubun &")>25 then convert(varchar(7),dateadd(month,+1,"& jungsandategubun &"),121) else convert(varchar(7),"& jungsandategubun &",121) end) as jMonth" & vbcrlf
	sql = sql & " 	,R.ismobile,isNULL(m.linkorderserial,m.orderserial),m.jumundiv,isNULL(m.rdsite,'')" & vbcrlf
	sql = sql & " 	,convert(varchar(10),m.regdate,21) as rDate" & vbcrlf
	sql = sql & " 	,convert(varchar(10),"& jungsandategubun &",21) as fixedDate" & vbcrlf
	sql = sql & " 	,d.itemname+(CASE WHEN d.itemoptionname<>'' THEN ' ['+d.itemoptionname+']' ELSE '' END)" & vbcrlf
	sql = sql & " 	,d.itemno*-1" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0)*-1 as SuppPrc" & vbcrlf
	sql = sql & " 	,(CASE WHEN R.isCharge='Y' THEN 0.02 ELSE 0 END)" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0)*(CASE WHEN R.isCharge='Y' THEN 0.02 ELSE 0 END)*-1" & vbcrlf
	sql = sql & " 	,CASE" & vbcrlf
	sql = sql & " 		WHEN m.jumundiv='9' THEN '��ǰ�����'" & vbcrlf
	sql = sql & " 		WHEN m.jumundiv='6' THEN '��ȯ�Ϸ�'  ELSE '��������' END" & vbcrlf
	sql = sql & " 	from db_order.dbo.tbl_order_master m with (nolock)" & vbcrlf
	sql = sql & " 	Join db_order.dbo.tbl_order_detail d with (nolock)" & vbcrlf
	sql = sql & " 		on m.orderserial=d.orderserial" & vbcrlf
	sql = sql & " 	Join db_item.dbo.tbl_Outmall_RdsiteGubun R with (nolock)" & vbcrlf
	sql = sql & " 		on m.rdsite=R.rdsite" & vbcrlf
	sql = sql & " 		and R.gubun='nvshop'" & vbcrlf
	sql = sql & " 	Join db_jungsan.dbo.tbl_nvshop_jungsan_detail J with (nolock)" & vbcrlf	'�̹� ���꿡 �� �ְ�.
	sql = sql & " 		on d.orderserial=J.orderserial" & vbcrlf
	sql = sql & " 		and d.itemid=j.itemid" & vbcrlf
	sql = sql & " 		and d.itemoption=j.itemoption and J.ordStatName<>'��������'" & vbcrlf
	sql = sql & " 	left join db_jungsan.dbo.tbl_nvshop_jungsan_detail J2 with (nolock)" & vbcrlf
	sql = sql & " 		on d.orderserial=J2.orderserial" & vbcrlf
	sql = sql & " 		and d.itemid=j2.itemid" & vbcrlf
	sql = sql & " 		and d.itemoption=j2.itemoption and J2.ordStatName='��������'" & vbcrlf
	sql = sql & " 	left join db_item.[dbo].[tbl_item_standing_item] si with (nolock)" & vbcrlf
	sql = sql & " 		on d.itemid = si.orgitemid" & vbcrlf
	sql = sql & " 		and d.itemoption = si.orgitemoption" & vbcrlf
	sql = sql & " 	left join db_item.[dbo].[tbl_item_standing_order] so with (nolock)" & vbcrlf		' ���ⱸ�� ������ȸ�� �Ϸ� �Ǹ� ���� ��ǰ�� ��ҷ� ���ư��� ���� ����. ����
	sql = sql & " 		on d.itemid = so.orgitemid" & vbcrlf
	sql = sql & " 		and d.itemoption = so.orgitemoption" & vbcrlf
	sql = sql & " 		and si.endreserveidx = so.reserveidx" & vbcrlf
	sql = sql & " 	where m.ipkumdiv>3" & vbcrlf
	sql = sql & " 	and (m.cancelyn<>'N' or d.cancelyn='Y')" & vbcrlf 	'����� CASE (�����)
	sql = sql & " 	and "& jungsandategubun &">@STDT and "& jungsandategubun &"<@EDDT" & vbcrlf
	sql = sql & " 	and d.itemid not in (0,100)" & vbcrlf
	sql = sql & " 	and J2.orderserial is NULL" & vbcrlf
	sql = sql & " 	and so.orgitemid is null" & vbcrlf

	'response.write sql & "<br>"
	dbget.CommandTimeout = 60*5   ' 5��
	dbget.execute sql

	sql = "DECLARE @STDT varchar(10)" & vbcrlf
	sql = sql & " DECLARE @EDDT varchar(10)" & vbcrlf
	sql = sql & " DECLARE @jMonth varchar(7)" & vbcrlf
	sql = sql & " DECLARE @STORDERSERIAL varchar(11)" & vbcrlf
	sql = sql & " SET @STDT='2014-05-26'" & vbcrlf '���� �ſ�26~ 25�ϱ���
	sql = sql & " SET @EDDT='"& DateSerial(yyyy, mm, "26") &"'" & vbcrlf
	sql = sql & " SET @STORDERSERIAL=RIGHT(REPLACE(@STDT,'-',''),6)+'00000'" & vbcrlf
	sql = sql & " SET @jMonth=LEFT(@EDDT,7)" & vbcrlf
	sql = sql & " insert into db_jungsan.dbo.tbl_nvshop_jungsan_detail" & vbcrlf
	sql = sql & " (orderserial,itemid,itemoption,cancelDT,jMonth,ismobile,orgorderserial,jumundiv,rdsite" & vbcrlf
	sql = sql & " ,rDate,fixedDate,itemNOptionName,itemno,suppPrc,commpro,commissoin,ordStatName)" & vbcrlf
	sql = sql & " 	select" & vbcrlf
	sql = sql & " 	d.orderserial,d.itemid,d.itemoption,''" & vbcrlf
	'sql = sql & " 	,convert(varchar(7),"& jungsandategubun &",21) as jMonth" & vbcrlf
	sql = sql & " 	,(case when day("& jungsandategubun &")>25 then convert(varchar(7),dateadd(month,+1,"& jungsandategubun &"),121) else convert(varchar(7),"& jungsandategubun &",121) end) as jMonth" & vbcrlf
	sql = sql & " 	,R.ismobile,isNULL(m.linkorderserial,m.orderserial),m.jumundiv,isNULL(m.rdsite,'')" & vbcrlf
	sql = sql & " 	,convert(varchar(10),m.regdate,21) as rDate" & vbcrlf
	sql = sql & " 	,convert(varchar(10),"& jungsandategubun &",21) as fixedDate" & vbcrlf
	sql = sql & " 	,d.itemname+(CASE WHEN d.itemoptionname<>'' THEN ' ['+d.itemoptionname+']' ELSE '' END)" & vbcrlf
	sql = sql & " 	,d.itemno" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0) as SuppPrc" & vbcrlf
	sql = sql & " 	,(CASE WHEN R.isCharge='Y' THEN 0.02 ELSE 0 END)" & vbcrlf
	sql = sql & " 	,Round((d.reducedprice*d.itemno)/(CASE WHEN d.vatinclude='N' THEN 1 ELSE 1.1 END),0)*(CASE WHEN R.isCharge='Y' THEN 0.02 ELSE 0 END)" & vbcrlf
	sql = sql & " 	,CASE" & vbcrlf
	sql = sql & " 		WHEN m.jumundiv='9' THEN '��ǰ�Ϸ�'" & vbcrlf
	sql = sql & " 		WHEN	m.jumundiv='6' THEN '��ȯ�Ϸ�' ELSE '���Ϸ�' END" & vbcrlf
	sql = sql & " 	from db_order.dbo.tbl_order_master m with (nolock)" & vbcrlf
	sql = sql & " 	Join db_order.dbo.tbl_order_detail d with (nolock)" & vbcrlf
	sql = sql & " 		on m.orderserial=d.orderserial" & vbcrlf
	sql = sql & " 	Join db_item.dbo.tbl_Outmall_RdsiteGubun R with (nolock)" & vbcrlf
	sql = sql & " 		on m.rdsite=R.rdsite" & vbcrlf
	sql = sql & " 		and R.gubun='nvshop'" & vbcrlf
	sql = sql & " 	left join db_jungsan.dbo.tbl_nvshop_jungsan_detail J with (nolock)" & vbcrlf
	sql = sql & " 		on d.orderserial=J.orderserial" & vbcrlf
	sql = sql & " 		and d.itemid=j.itemid" & vbcrlf
	sql = sql & " 		and d.itemoption=j.itemoption" & vbcrlf
	sql = sql & " 	where J.itemoption is NULL" & vbcrlf
	sql = sql & " 	and m.ipkumdiv>3" & vbcrlf
	sql = sql & " 	and m.ipkumdiv=8" & vbcrlf
	sql = sql & " 	and m.cancelyn='N'" & vbcrlf
	sql = sql & " 	and "& jungsandategubun &">@STDT and "& jungsandategubun &"<@EDDT" & vbcrlf
	sql = sql & " 	and d.itemid not in (0,100)" & vbcrlf
	sql = sql & " 	and d.cancelyn<>'Y'" & vbcrlf   '��Ҵ� ���� ����.

	'response.write sql & "<br>"
	dbget.CommandTimeout = 60*5   ' 5��
	dbget.execute sql

	response.write "<script type='text/javascript'>"
	response.write "	alert('"& yyyy &"�� "& mm&"�� naver ���������� �ۼ� �Ϸ�.');"
	response.write "	parent.location.reload()"
	response.write "</script>"
	dbget.close() : response.end

else
	response.write "<script type='text/javascript'>"
	response.write "	alert('�߸��� ��� �Դϴ�.');"
	response.write "	parent.location.reload()"
	response.write "</script>"
	dbget.close() : response.end
end if
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->