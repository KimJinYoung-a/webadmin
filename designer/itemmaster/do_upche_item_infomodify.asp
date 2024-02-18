<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim itemid, largeno, midno, smallno, itemdiv
dim keywords, sourcearea, makername, itemsource
dim itemsize, usinghtml, itemcontent, ordercomment, requireMakeDay
dim mode, itemoption, isusing, upchemanagecode
dim infoDiv, safetyYn, safetyDiv, safetyNum
dim freight_min, freight_max
dim 	itemweight,deliverarea,deliverOverseas
'// �Ǹ� ������
dim sellEndDate
sellEndDate= requestCheckVar(request("sellEndDate"),20)

On Error Resume Next
if (sellEndDate<>"") then
sellEndDate = CDate(sellEndDate)
If Err then sellEndDate=""
end if
On Error Goto 0

itemid = requestCheckVar(request("itemid"),10)
largeno = requestCheckVar(request("cd1"),10)
midno	= requestCheckVar(request("cd2"),10)
smallno	= requestCheckVar(request("cd3"),10)
itemdiv	= requestCheckVar(request("itemdiv"),10)
keywords	= html2db(request("keywords"))
sourcearea	= requestCheckVar(html2db(request("sourcearea")),300)
makername	= requestCheckVar(html2db(request("makername")),200)
itemsource	= requestCheckVar(html2db(request("itemsource")),300)
itemsize	= requestCheckVar(html2db(request("itemsize")),300)
usinghtml	= request("usinghtml")
itemcontent	= html2db(request("itemcontent"))
ordercomment	= html2db(request("ordercomment"))
requireMakeDay	= html2db(request("requireMakeDay"))
upchemanagecode = html2db(request("upchemanagecode"))
infoDiv = requestCheckVar(request("infoDiv"),10)
safetyYn = requestCheckVar(request("safetyYn"),10)
safetyDiv = requestCheckVar(request("safetyDiv"),20)
safetyNum = html2db(request("safetyNum"))
itemweight =  requestCheckvar(Request("itemWeight"),10)
deliverarea =  requestCheckvar(Request("deliverarea"),1)
deliverOverseas=  requestCheckvar(Request("deliverOverseas"),1)
if deliverOverseas = "" then deliverOverseas ="N"
if itemweight = "" then itemweight = 0
freight_min = getNumeric(request("freight_min"))
freight_max = getNumeric(request("freight_max"))
if freight_min="" then freight_min="0"
if freight_max="" then freight_max="0"

dim sqlStr,i
dim AssignedRow

'==============================================================================
sqlStr = "update [db_item].[dbo].tbl_item" & VbCrlf
sqlStr = sqlStr & " set cate_large='" & largeno & "'" & VbCrlf
sqlStr = sqlStr & " , cate_mid='" & midno & "'" & VbCrlf
sqlStr = sqlStr & " , cate_small='" & smallno & "'" & VbCrlf
sqlStr = sqlStr & " , itemdiv='" & CStr(itemdiv) & "'" & VbCrlf
sqlStr = sqlStr & " , upchemanagecode=convert(varchar(32),'" & upchemanagecode & "')" & vbCrlf

IF (sellEndDate<>"") then
    sqlStr = sqlStr & " ,sellEndDate='" & CStr(sellEndDate) & " 23:59:59" & "'" & vbCrlf 
ELSE
    sqlStr = sqlStr & " ,sellEndDate=NULL"  & vbCrlf 
End IF
sqlStr = sqlStr & " ,lastupdate=getdate()" & vbCrlf 
sqlStr = sqlStr & ", itemweight = "&itemweight&vbCrlf
sqlStr = sqlStr & ", deliverarea = '"&deliverarea&"'"&vbCrlf
sqlStr = sqlStr & ", deliverOverseas = '"&deliverOverseas&"'"&vbCrlf
sqlStr = sqlStr & " where itemid=" & CStr(itemid) & " "
sqlStr = sqlStr & " and makerid='" & CStr(session("ssBctID")) & "' "

dbget.Execute sqlStr, AssignedRow

if (AssignedRow>0) then
	'// ī�װ� �ߺ� Ȯ��(2008.07.31; ������)
	sqlStr = "select count(*) from db_item.dbo.tbl_item_category where itemid=" & itemid &_
			"	and code_large='" & largeno & "' " &_
			"	and code_mid='" & midno & "' " &_
			"	and code_small='" & smallno & "' and code_div='A' "
	rsget.Open sqlStr ,dbget,1

	if rsget(0)<1 then
	    sqlStr = "update [db_item].[dbo].tbl_item_Contents" + VbCrlf
	    sqlStr = sqlStr + " set keywords='" + keywords + "'" + VbCrlf
	    sqlStr = sqlStr + " , sourcearea='" + sourcearea + "'" + VbCrlf
	    sqlStr = sqlStr + " , makername='" + makername + "'" + VbCrlf
	    sqlStr = sqlStr + " , itemsource='" + itemsource + "'" + VbCrlf
	    sqlStr = sqlStr + " , itemsize='" + itemsize + "'" + VbCrlf
	    sqlStr = sqlStr + " , usinghtml='" + usinghtml + "'" + VbCrlf
	    sqlStr = sqlStr + " , itemcontent='" + itemcontent + "'" + VbCrlf
	    sqlStr = sqlStr + " , ordercomment='" + ordercomment + "'" + VbCrlf
	    sqlStr = sqlStr + " , requireMakeDay='" + requireMakeDay + "'" + VbCrlf
	    sqlStr = sqlStr + " , infoDiv='" + infoDiv + "'" + VbCrlf
	    sqlStr = sqlStr + " , safetyYn='" + safetyYn + "'" + VbCrlf
	    sqlStr = sqlStr + " , safetyDiv='" + safetyDiv + "'" + VbCrlf
	    sqlStr = sqlStr + " , safetyNum='" + safetyNum + "'" + VbCrlf
	    sqlStr = sqlStr + " , freight_min='" + freight_min + "'" + VbCrlf
	    sqlStr = sqlStr + " , freight_max='" + freight_max + "'" + VbCrlf

	    sqlStr = sqlStr + " where itemid=" + CStr(itemid) + " "

	    dbget.Execute sqlStr


	    '''�� ī�װ� : ��ü�� �⺻ ī�װ��� ����
	    sqlStr = "update [db_item].dbo.tbl_Item_category " 
	    sqlStr = sqlStr + " set code_large='" + largeno + "'"
	    sqlStr = sqlStr + " , code_mid='" + midno + "'"
	    sqlStr = sqlStr + " , code_small='" + smallno + "'"
	    sqlStr = sqlStr + " where itemid=" & CStr(itemid)
	    sqlStr = sqlStr + " and code_div='D'"
	    sqlStr = sqlStr + " and ("
	    sqlStr = sqlStr + "         code_large<>'" + largeno + "'"
	    sqlStr = sqlStr + "     or  code_mid<>'" + midno + "'"
	    sqlStr = sqlStr + "     or  code_small<>'" + smallno + "'"
	    sqlStr = sqlStr + " )"

	    dbget.Execute sqlStr

		'//��ǰ ǰ�������� ����
		if Request("infoDiv")<>"" then
			dim infoCd, infoCont, infoChk

			'�迭�� ó��
			redim infoCd(Request("infoCd").Count)
			redim infoCont(Request("infoCont").Count)
			redim infoChk(Request("infoChk").Count)
			for i=1 to Request("infoCd").Count
				infoCd(i) = Request("infoCd")(i)
				infoCont(i) = Request("infoCont")(i)
				infoChk(i) = Request("infoChk")(i)
			next

			'������ ����
			sqlStr = "Delete From db_item.dbo.tbl_item_infoCont Where itemid='" & CStr(itemid) & "'"
			dbget.execute(sqlStr)

			'DB�� ó��
			On Error Resume Next
			for i=1 to ubound(infoCd)
				'�Է°��� �ִ� ��츸 ����
				if infoChk(i)<>"" or infoCont(i)<>"" then
					sqlStr = "Insert into db_item.dbo.tbl_item_infoCont (itemid, infoCd, chkDiv, infoContent) values "
					sqlStr = sqlStr & "('" & CStr(itemid) & "'"
					sqlStr = sqlStr & ",'" & CStr(infoCd(i)) & "'"
					sqlStr = sqlStr & ",'" & CStr(infoChk(i)) & "'"
					sqlStr = sqlStr & ",'" & html2db(infoCont(i)) & "')"
					dbget.execute(sqlStr)
				end if
			next
			If Err.Number<>0 Then
				Response.Write "<script language=javascript>alert('��ǰǰ������ ó���� ������ �߻��߽��ϴ�.\n�Է� ������ �ٽ� �ѹ� Ȯ�����ּ���.');history.back();</script>"
				dbget.close()	:	response.End
			end if
			On Error Goto 0
		end if

	else
		Response.Write "<script language=javascript>alert('�̹� ��ǰ�� �����Ǿ��ִ� ī�װ��� �����Ͽ����ϴ�.\n\n���߰� ī�װ��� �����Ǿ����� ��찡 �����Ƿ� ���MD���� Ȯ��/������û�� ���ּ���.');history.back();</script>"
		dbget.close()	:	response.End
	end if

	rsget.Close
end if

					
dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>

<script language="javascript">
alert('���� �Ǿ����ϴ�.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->