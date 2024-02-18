<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doJust1Day_Process.asp
' Discription : ����Ʈ ������ ó�� ������
' History : 2008.04.09 ������ ����
'				2014.09.12 ������ ��������� ���� �ɿ��� ����
'###############################################

'// ���� ���� �� �Ķ���� ����
dim menupos, mode, sqlStr, lp
dim justDate, sale_code, itemid, salePrice, orgPrice, saleSuplyCash, limitNo, limitYn, justDesc, img1, img2, img3, img4
Dim orgsailprice, orgsailsuplycash, orgsailyn

menupos		= Request("menupos")
mode		= Request("mode")

justDate	= Request("justDate")
sale_code	= getNumeric(Request("sale_code"))
itemid		= getNumeric(Request("itemid"))
salePrice	= getNumeric(Request("salePrice"))
orgPrice	= getNumeric(Request("orgPrice"))
saleSuplyCash = getNumeric(Request("saleSuplyCash"))
limitNo		= getNumeric(Request("limitNo"))
limitYn		= Request("limitYn")
justDesc	= html2db(Request("justDesc"))
img1		= Request("image1")
img2		= Request("image2")
img3		= Request("image3")

If instr(img1, "http://webimage.10x10.co.kr/jikbang_just1day") = 0 Then
	img1 = "http://webimage.10x10.co.kr/jikbang_just1day/"&img1
End If

If instr(img2, "http://webimage.10x10.co.kr/jikbang_just1day") = 0 Then
	img2 = "http://webimage.10x10.co.kr/jikbang_just1day/"&img2
End If

If instr(img3, "http://webimage.10x10.co.kr/jikbang_just1day") = 0 Then
	img3 = "http://webimage.10x10.co.kr/jikbang_just1day/"&img3
End If
'// Ʈ������ ����
dbget.beginTrans

'// ��忡 ���� �б�
Select Case mode
	Case "add"
		'// �ű� ���
		rsget.Open "Select count(JustDate) from [db_etcmall].[dbo].tbl_jikbang_oneDay where JustDate='" & justDate & "'", dbget, 1
		if rsget(0)>0 then
			Alert_return("�̹� ��ϵ� ��¥�Դϴ�.\n�ٸ� ��¥�� �������ּ���.")
			dbget.close()	:	response.End
		end if
		rsget.Close

		'// ���ϵ� �����ݰ� ���Ͽ��θ� ������
		rsget.Open " Select top 1 sailprice, sailsuplycash, sailyn From db_item.dbo.tbl_item Where itemid='"&itemid&"' "
		If Not(rsget.bof Or rsget.eof) Then
			orgsailprice = rsget("sailprice")
			orgsailsuplycash = rsget("sailsuplycash")
			orgsailyn = rsget("sailyn")
		End If
		rsget.Close


		'' ���ΰ� 0, ���԰�0 �ΰ�� ���� �ȵ�. �ϴ� ���
		'���ο��� ���̺� ����(������)
		sqlStr = "Insert Into [db_event].[dbo].tbl_sale " &_
				" (sale_name, sale_rate, sale_margin, sale_marginvalue, sale_startdate, sale_enddate, availPayType, adminid, sale_status) values " &_
				" ('����_" & justDate & "' " &_
				" ," & cInt((1-cLng(salePrice)/cLng(orgPrice))*100) &_
				" , 5, " & cInt((1-cLng(salePrice)/cLng(orgPrice))*100) &_
				" ,'" & justDate & "', '" & justDate & "'" &_
				" , '8', '" & session("ssBctId") & "',7)"
		dbget.Execute(sqlStr)
		sqlStr = "select IDENT_CURRENT('[db_event].[dbo].tbl_sale') as sale_code"
		rsget.Open sqlStr, dbget, 1
		If Not rsget.Eof then
			sale_code = rsget("sale_code")
		end if
		rsget.close

		'���ο��� ���̺� ����(��ǰ����)
'		sqlStr = "Insert Into [db_event].[dbo].tbl_saleItem " &_
'				" (sale_code, itemid, saleprice, salesupplyCash, limitno, orgsailprice, orgsailsuplycash, orgsailyn, orglimityn, saleItem_status) values " &_
		sqlStr = "Insert Into [db_event].[dbo].tbl_saleItem " &_
				" (sale_code, itemid, saleprice, salesupplyCash, limitno, orglimityn, saleItem_status) values " &_
				" (" & sale_code &_
				" ," & itemid &_
				" ," & salePrice &_
				" ," & SaleSuplyCash &_
				" ," & limitNo &_
				" ,'" & limitYn & "', 7)"
		dbget.Execute(sqlStr)
    
		'����Ʈ ������ ����
		sqlStr = "Insert Into [db_etcmall].[dbo].tbl_jikbang_oneDay " &_
				" (JustDate,itemid,orgPrice,justSalePrice,SaleSuplyCash,justDesc,sale_code,limitNo,adminid,OutPutImgUrl1,OutPutImgUrl2,contentImgUrl) values " &_
				" ('" & justDate & "'" &_
				" ," & itemid &_
				" ," & orgPrice &_
				" ," & salePrice &_
				" ," & SaleSuplyCash &_
				" ,'" & justDesc & "'" &_
				" ," & sale_code &_
				" ," & limitNo &_
				" ,'" & session("ssBctId") & "'" &_
				" ,'" & img1 & "','" & img2 & "','" & img3 & "')"
				
		dbget.Execute(sqlStr)

	Case "edit"
		'// ���� ����
		sqlStr = "Update [db_etcmall].[dbo].tbl_jikbang_oneDay SET " &_
				" 	OutPutImgUrl1=''" &_
				" 	,OutPutImgUrl2=''" &_
				" 	,contentImgUrl=''" &_
				" Where justDate='" & justDate & "'"
		dbget.Execute(sqlStr)

		sqlStr = "Update [db_etcmall].[dbo].tbl_jikbang_oneDay " &_
				" Set justSalePrice=" & salePrice &_
				" 	,SaleSuplyCash=" & SaleSuplyCash &_
				" 	,limitNo=" & limitNo &_
				" 	,justDesc='" & justDesc & "'" &_
				" 	,OutPutImgUrl1='" & img1 & "'" &_
				" 	,OutPutImgUrl2='" & img2 & "'" &_
				" 	,contentImgUrl='" & img3 & "'" &_
				" Where justDate='" & justDate & "'"
		dbget.Execute(sqlStr)
        
        if sale_code<>"" then
    		sqlStr = "Update [db_event].[dbo].tbl_sale " &_
    				" Set sale_rate=" & cInt((1-cLng(salePrice)/cLng(orgPrice))*100) &_
    				" 	,sale_marginvalue=" & cInt((1-cLng(salePrice)/cLng(orgPrice))*100) &_
    				" Where sale_code=" & sale_code
    		dbget.Execute(sqlStr)
        
        
    		sqlStr = "Update [db_event].[dbo].tbl_saleItem " &_
    				" Set saleprice=" & saleprice &_
    				" 	,salesupplyCash=" & SaleSuplyCash &_
    				" 	,limitno=" & limitno &_
    				"	,lastupdate=getdate() " &_
    				" Where sale_code=" & sale_code & " and itemid=" & itemid
    		dbget.Execute(sqlStr)
        end if
	Case "delete"
		'// ����
		if justDate>cStr(date()) then
			if sale_code<>"" then
				'���� ���ο��� ���� ����
				sqlStr = "Update [db_event].[dbo].tbl_sale " &_
						" Set sale_using=0 " &_
						" Where sale_code=" & sale_code & ";" & vbCrLf
			end if
			'����Ʈ������ ���� ����
			sqlStr = sqlStr & "delete [db_etcmall].[dbo].tbl_jikbang_oneDay " &_
					" Where justDate='" & justDate & "';" & vbCrLf
			dbget.Execute(sqlStr)
		else
			Alert_return("���� �������̰ų� �Ϸ�� ��ǰ�� ������ �� �����ϴ�.")
			response.End
		end if

End Select


'// Ʈ������ �˻� �� ����
If Err.Number = 0 Then
        dbget.CommitTrans
Else
        dbget.RollBackTrans
		Alert_return("����Ÿ�� �����ϴ� ���߿� ������ �߻��Ͽ����ϴ�.")
		dbget.close()	:	response.End
End If

%>
<script language="javascript">
<!--
	// ������� ����
	alert("�����߽��ϴ�.");
	self.location = "Just1Day_list.asp?menupos=<%=menupos%>";
//-->
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->