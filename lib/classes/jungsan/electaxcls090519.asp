<%
class CElecTaxRegItem
	public Fidx

	public Fjungsanid
	public Fjungsanname
	public Fjungsangubun
	public Fmakerid

	public Fapi_no              '����/���/��ȸ         "1/2/3"
	public Funiq_id             '������ȣ               PK
	public Fbiz_no              '���� ����ڹ�ȣ
	public Fcorp_nm             '���� ��ȣ
	public Fceo_nm              '���� ��ǥ�ڸ�
	public Fbiz_status          '���� ����
	public Fbiz_type            '���� ����
	public Faddr                '���� �ּ�
	public Fdam_nm              '���� ����ڸ�
	public Femail               '���� �̸���
	public Fhp_no1              '���� �ڵ��� 1
	public Fhp_no2              '���� �ڵ��� 2
	public Fhp_no3              '���� �ڵ��� 3


	public Fwrite_date          '���ݰ�꼭������
	public Fsb_type             '����/����              "01/02"
	public Ftax_type            '����/���/����         "01/02/03"
	public Fbill_type           '����/û��              "01/18"
	public Fpc_gbn              '����/���              "P/C"
	public Fvol_no
	public Fissue_no
	public Fserial_no           '��꼭 �Ϸù�ȣ
	public Fremark

	public Fitem_count          'ǰ�񰹼�               �ִ�4��
	public Fitem_nm             'ǰ���                 "ǰ��1|ǰ��2|ǰ��3"
	public Fitem_std
	public Fitem_qty            'ǰ�����               "1|2|3"
	public Fitem_price          'ǰ����ް�             "1000|2000|3000"
	public Fapprove_type   		''01���޹޴��ڰ� ����   11�����ڰ�����
	public Fitem_amt
	public Fitem_vat
	public Fitem_remark


	public Fcur_c_corp_no '10001568
	public Fcur_u_user_no '1000394
	public Fcur_biz_no '2118700620
	public Fcur_corp_nm '(��)�ٹ�����
	public Fcur_ceo_nm '��â��
	public Fcur_biz_status '����
	public Fcur_biz_type '���ڻ�ŷ�
	public Fcur_addr '����� ���α� ������ 1-45 ��������2��

	public Fcur_dam_nm '�̹���
	public Fcur_email 'moon@10x10.co.kr
	public Fcur_hp_no1 '017
	public Fcur_hp_no2 '360
	public Fcur_hp_no3 '6991

	public Fcash_amt
	public Fcredit_amt

	Private Sub Class_Initialize()
		Fapi_no = "1"
		Fapprove_type = "01"

		Fcur_c_corp_no = "57911"   ''10001568
		Fcur_u_user_no = "244730"  ''1000394  '' 261746 (customer)

		Fcur_biz_no = "2118700620"
		Fcur_corp_nm = "(��)�ٹ�����"
		Fcur_ceo_nm = "��â��"
		Fcur_biz_status = "����"
		Fcur_biz_type = "���ڻ�ŷ�"
		Fcur_addr = "����� ���α� ������ 1-45 ��������2��"
		Fcur_dam_nm = "�̹���"
		Fcur_email = "moon@10x10.co.kr"
		Fcur_hp_no1 = "017"
		Fcur_hp_no2 = "360"
		Fcur_hp_no3 = "6991"
	end sub

	Private Sub Class_Terminate()

	End Sub
end Class

class CElecTaxReg
	public FRectOneRegitem
	public Ftax_no
	public Fresultmsg
	public FResultMsgALL

	public sub ExecDTIXmlDom()
		dim objXMLHTTP, reqParam
		dim tmpArr
		dim sqlStr


		reqParam = "uniq_id=" + FRectOneRegitem.Funiq_id
		''//���޹޴���---------------------------------------
		reqParam = reqParam + "&biz_no=" + FRectOneRegitem.Fbiz_no
		reqParam = reqParam + "&corp_nm=" + Server.URLEncode(FRectOneRegitem.Fcorp_nm)
		reqParam = reqParam + "&ceo_nm=" + Server.URLEncode(FRectOneRegitem.Fceo_nm)
		reqParam = reqParam + "&biz_status=" + Server.URLEncode(FRectOneRegitem.Fbiz_status)
		reqParam = reqParam + "&biz_type=" + Server.URLEncode(FRectOneRegitem.Fbiz_type)
		reqParam = reqParam + "&addr=" + Server.URLEncode(FRectOneRegitem.Faddr)
		reqParam = reqParam + "&dam_nm=" + Server.URLEncode(FRectOneRegitem.Fdam_nm)
		reqParam = reqParam + "&email=" + Server.URLEncode(FRectOneRegitem.Femail)
		reqParam = reqParam + "&hp_no1=" + Left(FRectOneRegitem.Fhp_no1,3)
		reqParam = reqParam + "&hp_no2=" + Left(FRectOneRegitem.Fhp_no2,4)
		reqParam = reqParam + "&hp_no3=" + Left(FRectOneRegitem.Fhp_no3,4)
		''//-------------------------------------------------
		reqParam = reqParam + "&write_date=" + replace(replace(FRectOneRegitem.Fwrite_date,"-",""),"/","")
		reqParam = reqParam + "&sb_type=" + FRectOneRegitem.Fsb_type
		reqParam = reqParam + "&tax_type=" + FRectOneRegitem.Ftax_type
		reqParam = reqParam + "&bill_type=" + FRectOneRegitem.Fbill_type
		reqParam = reqParam + "&pc_gbn=" + FRectOneRegitem.Fpc_gbn
		reqParam = reqParam + "&serial_no=" + FRectOneRegitem.Fserial_no
		reqParam = reqParam + "&item_count=" + FRectOneRegitem.Fitem_count
		reqParam = reqParam + "&item_nm=" + Server.URLEncode(FRectOneRegitem.Fitem_nm)
		reqParam = reqParam + "&item_qty=" + FRectOneRegitem.Fitem_qty
		reqParam = reqParam + "&item_price=" + FRectOneRegitem.Fitem_price
		reqParam = reqParam + "&item_amt=" + FRectOneRegitem.Fitem_amt
		reqParam = reqParam + "&item_vat=" + FRectOneRegitem.Fitem_vat
		reqParam = reqParam + "&item_remark=" + Server.URLEncode(FRectOneRegitem.Fitem_remark)

		reqParam = reqParam + "&approve_type=" + FRectOneRegitem.Fapprove_type

		reqParam = reqParam + "&cur_c_corp_no=" + FRectOneRegitem.Fcur_c_corp_no
		reqParam = reqParam + "&cur_u_user_no=" + FRectOneRegitem.Fcur_u_user_no
		reqParam = reqParam + "&cur_biz_no=" + FRectOneRegitem.Fcur_biz_no
		reqParam = reqParam + "&cur_corp_nm=" + Server.URLEncode(FRectOneRegitem.Fcur_corp_nm)
		reqParam = reqParam + "&cur_ceo_nm=" + Server.URLEncode(FRectOneRegitem.Fcur_ceo_nm)
		reqParam = reqParam + "&cur_biz_status=" + Server.URLEncode(FRectOneRegitem.Fcur_biz_status)
		reqParam = reqParam + "&cur_biz_type=" + Server.URLEncode(FRectOneRegitem.Fcur_biz_type)
		reqParam = reqParam + "&cur_addr=" + Server.URLEncode(FRectOneRegitem.Fcur_addr)

		reqParam = reqParam + "&cur_dam_nm=" + Server.URLEncode(FRectOneRegitem.Fcur_dam_nm)
		reqParam = reqParam + "&cur_email=" + Server.URLEncode(FRectOneRegitem.Fcur_email)
		reqParam = reqParam + "&cur_hp_no1=" + FRectOneRegitem.Fcur_hp_no1
		reqParam = reqParam + "&cur_hp_no2=" + FRectOneRegitem.Fcur_hp_no2
		reqParam = reqParam + "&cur_hp_no3=" + FRectOneRegitem.Fcur_hp_no3
		'reqParam = reqParam + "&cash_amt=0"
		reqParam = reqParam + "&credit_amt=" + FRectOneRegitem.Fcredit_amt
		reqParam = reqParam + "&enc_yn=N"
		reqParam = reqParam + "&final_status=12"




'response.write "<!--" & reqParam & "-->"
'response.write "�˼��մϴ�. ��� �������Դϴ�."
'dbget.close()	:	response.End

		Set objXMLHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP")
		objXMLHTTP.Open "POST",	"http://web1.neoport.net:8383/tx_create.req", False
		''objXMLHTTP.Open "POST",	"http://api.neoport.net/tx_create.req", False  ''80Port
		
		objXMLHTTP.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		objXMLHTTP.Send reqParam

		FResultMsgALL = trim(objXMLHTTP.responseText)
		FResultMsgALL = replace(FResultMsgALL,Vbcrlf,"")
		FResultMsgALL = replace(FResultMsgALL,Vbcr,"")
		FResultMsgALL = replace(FResultMsgALL,Vblf,"")
		Set objXMLHTTP = Nothing

		'response.write FResultMsgALL

		If FResultMsgALL <> "" Then
		    tmpArr = Split(FResultMsgALL, "|")

		    if UBound(tmpArr)>=0 then
				Ftax_no = trim(Left(tmpArr(0),32))
			end if

			if UBound(tmpArr)>=1 then
				Fresultmsg = trim(Left(tmpArr(1),128))
			end if

		    sqlStr = " update [db_jungsan].[dbo].tbl_tax_history_master" + vbCrlf
			sqlStr = sqlStr + " set tax_no='" + Ftax_no + "'" + vbCrlf
			sqlStr = sqlStr + " , resultmsg='" + Fresultmsg + "'" + vbCrlf
			sqlStr = sqlStr + " where idx=" + CStr(FRectOneRegitem.Fidx) + vbCrlf

			rsget.Open sqlStr,dbget,1

			if Fresultmsg="OK" then
				if FRectOneRegitem.Fjungsangubun="ON" then
					sqlStr = " update [db_jungsan].[dbo].tbl_designer_jungsan_master" + vbCrlf
					sqlStr = sqlStr + " set taxlinkidx=" + CStr(FRectOneRegitem.Fidx) + vbCrlf
					sqlStr = sqlStr + " ,neotaxno='" + CStr(Ftax_no) + "'" + vbCrlf
					sqlStr = sqlStr + " ,finishflag='3'"  + vbCrlf
					sqlStr = sqlStr + " ,taxinputdate=getdate()"  + vbCrlf
					sqlStr = sqlStr + " ,taxregdate='" + FRectOneRegitem.Fwrite_date + "'"  + vbCrlf
					sqlStr = sqlStr + " where id=" + CStr(FRectOneRegitem.Fjungsanid)

					rsget.Open sqlStr,dbget,1
				elseif (FRectOneRegitem.Fjungsangubun="OFF") or (FRectOneRegitem.Fjungsangubun="FRN") then
				''������.
					sqlStr = " update [db_shop].[dbo].tbl_shop_jungsanmaster" + vbCrlf
					sqlStr = sqlStr + " set taxlinkidx=" + CStr(FRectOneRegitem.Fidx) + vbCrlf
					sqlStr = sqlStr + " ,neotaxno='" + CStr(Ftax_no) + "'" + vbCrlf
					sqlStr = sqlStr + " ,currstate='3'"  + vbCrlf
					sqlStr = sqlStr + " ,taxregdate=getdate()"  + vbCrlf
					sqlStr = sqlStr + " ,segumil='" + FRectOneRegitem.Fwrite_date + "'"  + vbCrlf
					sqlStr = sqlStr + " where idx=" + CStr(FRectOneRegitem.Fjungsanid)

					rsget.Open sqlStr,dbget,1
			    elseif (FRectOneRegitem.Fjungsangubun="OF") then
					sqlStr = " update [db_jungsan].[dbo].tbl_off_jungsan_master" + vbCrlf
					sqlStr = sqlStr + " set taxlinkidx=" + CStr(FRectOneRegitem.Fidx) + vbCrlf
					sqlStr = sqlStr + " ,neotaxno='" + CStr(Ftax_no) + "'" + vbCrlf
					sqlStr = sqlStr + " ,finishflag='3'"  + vbCrlf
					sqlStr = sqlStr + " ,taxinputdate=getdate()"  + vbCrlf
					sqlStr = sqlStr + " ,taxregdate='" + FRectOneRegitem.Fwrite_date + "'"  + vbCrlf
					sqlStr = sqlStr + " where idx=" + CStr(FRectOneRegitem.Fjungsanid)

					rsget.Open sqlStr,dbget,1
				end if
			end if
		End If
	end sub

	public sub ExecDTI()
		dim dtiObj, reqParam
		dim tmpArr
		dim sqlStr

		Set dtiObj = Server.CreateObject("NeoportDtiX.NeoportDti")
		dtiObj.InitConfig( "E:\NeoPort\config_dev.ini")

		reqParam = "uniq_id=" + FRectOneRegitem.Funiq_id
		''//���޹޴���---------------------------------------
		reqParam = reqParam + "&biz_no=" + FRectOneRegitem.Fbiz_no
		reqParam = reqParam + "&corp_nm=" + FRectOneRegitem.Fcorp_nm
		reqParam = reqParam + "&ceo_nm=" + FRectOneRegitem.Fceo_nm
		reqParam = reqParam + "&biz_status=" + FRectOneRegitem.Fbiz_status
		reqParam = reqParam + "&biz_type=" + FRectOneRegitem.Fbiz_type
		reqParam = reqParam + "&addr=" + FRectOneRegitem.Faddr
		reqParam = reqParam + "&dam_nm=" + FRectOneRegitem.Fdam_nm
		reqParam = reqParam + "&email=" + FRectOneRegitem.Femail
		reqParam = reqParam + "&hp_no1=" + Left(FRectOneRegitem.Fhp_no1,4)
		reqParam = reqParam + "&hp_no2=" + Left(FRectOneRegitem.Fhp_no2,4)
		reqParam = reqParam + "&hp_no3=" + Left(FRectOneRegitem.Fhp_no3,4)
		''//-------------------------------------------------
		reqParam = reqParam + "&write_date=" + replace(replace(FRectOneRegitem.Fwrite_date,"-",""),"/","")
		reqParam = reqParam + "&sb_type=" + FRectOneRegitem.Fsb_type
		reqParam = reqParam + "&tax_type=" + FRectOneRegitem.Ftax_type
		reqParam = reqParam + "&bill_type=" + FRectOneRegitem.Fbill_type
		reqParam = reqParam + "&pc_gbn=" + FRectOneRegitem.Fpc_gbn
		reqParam = reqParam + "&vol_no="
		reqParam = reqParam + "&issue_no="
		reqParam = reqParam + "&serial_no=" + FRectOneRegitem.Fserial_no
		reqParam = reqParam + "&item_count=" + FRectOneRegitem.Fitem_count
		reqParam = reqParam + "&item_nm=" + FRectOneRegitem.Fitem_nm
		reqParam = reqParam + "&item_qty=" + FRectOneRegitem.Fitem_qty
		reqParam = reqParam + "&item_price=" + FRectOneRegitem.Fitem_price
		reqParam = reqParam + "&item_amt=" + FRectOneRegitem.Fitem_amt
		reqParam = reqParam + "&item_vat=" + FRectOneRegitem.Fitem_vat
		reqParam = reqParam + "&item_remark=" + FRectOneRegitem.Fitem_remark
		reqParam = reqParam + "&approve_type=" + FRectOneRegitem.Fapprove_type

		'reqParam = reqParam + "&cur_u_user_no=" + FRectOneRegitem.Fcur_u_user_no
		'reqParam = reqParam + "&cur_dam_nm=" + FRectOneRegitem.Fcur_dam_nm
		'reqParam = reqParam + "&cur_email=" + FRectOneRegitem.Fcur_email
		'reqParam = reqParam + "&cur_hp_no1=" + FRectOneRegitem.Fcur_hp_no1
		'reqParam = reqParam + "&cur_hp_no2=" + FRectOneRegitem.Fcur_hp_no2
		'reqParam = reqParam + "&cur_hp_no3=" + FRectOneRegitem.Fcur_hp_no3

		response.write reqParam
		'------------------------------------------------------------------------------
		' CallAPI()  : API�� ȣ���Ͽ� ������ ����Ÿ�� �����Ѵ�.
		'
		' Input
		'   api_no : 1=����,2=���
		'   reqParam : ���ݰ�꼭 ����Ÿ (�ڼ��� ������ ��÷ ��������)
		'   reserved : ����� ������
		' Return
		'   ������ : ���ݰ�꼭��ȣ|OK (���ݰ�꼭 ��ȣ�� 0���� ū ��)
		'   ���н� : ������ȣ|��������
		'               -1 : �Ϲ����� ����
		'               -2 : �����ڰ� ȸ������ �ȵ�
		'               -3 : �����ڰ� ���Ҿ�ü�ε� ���ݾȵ�
		'               -4 : �����ڰ� ����ȸ����
		'------------------------------------------------------------------------------
		FResultMsgALL = dtiObj.CallAPI("1", reqParam, "")

		'------------------------------------------------------------------------------

		response.write FResultMsgALL

		If FResultMsgALL <> "" Then
		    tmpArr = Split(FResultMsgALL, "|")

		    if UBound(tmpArr)>=0 then
				Ftax_no = tmpArr(0)
			end if

			if UBound(tmpArr)>=1 then
				Fresultmsg = tmpArr(1)
			end if

		    sqlStr = " update [db_jungsan].[dbo].tbl_tax_history_master" + vbCrlf
			sqlStr = sqlStr + " set tax_no='" + Ftax_no + "'" + vbCrlf
			sqlStr = sqlStr + " , resultmsg='" + Fresultmsg + "'" + vbCrlf
			sqlStr = sqlStr + " where idx=" + CStr(FRectOneRegitem.Fidx) + vbCrlf

			rsget.Open sqlStr,dbget,1

		End If

		Set dtiObj = Nothing
	end sub

	public sub SavePreData()
		dim sqlstr
		dim iid, ouniq_id

		sqlStr = "select * from [db_jungsan].[dbo].tbl_tax_history_master where 1=0"
		rsget.Open sqlStr,dbget,1,3
		rsget.AddNew
		rsget("jungsanid") = CLng(FRectOneRegitem.Fjungsanid)
		rsget("jungsangubun") = FRectOneRegitem.Fjungsangubun
		rsget("makerid") = FRectOneRegitem.Fmakerid
		rsget("jungsanname") = html2db(FRectOneRegitem.Fjungsanname)

		rsget("biz_no") = FRectOneRegitem.Fbiz_no
		rsget("corp_nm") = html2db(FRectOneRegitem.Fcorp_nm)
		rsget("ceo_nm") = html2db(FRectOneRegitem.Fceo_nm)
		rsget("biz_status") = html2db(FRectOneRegitem.Fbiz_status)
		rsget("biz_type") = html2db(FRectOneRegitem.Fbiz_type)
		rsget("addr") = html2db(FRectOneRegitem.Faddr)
		rsget("dam_nm") = html2db(FRectOneRegitem.Fdam_nm)
		rsget("email") = html2db(FRectOneRegitem.Femail)

		rsget("hp_no") = FRectOneRegitem.Fhp_no1 + FRectOneRegitem.Fhp_no2 + FRectOneRegitem.Fhp_no3

		rsget("write_date") = FRectOneRegitem.Fwrite_date
		rsget("sb_type") = FRectOneRegitem.Fsb_type
		rsget("tax_type") = FRectOneRegitem.Ftax_type
		rsget("bill_type") = FRectOneRegitem.Fbill_type
		rsget("pc_gbn") = FRectOneRegitem.Fpc_gbn

		rsget("item_count") = FRectOneRegitem.Fitem_count
		rsget("item_nm") = html2db(FRectOneRegitem.Fitem_nm)
		rsget("item_qty") = FRectOneRegitem.Fitem_qty
		rsget("item_price") = FRectOneRegitem.Fitem_price
		rsget("item_amt") = FRectOneRegitem.Fitem_amt
		rsget("item_vat") = FRectOneRegitem.Fitem_vat
		rsget("item_remark") = html2db(FRectOneRegitem.Fitem_remark)

		rsget("cur_dam_nm") = html2db(FRectOneRegitem.Fcur_dam_nm)
		rsget("cur_email") = html2db(FRectOneRegitem.Fcur_email)
		rsget("cur_hp_no") = FRectOneRegitem.Fcur_hp_no1 + FRectOneRegitem.Fcur_hp_no2 + FRectOneRegitem.Fcur_hp_no3

		rsget.update
			iid = rsget("idx")
		rsget.close


		ouniq_id = replace(replace(CStr(DateSerial(Year(now),month(now),Day(now))),"-",""),"/","")
		ouniq_id = ouniq_id & Format00(5,Right(CStr(iid),5))

		sqlStr = " update [db_jungsan].[dbo].tbl_tax_history_master" + vbCrlf
		sqlStr = sqlStr + " set uniq_id='" + ouniq_id + "'" + vbCrlf
		sqlStr = sqlStr + " where idx=" + CStr(iid) + vbCrlf

		rsget.Open sqlStr,dbget,1

		FRectOneRegitem.Fidx = iid
		FRectOneRegitem.Funiq_id = ouniq_id
		FRectOneRegitem.Fserial_no = ouniq_id
		''FRectOneRegitem.Fserial_no = Right(ouniq_id,10)
	end sub

	Private Sub Class_Initialize()

	end sub

	Private Sub Class_Terminate()

	End Sub
end class
%>