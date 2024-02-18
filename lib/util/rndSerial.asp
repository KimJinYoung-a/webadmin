<%
	'// ���� �ø���Ű ����(���ڸ�����)
	Function rdmSerialEnc(strNum)
		Dim rstStr, keyNo, tmpNo, strProc
		Dim lp
		if strNum = "" then
			rdmSerialEnc = ""
			Exit Function
		end if

		'Ű�� ����
		randomize
		keyNo = int(rnd*10)

		'������ �߰�
		strProc = strNum & getKeyHash(keyNo & strNum)

		'��Ʈ�� ���� �� ���ڵ�
		For lp=1 to len(strProc)
			tmpNo = mid(strProc, lp, 1)
			if asc(tmpNo)>47 and asc(tmpNo)<58 then
				rstStr = rstStr & getKeyString(keyNo,tmpNo,"enc")
			else
				'���ڰ� �ƴѹ��ڰ� ������ �󰪹�ȯ �� ����
				rstStr = ""
				exit for
			end if
		Next

		if rstStr<>"" then
			rstStr = getKeyString("0",keyNo,"enc") & rstStr
		end if

		'��� ��ȯ
		rdmSerialEnc = rstStr
	end Function

	'// ���� �ø���Ű ��ȣȭ
	Function rdmSerialDec(strNum)
		Dim rstStr, keyNo, tmpNo
		Dim lp
		if strNum = "" or len(strNum)<2 then
			rdmSerialDec = ""
			Exit Function
		end if

		'Ű�� ����
		keyNo = getKeyString("0",left(strNum,1),"dec")
		if keyNo = "" then
			rdmSerialDec = ""
			Exit Function
		end if

		'��Ʈ�� ���� �� ���ڵ�
		For lp=2 to len(strNum)
			tmpNo = mid(strNum, lp, 1)

			rstStr = rstStr & getKeyString(keyNo,tmpNo,"dec")

			'���� �ڵ尡 ������ �󰪹�ȯ �� ����
			if getKeyString(keyNo,tmpNo,"dec")="" then
				rstStr = ""
				Exit For
			end if
		Next

		'������ Ȯ��
		if rstStr<>"" then
			if getKeyHash(keyNo & left(rstStr,len(rstStr)-1))=right(rstStr,1) then
				rstStr = left(rstStr,len(rstStr)-1)
			else
				rstStr = ""
			end if
		end if

		'��� ��ȯ
		rdmSerialDec = rstStr
	end Function

	'// Ű���� ���� ���ڿ� ��ȯ
	function getKeyString(strK,strS,mode)
		dim ArrKeyString, i
		Select Case strK
			Case "G",0
				ArrKeyString = "G,E,C,Z,T,9,W,1,B,A"
			Case "E",1
				ArrKeyString = "P,M,A,C,2,W,Z,Q,0,1"
			Case "C",2
				ArrKeyString = "3,S,Z,X,5,B,I,Y,M,V"
			Case "Z",3
				ArrKeyString = "L,O,H,U,R,4,8,1,5,Z"
			Case "T",4
				ArrKeyString = "6,A,K,X,B,Y,N,7,J,U"
			Case "9",5
				ArrKeyString = "1,K,M,H,6,U,Y,T,E,G"
			Case "W",6
				ArrKeyString = "W,2,R,F,D,7,X,C,V,K"
			Case "1",7
				ArrKeyString = "U,I,3,L,K,J,8,N,H,Y"
			Case "B",8
				ArrKeyString = "T,Y,G,4,B,N,R,9,F,J"
			Case "A",9
				ArrKeyString = "P,Q,L,A,5,Z,X,E,0,D"
		end Select

		ArrKeyString = Split(ArrKeyString,",")

		Select Case mode
			Case "enc"
				getKeyString = ArrKeyString(strS)
			Case "dec"
				for i=0 to 9
					if Cstr(ArrKeyString(i))=Cstr(strS) then
						getKeyString = i
						exit for
					end if
				next
		end Select
	end Function

	'// �ùٸ� ������ Ȯ���ϴ� HASH �� ����(���ڸ�����)
	Function getKeyHash(strH)
		Dim tmpNo, rstNo, lp

		For lp=1 to len(strH)
			tmpNo = mid(strH, lp, 1)
			if asc(tmpNo)>47 and asc(tmpNo)<58 then
				rstNo = rstNo + (tmpNo * lp)
			end if
		Next
		rstNo = rstNo mod len(strH)
		rstNo = right(rstNo,1)

		getKeyHash = rstNo
	End Function
%>