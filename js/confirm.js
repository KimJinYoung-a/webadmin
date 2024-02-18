//-----------------------Written by Duke Kim ---------------------------

/*
purpose : ��ü���� �� ������Ʈ ��ȿ�� �˻�
input : �� ��ü
return : ��� ��ҿ� ������ ������ true, ������ false
remark : form ������ onsumbit �̺�Ʈ�� �ڵ鷯�� �ٿ� submit������ �����ϴ� ������� ���
				 ��) <form name="register" method="post" action="join_joining_p.asp" onsubmit="return validate(this)">
				 �� �Լ��� ����ϱ� ���ؼ��� �� ������Ʈ�� id �Ӽ��� ������ �°� �����ؾ� �Ѵ�.

				 ---------------------------------ID Properties GuideLine---------------------------------
				 1. ������ ���� �޽����ڽ� ��½� ID ������Ƽ�� ����ϹǷ� ����ڿ��� �����Ϸ��� ������ �̸��� ����Ѵ�.
				 2. ������׿� ���ؼ��� [off,off,off,off] �� ���� ���ȣ���� ���ڸ� �ĸ��и������� ����Ѵ�. �׸���� ������� ����.
				 3. �ʼ��Է�(��������) ���������� ù�ڸ� ���� --> [on,off,off,off]
				 4. ���ڸ���� ���������� ��°�ڸ� ���� -->			[off,on,off,off]
				 5. �ּұ������� ���������� ��°�ڸ� ���� -->		[off,off,5,off]
				 6. �ִ�������� ���������� ��°�ڸ� ���� -->		[off,off,off,10]
				 7. ��� ������ �����Ҷ��� ����ڸ� ���� -->		[on,on,5,10]

				 example)
					<input type="text" id="[on,off,3,8]���̵�" name="id">
					���� �ʼ��Է�, 3~8�� ���������� �а��̴�.
					���̵� ������ ��� "���̵���� ����׿�" ������ �޼����ڽ��� ��µȴ�.
				 ---------------------------------ID Properties GuideLine---------------------------------
*/

function validate(form)
{

	var rtn_blank, rtn_digit, rtn_length;

	rtn_blank = check_form_blank(form);

	//���� üũ
	if (rtn_blank >= 0)
	{
		form.elements[rtn_blank].focus();
		return false;
	}
	//���ڸ� ���Ǵ� �Է��� üũ
	else
		rtn_digit = check_form_digit(form);
		if (rtn_digit >= 0)
			{
			//form.elements[rtn_digit].value = '';
			form.elements[rtn_digit].focus();
			return false;
			}

	//�ּ�,�ִ���� ���� üũ
	rtn_length = check_form_length(form);
	if (rtn_length >= 0)
		{
		//form.elements[rtn_length].value = '';
		form.elements[rtn_length].focus();
		return false;
		}

	return true;
}

function validate3(form)
{
	var rtn_blank, rtn_digit, rtn_length;
	rtn_blank = check_form_blank(form);

	//���� üũ
	if (rtn_blank >= 0)
	{
		form.elements[rtn_blank].focus();
		return false;
	}
	//���ڸ� ���Ǵ� �Է��� üũ
	else
		rtn_digit = check_form_digit(form);
		if (rtn_digit >= 0)
			{
			//form.elements[rtn_digit].value = '';
			form.elements[rtn_digit].focus();
			return false;
			}
	//�ּ�,�ִ���� ���� üũ
	rtn_length = check_form_length(form);
	if (rtn_length >= 0)
		{
		//form.elements[rtn_length].value = '';
		form.elements[rtn_length].focus();
		return false;
		}

	//�̸����ּ��� ��ȿ���� üũ(@�� ���ԵǾ� �ִ��� Ȯ��)
	else if (check_form_email(form.email.value) == false)
		{
		alert('�̸��� �ּҰ� ��ȿ���� �ʽ��ϴ�.');
		form.email.focus();
		return false;
		}
	//Ȩ������ �ּ��� ��ȿ���� üũ(http://�� ���ԵǾ� �ִ��� Ȯ��)
	/*
	else if  (check_form_url(form.g_url.value) == false)
		{
		alert('Ȩ������ �ּҰ� ��ȿ���� �ʽ��ϴ�.')
		form.g_url.focus();
		return false;
		}

	*/

	// ������ ���� ���θ� üũ
	else if (form.question.value =='')
		{
		alert('������ ���� �Ͻʽÿ�.');

		return false;

		}


	/*
	//�����͸� �ѱ���� ��Ȱ�� �� ������Ʈ�� Ȱ��ȭ �ؾ� �����Ͱ� ���޵�
	var i;
	for (i=0; i < form.elements.length; i++)
		form.elements[i].disabled = false;
	*/
	//�� ����
	return true;
}

function validate2(form)
{
	var rtn_blank, rtn_digit, rtn_length;
	rtn_blank = check_form_blank(form);

	//���� üũ
	if (rtn_blank >= 0)
	{
		form.elements[rtn_blank].focus();
		return false;
	}
	//���ڸ� ���Ǵ� �Է��� üũ
	else
		rtn_digit = check_form_digit(form);
		if (rtn_digit >= 0)
			{
			//form.elements[rtn_digit].value = '';
			form.elements[rtn_digit].focus();
			return false;
			}
	//�ּ�,�ִ���� ���� üũ
	rtn_length = check_form_length(form);
	if (rtn_length >= 0)
		{
		//form.elements[rtn_length].value = '';
		form.elements[rtn_length].focus();
		return false;
		}
	//�̸����ּ��� ��ȿ���� üũ(@�� ���ԵǾ� �ִ��� Ȯ��)
	else if (check_form_email(form.buy_email.value) == false)
		{
		alert('�̸��� �ּҰ� ��ȿ���� �ʽ��ϴ�.');
		form.buy_email.focus();
		return false;
		}
	//Ȩ������ �ּ��� ��ȿ���� üũ(http://�� ���ԵǾ� �ִ��� Ȯ��)
	/*
	else if  (check_form_url(form.g_url.value) == false)
		{
		alert('Ȩ������ �ּҰ� ��ȿ���� �ʽ��ϴ�.')
		form.g_url.focus();
		return false;
		}

	*/

	return true;
}



/*
purpose : ���̵� �ߺ����� üũ�������� ��â�� �ε�
input : �� ��ü
remark : ���̵���� ��������� �����޽���, �׷��� ������ ��â �ε�
*/
function checkid( form )
{
	var id;
	id = 	form.id.value;

	if (id == '')
	{
		alert('���̵���� ����ֳ׿�');
		form.id.focus();
	}
	else
	{
		window.open('/lib/searchid.asp?id=' + id, 'searchid', 'width=400,height=230,location=no,menubar=no,resizable=no,scrollbars=no,status=no,toolbar=no');
	}
}
function checkid2()
{
	var id;
	id = 	document.forms[0].id.value;

	if (id == '')
	{
		alert('���̵���� ����ֳ׿�');
		form.id.focus();
	}
	else
	{
		window.open('/lib/searchid.asp?id=' + id, 'searchid', 'width=400,height=230,location=no,menubar=no,resizable=no,scrollbars=no,status=no,toolbar=no');
	}
}



/*
purpose : �����ȣ �˻� �� ���� �������� ��â�� �ε�
input : �� ��ü
*/
function searchzipcode(type)
{
	window.open('/lib/searchzip.asp?target=' + type, 'searchzip', 'width=460,height=250,scrollbars=yes');
}



/*
purpose : ���� �̸��ϰ��� ��������Ʈ�� ��â���� �ε�
*/
function getemail()
	{
	window.open('http://register.daum.net/', 'getemail');
	}



/*
purpose : ���� ������Ʈ�� value ������Ƽ�� ��ĭ���� �����ִ� ���� üũ
input : �� ��ü
return : ��ĭ�� ������ �ش� ������Ʈ�� �÷��ǳ� �ε�����ȣ, ������ -1
remark : ������Ʈ id ������Ƽ�� ù��° ���������� on �̸� ��ĭ�� ������� �ʴ� ���� �ǹ�
*/
function check_form_blank(form)
	{
	var i, con, id, pos1, pos2, pos3, pos4;
	for (i=0; i < form.elements.length; i++)
		{
		//�ؽ�Ʈ�� �н�����, �ؽ�Ʈ����� �� ���� Ÿ�Կ� ���ؼ���
		if (form.elements[i].type == 'text' || form.elements[i].type == 'password' || form.elements[i].type == 'textarea' || form.elements[i].type == 'hidden')
			{
			//���༳���κ��� �߶�
			id = form.elements[i].id;
			pos1 = id.indexOf(']');							//�������ǰ� ID�� �����ϴ� ��ġ - 1
			pos2 = id.indexOf('[');							//�������������� ���۵Ǵ� ��ġ - 1
			pos3 = id.indexOf(',', pos2 + 1);		//�������������� ������ ��ġ
			pos4 = id.length;										//ID�� ������ ��ġ
			con = id.substring(pos2 + 1, pos3);		//alert(con);
			id = id.substring(pos1 + 1, pos4);		//alert(id);

			//���༳���� �Ǿ��ְ� ��ĭ�̸�
			if (con == 'on' && form.elements[i].value == '')
				{
				alert (id + '���� �Է����ּ���.');
				return(i);
				}
			}
		}
	return(-1);
	}



/*
purpose : ���� ������Ʈ�� ���ڸ� �Է¹޾ƾ� �ϴ� ���� ã�� value ������Ƽ�� üũ
input : �� ��ü
return : ���ڹ����� �Ѿ�� ���� ������ �ش� ������Ʈ�� �÷��ǳ� �ε�����ȣ, ������ -1
remark :
*/
function check_form_digit(form)
	{
	var i, con, id, pos1, pos2, pos3, pos4, j, digit;
	for (i=0; i < form.elements.length; i++)
		{
		//�ؽ�Ʈ�� �н�����, �ؽ�Ʈ����� Ÿ�Կ� ���ؼ���
		if (form.elements[i].type == 'text' || form.elements[i].type == 'password' || form.elements[i].type == 'textarea')
			{
			//���༳���κ��� �߶�
			id = form.elements[i].id;
			pos1 = id.indexOf(']');						//�������ǰ� ID �� �����ϴ� ��ġ - 1
			pos2 = id.indexOf(',');						//�������������� ���۵Ǵ� ��ġ - 1
			pos3 = id.indexOf(',', pos2 + 1);	//�������������� ������ ��ġ
			pos4 = id.length;									//ID�� ������ ��ġ
			con = id.substring(pos2 + 1, pos3);		//alert(con);
			id = id.substring(pos1 + 1, pos4);		//alert(id);

			//���༳���� �Ǿ� ������
			if (con == 'on')
				{
				digit = form.elements[i].value;
				for (j=0; j < digit.length; j++)
					if ((digit.charAt(j) * 0 == 0) == false)
						{
						alert(id + '���� ���ڸ� ���˴ϴ�.');
						return(i);
						}
				}
			}
		}
	return(-1);
	}




/*
purpose : ���� ������Ʈ�� �ּұ������Ѱ� �ִ���������� üũ
input : �� ��ü
return : �ִ� �Ǵ� �ּұ��̸� ����� ���� ������ �ش� ������Ʈ�� �÷��ǳ� �ε�����ȣ, ������ -1
remark :
*/
function check_form_length(form)
	{
	var i, id, pos1, pos2, pos3, pos4, max, min, length;
	for (i=0; i < form.elements.length; i++)
		{
		//�ؽ�Ʈ�� �н�����, �ؽ�Ʈ����� Ÿ�Կ� ���ؼ���
		if (form.elements[i].type == 'text' || form.elements[i].type == 'password' || form.elements[i].type == 'textarea')
			{
			//���༳���κ��� �߶�
			id = form.elements[i].id;
			pos1 = id.indexOf(']');						//�������ǰ� ID �� �����ϴ� ��ġ - 1
			pos2 = id.indexOf(',');						//������������ �ǳʶ�
			pos2 = id.indexOf(',', pos2 + 1);	//������������ �ǳʶ�
			pos2 = id.indexOf(',', pos2);			//�ּұ������������� ���۵Ǵ� ��ġ - 1
			pos3 = id.indexOf(',', pos2 + 1);	//�ּұ������������� ������ ��ġ
			min  = id.substring(pos2 + 1, pos3);		//alert(min);
			pos2 = id.indexOf(',', pos2 + 1);	//�ִ�������������� ���۵Ǵ� ��ġ - 1
			pos3 = id.indexOf(']', pos2);			//�ִ�������������� ������ ��ġ
			max  = id.substring(pos2 + 1, pos3);		//alert(max);
			pos4 = id.length;									//ID�� ������ ��ġ
			id = id.substring(pos1 + 1, pos4);			//alert(id);

			length = (form.elements[i].value).length
			if (id!='' && min != 'off' && max != 'off' )	//�ּ� �Ǵ� �ִ���̰� �����Ǿ� �ִ� ���
				{
				if (max == 'off')
					{
					if (min >= length)		//�ּұ��� ���������� ����
						{
						alert(id + '���� �ּ��� ' + min + '�� �̻��̾�� �մϴ�.');
						return(i);
						}
					}
				else if (min == 'off')
					{
					if (length > max)		//�ִ���� ���������� ����
						{
						alert(id + '���� �ִ� ' + max + '�ڱ����� �Է°����մϴ�.');
						return(i);
						}
					}
				else
					{
					if (min > length || max < length)		//�ּ�, �ִ���� ���������� ����
						{
						alert(id + '���� ' + min + '~' + max + '�� ���̷� �Է��ϼž� �մϴ�.');
						return(i);
						}
					}
				}
			}
		}
	return(-1);
	}




/*
purpose : �ֹε�Ϲ�ȣ�� ��ȿ���� üũ
input : �ֹι�ȣ(�����¾��� �ٿ���)
return : �ùٸ��� true, �ùٸ��� ������ false
remark : �ٷ� ���� jumin_chk �Լ��� �ݵ�� �ʿ��ϴ�.
*/

function check_form_ssn(it1, it2){
	var forigndigit = it2.substring(0,1);
	jumin=it1+it2;

	if ((forigndigit=="5")||(forigndigit=="6")){
		return isRegNo_fgnno(jumin);
	}else{
		if(jumin_chk(jumin)){
			return false;
		}else	{
			return true;
		}
	}
}

function isRegNo_fgnno(fgnno) {
        var sum=0;
        var odd=0;
        buf = new Array(13);
        for(i=0; i<13; i++) { buf[i]=parseInt(fgnno.charAt(i)); }
        odd = buf[7]*10 + buf[8];
        if(odd%2 != 0) { return false; }
        if( (buf[11]!=6) && (buf[11]!=7) && (buf[11]!=8) && (buf[11]!=9) ) {
                return false;
        }
        multipliers = [2,3,4,5,6,7,8,9,2,3,4,5];
        for(i=0, sum=0; i<12; i++) { sum += (buf[i] *= multipliers[i]); }
        sum = 11 - (sum%11);
        if(sum >= 10) { sum -= 10; }
        sum += 2;
        if(sum >= 10) { sum -= 10; }
        if(sum != buf[12]) { return false }
        return true;
}

function jumin_chk(it)
{
	IDtot = 0;
	IDAdd="234567892345";

	for(i=0;i<12;i++)
	{
		IDtot=IDtot+parseInt(it.substring(i,i+1))*parseInt(IDAdd.substring(i,i+1));
	}

	IDtot=11-(IDtot%11);
	if(IDtot==10)
	{
		IDtot=0;
	}
	else if(IDtot==11)
	{
		IDtot=1;
	}

	if(parseInt(it.substring(12,13))!=IDtot)
		return true;
	else
		return false;
}




/*
purpose : �̸��� �ּ��� ��ȿ���� üũ
input : �̸��� �ּ�
return : �ùٸ��� true, �ùٸ��� ������ false
remark : �ּҿ� @�� ���ԵǾ� �ִ���, �Ǵ� �ι��̻� ���Ե����� �ʾҴ��� Ȯ��
*/

function check_form_email(email)
{

	var pos;


	pos = email.indexOf('@');

	if (pos < 0)				//@�� ���ԵǾ� ���� ����
		return(false);
	else
		{
		pos = email.indexOf('@', pos + 1)
		if (pos >= 0)			//@�� �ι��̻� ���ԵǾ� ����
			return(false);
		}


	pos = email.indexOf('.');

	if (pos < 0)				//@�� ���ԵǾ� ���� ����
		return false;


	return(true);

}





/*
purpose : URL�� ��ȿ���� üũ
input : URL
return : �ùٸ��� true, �ùٸ��� ������ false
remark : �ּҰ� http://�� �����ϴ��� Ȯ��
*/

function check_form_url(url)
	{
	var protocol;
	protocol = url.substring(0, 7)

	if (protocol != 'http://')				//http://�� �������� ����
		return(false);
	else
		return(true);
	}

