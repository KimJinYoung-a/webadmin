/* ============================================================================
*
* ��¥���� �ڹٽ�ũ��Ʈ �����Լ�
*
* ��¥�� ���ڿ��� ��ȯ�ϰų� �� �ݴ��
* ���ڿ��� "2010-01-01 13:21:33" �Ǵ� "2010-01-01" �� ������ �����Ѵ�.
*
============================================================================ */

/* ============================================================================
��ȿ�� üũ

��ȿ��(�����ϴ�) ��(��)���� üũ
function isValidMonth(mm);

��ȿ��(�����ϴ�) ��(��)���� üũ
function isValidDay(yyyy, mm, dd);

��ȿ��(�����ϴ�) ��(��)���� üũ
function isValidHour(hh);

��ȿ��(�����ϴ�) ��(��)���� üũ
function isValidMin(mi);

��ȿ��(�����ϴ�) ������ üũ
function isValidSec(sec);

============================================================================ */



/* ============================================================================
��ȯ

�ڹٽ�ũ��Ʈ Date ��ü�� ��Ʈ������ ��ȯ
(�Է� : Date)
(��� : "YYYY-MM-DD")
function toDateString(Date date);

�ڹٽ�ũ��Ʈ Date ��ü�� ��Ʈ������ ��ȯ
(�Է� : Date)
(��� : "YYYY-MM-DD HH:MM:SS")
function toTimeString(Date date);

��Ʈ���� �ڹٽ�ũ��Ʈ Date ��ü�� ��ȯ
(�Է� : "YYYY-MM-DD HH:MM:SS" �Ǵ� "YYYY-MM-DD")
(��� : Date)
function toDate(time);

============================================================================ */



/* ============================================================================
��¥ ����

��¥�� ���Ѵ�.
(�Է� : Date)
(��� : Date)
function addDate(datefrom, day)
function addMonth(datefrom, month)

�� ��¥ ������ ��¥ ���̸� ���Ѵ�.
(�Է� : Date)
function getDayInterval(time1,time2)

Ҵ�� YYYY�������� ����
(�Է� : Date)
function getYear(time)

���� MM�������� ����
(�Է� : Date)
function getMonth(time)

���� DD�������� ����
(�Է� : Date)
function getDay(time)

Ư�� ��¥�� ����
(�Է� : Date)
(��� : '��','��','ȭ','��','��','��','��')
function getDayOfWeek(time)

============================================================================ */



// ============================================================================
function isValidMonth(mm) {
	var m = parseInt(mm,10);
	return (m >= 1 && m <= 12);
}

function isValidDay(yyyy, mm, dd) {
	var m = parseInt(mm,10) - 1;
	var d = parseInt(dd,10);

	var end = new Array(31,28,31,30,31,30,31,31,30,31,30,31);
	if ((yyyy % 4 == 0 && yyyy % 100 != 0) || yyyy % 400 == 0) {
		end[1] = 29;
	}

	return (d >= 1 && d <= end[m]);
}

function isValidHour(hh) {
	var h = parseInt(hh,10);
	return (h >= 1 && h <= 24);
}

function isValidMin(mi) {
	var m = parseInt(mi,10);
	return (m >= 1 && m <= 60);
}

function isValidSec(sec) {
	var s = parseInt(sec,10);
	return (s >= 1 && s <= 60);
}



// ============================================================================
function toDateString(date) {
	var year  = date.getFullYear();
	var month = date.getMonth() + 1; // 1��=0,12��=11�̹Ƿ� 1 ����
	var day   = date.getDate();

	if (("" + month).length == 1) { month = "0" + month; }
	if (("" + day).length   == 1) { day   = "0" + day;   }

	return ("" + year + "-" + month + "-" + day);
}

function toTimeString(date) {
	var hour  = date.getHours();
	var min   = date.getMinutes();
	var sec   = date.getSeconds();

	if (("" + hour).length  == 1) { hour  = "0" + hour;  }
	if (("" + min).length   == 1) { min   = "0" + min;   }
	if (("" + sec).length   == 1) { sec   = "0" + sec;   }

	return ("" + toDateString(date) + " " + hour + ":" + min + ":" + sec)
}

function toDate(time) {
	// (�Է� : "YYYY-MM-DD HH:MM:SS" �Ǵ� "YYYY-MM-DD")
	var year, month, day, hour, min, sec;

	if (time.length >= 10) {
		year  = time.substr(0,4);
		month = time.substr(5,2) - 1; // 1��=0,12��=11
		day   = time.substr(8,2);
	}

	if (time.length >= 19) {
		hour  = time.substr(11,2);
		min   = time.substr(14,2);
		sec   = time.substr(17,2);

		return new Date(year,month,day,hour,min, sec);
	}

	return new Date(year,month,day);
}



// ============================================================================
function addDate(datefrom, day) {
	// var dateto = datefrom; <=========== ������������. �������̴�.
	var dateto = new Date(datefrom);

	dateto.setDate(datefrom.getDate() + day);

	return dateto;
}

function addMonth(datefrom, month) {
	var dateto = new Date(datefrom);

	dateto.setMonth(datefrom.getMonth() + month);

	return dateto;
}

function getDayInterval(date1,date2) {
	var day   = 1000 * 3600 * 24; // 24�ð�

	return parseInt((date2 - date1) / day, 10);
}

function getYear(time) {
   return time.getFullYear();
}

function getMonth(time) {
   var month = time.getMonth() + 1; // 1��=0,12��=11�̹Ƿ� 1 ����
   if (("" + month).length == 1) { month = "0" + month; }

   return month;
}

function getDay(time) {
   var day = time.getDate();
   if (("" + day).length == 1) { day = "0" + day; }

   return day;
}

function getDayOfWeek(time) {
   var day = time.getDay(); //�Ͽ���=0,������=1,...,�����=6
   var week = new Array('��','��','ȭ','��','��','��','��');

   return week[day];
}
