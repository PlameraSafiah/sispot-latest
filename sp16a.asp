<html>
<head>
<title>Senarai Pekerja Memperolehi Kelulusan Tuntutan Sepertiga</title>
</head>
<body>
<script type="text/javascript" src="menu1.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
function check(f){
	if (f.blnthn.value=="")
	{
		alert("Sila masukkan bulan dan tahun");
		f.blnthn.focus();
		return false
	}
}
//  End -->
</script>


<form action="sp16a.asp" method="POST">
<%Set objConn = Server.CreateObject("ADODB.Connection")
  objConn.Open "dsn=12c;uid=majlis;pwd=majlis;"
  
  

tahun=request.querystring("tahun")
bulan=request.querystring("bulan")
proses = request.querystring("proses")
kod = request.querystring("kod")
ptj=plokasi


   'check jabatan by id pekerja
    pekd = request.cookies("gnop")
    a =     "select jabatan from majlis.pengguna where no_pekerja = '"& pekd &"' "
    Set rsa = objConn.Execute(a)


    'set bulan/tahun - arahan kerja 
    q4="select distinct to_char(tkh_ot,'mm/yyyy') as tkh_ot"
    q4 = q4 & " from payroll.jadual_ot order by tkh_ot"  
    set oq4 = objConn.Execute(q4)
    tkh_ot2 = oq4 ("tkh_ot")

    
	'data pengguna login
    idd = "select a.nama, a.lokasi, initcap(b.keterangan) keterangan, to_char(sysdate,'dd/mm/yyyy') tkh  from payroll.paymas a, iabs.jabatan b "
    idd = idd & " where a.no_pekerja = '"&pekd&"' and b.kod = a.lokasi "
    set idd2 = objConn.Execute(idd)
    'response.write idd
    lokasi = idd2 ("lokasi")
	ket = idd2 ("keterangan")
	

   se = " select distinct a.id , b.nama"
   se = se & " from payroll.selenggara_kerja a , payroll.paymas b"
   se = se & " where a.id = b.no_pekerja and bulan = '"&bulan&"'  and tahun = '"&tahun&"' and lokasi='"& kod &"' order by a.id asc " 
   Set rse= objConn.Execute (se)
   'response.write se
   id = rse("id")
   nama = rse("nama")
   
   
   
  ff = " select keterangan from iabs.jabatan where kod='"&kod&"' "
  Set rsff= objConn.Execute (ff)
  keterangan = rsff ("keterangan")
  	
%>


<table align="center" border="0" cellpadding="0" cellspacing="0" width="80%">
<tr>
  <td align="center" width="10%" valign="top" rowspan="3" >
  <img border="0" src="images/logompsp.jpg"></td>
  <td valign="top" align="center">
      <strong><font size="3" face="Verdana">
      MAJLIS PERBANDARAN SEBERANG PERAI<br>
      SENARAI PEKERJA<br>
      KELULUSAN TUNTUTAN MELEBIHI SEPERTIGA<br> 
      JABATAN <%=keterangan%><br>
      BULAN <%=bulan%>/<%=tahun%></font></strong></td>
  <td rowspan="3" width="10%"></td>
</tr>
</table>
<br><br>


<TABLE width="60%" align="center"> 
    <tr> 
    <td width="6%"   align="center"><B>Bil</B></td>
    <td width="94%"  align="left"><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Senarai Nama Pekerja Mendapat Kelulusan Tuntutan Melebihi Sepertiga</B></td>
    </tr>
</table>
<hr width="60%">
  
  
<%  
bil=0
Do While Not rse.eof 
bil=bil + 1
%>

<TABLE width="60%" align="center"> 
  <tr>
    <td width="6%" align="center"><%=bil%></td>
    <td width="94%" align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rse ("id")%>-<%=rse ("nama")%></td>
  </tr>
</table>
<%  
rse.movenext
loop 
%>  
<hr width="60%"> 