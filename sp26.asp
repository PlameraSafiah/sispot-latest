<%'<--#include file="connection.asp" -->
'<--#INCLUDE FILE="adovbs.inc" -->%>
<!--#include file="connection.asp" -->
<%response.buffer=true%>
<%response.cookies("amenu") = "sp26.asp"%>
<!--'#include file="spmenu.asp"-->


<html>
<head>
<title>Cetak Arahan Kerja</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" href="css/menu1.css" REL="stylesheet">

<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">
<script type="text/javascript" src="menu1.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
function check(f){
	if (f.blnthn.value=="")
	{
		alert("Sila masukkan bulan dan tahun arahan kerja");
		f.blnthn.focus();
		return false
	}
}
//  End -->
</script>
<script language="javascript">
function submitForm1()
{
	document.test.submit();
}
</script>
</head>


<body leftmargin="0" topmargin="0" OnLoad="cssHorScrolls();">
<% 
proses = request.form("proses")
blnthn = request.form("blnthn")
tkh_ot = request.form("tkh_ot")
pemohon = request.form("pemohon")
id = request.form("id_ak")
p6 = request.form("b6")	'cetak arahan kerja
p7 = request.form("b7")	'cetak permohonan
pekd = request.cookies("gnop")
ptj=plokasi


papar
if proses <> "" then	
			borang
end if


sub papar ()
if blnthn = "" then
Dim TodayYearNumber, TodayMonthName
TodayYearNumber = DatePart("yyyy",Date)
TodayMonth = DatePart("m",Date)
if len(TodayMonth)="1" then TodayMonth="0"& TodayMonth
blnthn=TodayMonth&TodayYearNumber
end if


 'check jabatan by id pekerja
   a =     "select jabatan from majlis.pengguna where no_pekerja = '"& pekd &"' "
   Set rsa = objConn.Execute(a)
   
    'check details pengguna by id
   idd = "select a.nama, a.lokasi, b.keterangan, to_char(sysdate,'dd/mm/yyyy') tkh  from payroll.paymas a, iabs.jabatan b "
   idd = idd & " where a.no_pekerja = '"&pekd&"' and b.kod = a.lokasi "
   set idd2 = objConn.Execute(idd)
   'response.write idd
   lokasi = idd2 ("lokasi")
   keterangan = idd2 ("keterangan")
   
%>
<br>


<form name="test" method="post" action="sp26.asp" onSubmit="return check1(this)">
  <TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd">
    <tr class="hd">
      <td colspan="2">Semak & Cetak Arahan Kerja &nbsp;&nbsp;(<font color="#FF0000"><b>*</b></font>Pastikan pengesahan arahan kerja telah dilaksanakan sebelum membuat cetakan)</td>
    </tr> 
  <tr>
    <td width="15%"><b>Bulan/Tahun (mm/yyyy) : </b></td>
  <td><!--<input type="text" name="blnthn" size="8" value="<%=tkh_ot%>">-->
     <select size="1" name="tkh_ot" onChange="submitForm1(this)">
          <option value="">Pilih Bulan/Tahun</option>
          <%     
   'semak bulan/tahun OT 
   q2="select distinct to_char(a.tkh_ot,'mm/yyyy') as tkh_ot "
   q2 = q2 & " from payroll.jadual_ot a , payroll.arahan_kerja b where a.id_ak=b.id_ak and b.ptj = '"&ptj&"' and b.pengesahan='Y' "
   q2 = q2 & " order by tkh_ot desc"  
   set oq2 = objConn.Execute(q2)
   'response.write q2
   if not oq2.bof and oq2.eof then
   tkh_ot = oq2 ("tkh_ot")
   end if
		  		  
		  do while not oq2.EOF  %>
          <option <%if tkh_ot = oq2("tkh_ot") then%> selected <%end if%>value="<%=oq2("tkh_ot")%>"><%=oq2("tkh_ot")%></option>
          <%
            oq2.MoveNext
            loop
'------
oq2.close
'------

           %>
    </select>
    </td>
    </tr> 
   <tr>
  <td><b>Nama Pemohon : </b></td>
  <td><!--<input type="text" name="pekerja" size="8" value="<%=pemohon%>">-->
  <select size="1" name="pemohon">
  <option value="">Pilih Pemohon</option>
   <%	
   if tkh_ot <> "" then 
   q2b="select distinct a.pemohon, to_char(c.tkh_ot,'mmyyyy') , b.nama "
   q2b = q2b & " from payroll.arahan_kerja a, payroll.jadual_ot c , payroll.paymas b where a.ptj= "&ptj&" and to_char(c.tkh_ot,'mm/yyyy') like '"& tkh_ot &"' "
   q2b = q2b & "and a.id_ak = c.id_ak and a.pemohon = b.no_pekerja and a.pengesahan = 'Y' order by a.pemohon"
   set oq2b = objConn.Execute(q2b)
   'response.write  q2b
   pemohon = oq2b ("pemohon")
   nama = oq2b ("nama")
     
		  do while not oq2b.EOF 
		  %>
          <option <%if cstr(pemohon) =  cstr(oq2b("pemohon")) then%> selected <%end if%>value="<%=oq2b("pemohon")%>"><%=oq2b("pemohon")%> - <%=oq2b("nama")%></option>
          <%
            oq2b.MoveNext
            loop
'------
oq2b.close
'------
            end if   
		%>
    </select>
  </td>
</tr>
   
       
    <tr align="center">
      <td colspan="2"><input type="submit" name="proses" value="Hantar"></td>
    </tr>
  </table>
</form>
<% end sub %>



<%
sub borang()
pemohon = request.form("pemohon")
id = request.form("id_ak")
tkh_ot = request.form("tkh_ot")
ptj=plokasi
p7 = request.form("b7")



  'papar maklumat arahan kerja
  ss = " select distinct a.id_ak , a.penyelia , a.keterangan , a.pemohon , a.ptj , a.nota , a.tahun , a.bulan , to_char(b.tkh_ot,'mm/yyyy') as tkh_ot , a.pengesahan "
  ss = ss & " from payroll.arahan_kerja a , payroll.jadual_ot b where a.id_ak=b.id_ak and to_char(b.tkh_ot,'mm/yyyy') = '"&tkh_ot&"' and a.ptj = '"&ptj&"' and "
  ss = ss & " a.pemohon='"& pemohon &"' and a.pengesahan='Y' " 
  ss = ss & " order by a.id_ak "
  'response.write ss
  Set rsss = objConn.Execute (ss)
  	id = rsss ("id_ak")
	penyelia = rsss ("penyelia")
	pemohon = rsss ("pemohon")
	keterangan = rsss("keterangan")
    ptj = rsss ("ptj")
	nota = rsss ("nota")
    bln = rsss ("bulan")
	thn = rsss ("tahun")
	
		
  'papar keterangan jabatan
  m = " select kod , initcap (keterangan) keterangan from iabs.jabatan "
  m = m & " where kod = '"&ptj&"' "
  Set rsm = objConn.Execute(m)
  ket = rsm ("keterangan")	
  
  
   'papar maklumat pemohon
   qqp = " select no_pekerja , nama , jawatan from payroll.paymas " 
   qqp = qqp & " where no_pekerja = '"&pemohon&"' " 
   Set rsqqp = objConn.Execute (qqp)
   nama3 = rsqqp ("nama")
   jaw1 = rsqqp ("jawatan")	
		
%>
<br><br>



<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all class="hd" height="20%">
  <TBODY>
  <tr class="hd"> 
    <td colspan="10" >Cetakan Arahan Kerja<br>
     Jabatan <%=ket%> <br> 
    Nama Pemohon : <%=nama3%> (<%=pemohon%>)<br>
    Bulan <%=tkh_ot%>
  </td>
  </tr> 
  <tr style="BACKGROUND-COLOR: #ffffff"><td colspan="9">&nbsp;</td></tr>
  <tr class="hd"> 
    <td width="2%"  align="center">Bil</td>
    <td width="8%"  align="center">No Rujukan Arahan Kerja</td>
    <td width="7%"  align="center">Bulan/Tahun</td>
    <td width="8%"  align="center">Jabatan</td>
    <td width="19%" align="center">Penyelia</td>
    <td width="7%"  align="center">Status Pengesahan</td>
    <td width="23%" align="center">Keterangan Kerja</td>
    <td width="8%"  align="center">Anggaran OT (RM)</td>
    <td width="25%" align="center">Proses</td>
  </tr>

  
<%  
bil=0
Do While Not rsss.eof 
bil=bil + 1 
id = rsss ("id_ak")	
penyelia = rsss ("penyelia")
pemohon = rsss ("pemohon")	
   
   'nama penyelia
   pp = " select no_pekerja , nama , jawatan from payroll.paymas " 
   pp = pp & " where no_pekerja = '"&penyelia&"' " 
   Set rspp = objConn.Execute (pp)
if not rspp.eof then
   nama1 = rspp ("nama")
   jaw1 = rspp ("jawatan")	
end if	
	
   'anggaran amaun OT
   se = " select sum(nvl(angg_ot_s,0) + nvl(angg_ot_m,0)) anggaran "
   se = se & " from payroll.pekerja_ot"
   se = se & " where id_ak='"&id&"' and to_char(tkh_ot,'mm/yyyy') ='"& tkh_ot &"' " 
   Set rse= objConn.Execute (se)
   'response.write se
   anggaran = rse("anggaran")
  
%>
  <tr bgcolor="#EFF5F5">
    <td rowspan="4" align="center"><%=bil%></td>
    <td rowspan="4" align="center"><%=rsss ("id_ak")%></td>
    <td rowspan="4" align="center"><%=tkh_ot%></td>
    <td rowspan="4" align="center"><%=rsss ("ptj")%></td>
    <td rowspan="4" align="left"><%=rsss ("penyelia")%>-<%=nama1%></td>
    <td rowspan="4" align="center"><%=rsss ("pengesahan")%></td>
    <td rowspan="4" align="left"><%=rsss ("keterangan")%></td>
    <td rowspan="4" align="center"><%=rse("anggaran")%></td>
    </tr>
     
     <tr bgcolor="#EFF5F5">
    <td align="center" height="20">&nbsp;&nbsp;&nbsp; 
    <%if rse("anggaran") = "" then%>
   <form name="test<%=bil%>"  method="post" action="sp26a.asp?id_ak=<%=id%>&pemohon=<%=pemohon%>&amenu='sp26a'" target="_new">
   <input type="hidden" name="b6" value="Cetak Arahan Kerja" class="button" onFocus="nextfield='done';" > 
   </form>&nbsp;&nbsp;&nbsp;  
   <%end if%>
   <%if rse("anggaran") <> "" then%>
   <form name="test<%=bil%>"  method="post" action="sp26a.asp?id_ak=<%=id%>&pemohon=<%=pemohon%>&amenu='sp26a'" target="_new">
   <input type="submit" name="b6" value="Cetak Arahan Kerja" class="button" onFocus="nextfield='done';" >    
   </form>&nbsp;&nbsp;&nbsp;  
   <%end if%>   
   </td> 
  </tr>
  
    <tr bgcolor="#EFF5F5">
   <td align="center" height="5">
   <form name="test<%=bil%>"  method="post" action="sp26b.asp?id_ak=<%=id%>&amenu='sp26b'" target="_new">
   <input type="hidden" name="b7" value="Cetak Permohonan Seorang Pekerja" class="button" onFocus="nextfield='done';">
    </form>
   </td>
  </tr>
  
   <tr bgcolor="#EFF5F5">
   <td align="center" height="20">&nbsp;&nbsp;&nbsp;
    <%if rse("anggaran") = "" then%>
   <form name="test<%=bil%>"  method="post" action="sp26c.asp?id_ak=<%=id%>&amenu='sp26c'" target="_new">
   <input type="hidden" name="b7" value="Cetak Senarai Pekerja" class="button" onFocus="nextfield='done';">
    </form>&nbsp;&nbsp;&nbsp;
    <%end if%>
        <%if rse("anggaran") <> "" then%>
   <form name="test<%=bil%>"  method="post" action="sp26c.asp?id_ak=<%=id%>&amenu='sp26c'" target="_new">
   <input type="submit" name="b7" value="Cetak Senarai Pekerja" class="button" onFocus="nextfield='done';">
   </form>&nbsp;&nbsp;&nbsp;
   <%end if%>   
   </td>
  </tr>
    
 <% 
  rsss.movenext
  loop
'------
rsss.close
objConn.close
'------ 
 %>
   
  </tbody>  
</table>
 <br>
  <TABLE border=1 align="center">
    <tr align="center">
    <td colspan="2">
   <form name="test<%=bil%>"  method="post" action="sp26a1.asp?pemohon=<%=pemohon%>&amenu='sp26a1'" target="_new">
   <input type="submit" name="b6" value="Cetak Arahan Kerja Keseluruhan" class="button" onFocus="nextfield='done';" >    
   </form> 
    </td>
    <td colspan="2">
   <form name="test<%=bil%>"  method="post" action="sp26c1.asp?pemohon=<%=pemohon%>&amenu='sp26c1'" target="_new">
   <input type="submit" name="b6" value="Cetak Senarai Pekerja Keseluruhan" class="button" onFocus="nextfield='done';" >    
   </form> 
    </td>
     </tr>
  </table>
  
  
<% end sub %>

<input type="hidden" name="bilrec" value="<%=ctrz%>" >
</body>
</html>