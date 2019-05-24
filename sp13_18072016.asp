<!--#include file="connection.asp" -->
<!--'#include file="spmenu.asp"-->
<%response.buffer=true%>
<html>
<head>


<title>Kelulusan Melebihi Sepertiga</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link type="text/css" href="css/menu1.css" REL="stylesheet">
<BODY leftMargin=0 onLoad="" topMargin=0 bgcolor="<%=color4%>">

<script type="text/javascript" src="http://code.jquery.com/jquery-latest.min.js"/></script>
<script type="text/javascript" src="../sispot/javascript/jquery-latest.min.js"/></script>
<script type="text/javascript">
$(document).ready(function(){
$('input[name="all"],input[name="title"]').bind('click', function(){
var status = $(this).is(':checked');
$('input[type="checkbox"]', $(this).parent('li')).attr('checked', status);
});

});
</script>
<script type="text/javascript" src="menu1.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin
function check(f){
	if (f.id.value=="")
	{
		alert("Sila masukkan no pekerja");
		f.id.focus();
		return false
	}
	if (f.bulan.value=="")
	{
		alert("Sila masukkan bulan");
		f.bulan.focus();
		return false
	}
	if (f.tahun.value=="")
	{
		alert("Sila masukkan tahun");
		f.tahun.focus();
		return false
	}
	
}
//  End -->
</script>
</head>



<%response.cookies("amenu") = "sp13.asp"%>

<% 
 
   proses = request.form("proses")

   if proses="" then 
       simpan				
   else
     id = request.form("id")
	 bulan1 = request.form("bulan1")
	 bulan2 = request.form("bulan2")
	 bulan3 = request.form("bulan3")
	 bulan4 = request.form("bulan4")
	 bulan5 = request.form("bulan5")
	 bulan6 = request.form("bulan6")
	 bulan7 = request.form("bulan7")
	 bulan8 = request.form("bulan8")
	 bulan9 = request.form("bulan9")
	 bulan10 = request.form("bulan10")
	 bulan11 = request.form("bulan11")
	 bulan12 = request.form("bulan12")
     tahun = request.form("tahun")
	
	if proses="Simpan" then 
	  simpan
	  papar1
	end if
    end if

%>

   
<%
sub simpan()
proses=request("proses")

if proses="Simpan" then
rekod=""
bln = mid(bulan1,3,2)
bln = mid(bulan2,3,2)
bln = mid(bulan3,3,2)
bln = mid(bulan4,3,2)
bln = mid(bulan5,3,2)
bln = mid(bulan6,3,2)
bln = mid(bulan7,3,2)
bln = mid(bulan8,3,2)
bln = mid(bulan9,3,2)
bln = mid(bulan10,3,2)
bln = mid(bulan11,3,2)
bln = mid(bulan12,3,2)
thn = mid(tahun,5,4)

    
	'papar1 data selenggara ---> nadia (24032016)
	if bulan1 <> "" then 
    q0="select * from payroll.selenggara_kerja where tahun like '"& tahun &"' and bulan like '"& bulan1 &"' and id like '"& id &"'  "
	end if
	if bulan2 <> "" then 
    q0="select * from payroll.selenggara_kerja where tahun like '"& tahun &"' and bulan like '"& bulan2 &"' and id like '"& id &"'  "
	end if
	if bulan3 <> "" then 
    q0="select * from payroll.selenggara_kerja where tahun like '"& tahun &"' and bulan like '"& bulan3 &"' and id like '"& id &"'  "
	end if
	if bulan4 <> "" then 
    q0="select * from payroll.selenggara_kerja where tahun like '"& tahun &"' and bulan like '"& bulan4 &"' and id like '"& id &"'  "
	end if
	if bulan5 <> "" then 
    q0="select * from payroll.selenggara_kerja where tahun like '"& tahun &"' and bulan like '"& bulan5 &"' and id like '"& id &"'  "
	end if
	if bulan6 <> "" then 
    q0="select * from payroll.selenggara_kerja where tahun like '"& tahun &"' and bulan like '"& bulan6 &"' and id like '"& id &"'  "
	end if
	if bulan7 <> "" then 
    q0="select * from payroll.selenggara_kerja where tahun like '"& tahun &"' and bulan like '"& bulan7 &"' and id like '"& id &"'  "
	end if
	if bulan8 <> "" then 
    q0="select * from payroll.selenggara_kerja where tahun like '"& tahun &"' and bulan like '"& bulan8 &"' and id like '"& id &"'  "
	end if
	if bulan9 <> "" then 
    q0="select * from payroll.selenggara_kerja where tahun like '"& tahun &"' and bulan like '"& bulan9 &"' and id like '"& id &"'  "
	end if
	if bulan10 <> "" then 
    q0="select * from payroll.selenggara_kerja where tahun like '"& tahun &"' and bulan like '"& bulan10 &"' and id like '"& id &"'  "
	end if
	if bulan11 <> "" then 
    q0="select * from payroll.selenggara_kerja where tahun like '"& tahun &"' and bulan like '"& bulan11 &"' and id like '"& id &"'  "
	end if
	if bulan12 <> "" then 
    q0="select * from payroll.selenggara_kerja where tahun like '"& tahun &"' and bulan like '"& bulan12 &"' and id like '"& id &"'  "
	end if
    set r0=objConn.Execute(q0)
	
	
	if r0.eof then
		'q2="Insert into payroll.selenggara_kerja (id , bulan , tahun) values ('"& id &"','"& bulan &"','"& tahun &"')"
		'objConn.Execute(q2) 
		
		'simpan data selenggara ---> nadia (15032016)					
		q2="Insert all "
		if bulan1 <> "" then
		q2 = q2 & " into payroll.selenggara_kerja (id , bulan , tahun) values ('"& id &"','"& bulan1 &"','"& tahun &"') "
		end if
		if bulan2 <> "" then
		q2 = q2 & " into payroll.selenggara_kerja (id , bulan , tahun) values ('"& id &"','"& bulan2 &"','"& tahun &"') "
		end if
		if bulan3 <> "" then 
		q2 = q2 & " into payroll.selenggara_kerja (id , bulan , tahun) values ('"& id &"','"& bulan3 &"','"& tahun &"') "
		end if
		if bulan4 <> "" then
		q2 = q2 & " into payroll.selenggara_kerja (id , bulan , tahun) values ('"& id &"','"& bulan4 &"','"& tahun &"') "
		end if
		if bulan5 <> "" then
		q2 = q2 & " into payroll.selenggara_kerja (id , bulan , tahun) values ('"& id &"','"& bulan5 &"','"& tahun &"') "
		end if
		if bulan6 <> "" then 
		q2 = q2 & " into payroll.selenggara_kerja (id , bulan , tahun) values ('"& id &"','"& bulan6 &"','"& tahun &"') "
		end if
		if bulan7 <> "" then
		q2 = q2 & " into payroll.selenggara_kerja (id , bulan , tahun) values ('"& id &"','"& bulan7 &"','"& tahun &"') "
		end if
		if bulan8 <> "" then 
		q2 = q2 & " into payroll.selenggara_kerja (id , bulan , tahun) values ('"& id &"','"& bulan8 &"','"& tahun &"') "
		end if
		if bulan9 <>"" then 
		q2 = q2 & " into payroll.selenggara_kerja (id , bulan , tahun) values ('"& id &"','"& bulan9 &"','"& tahun &"') "
		end if
		if bulan10 <> "" then 
		q2 = q2 & " into payroll.selenggara_kerja (id , bulan , tahun) values ('"& id &"','"& bulan10 &"','"& tahun &"') "
		end if
		if bulan11 <> "" then 
		q2 = q2 & " into payroll.selenggara_kerja (id , bulan , tahun) values ('"& id &"','"& bulan11 &"','"& tahun &"') "
		end if
		if bulan12 <> "" then 
		q2 = q2 & " into payroll.selenggara_kerja (id , bulan , tahun) values ('"& id &"','"& bulan12 &"','"& tahun &"') "
		end if
		q2 = q2 & " select * from dual  "
		objConn.Execute(q2) 	
		
		
		
		rekod="simpan"
		msg="Rekod baru telah selesai ditambah"
	else
		rekod="duplicate"
		if bulan1 <> "" then
		msg="Rekod baru untuk id <b>"& id &"</b> , bulan <b> '"&bulan1&"'</b>  ,tahun <b>"& tahun &"</b> telah wujud.<br>Proses tidak dapat diteruskan"
		end if
		if bulan2 <> "" then
		msg="Rekod baru untuk id <b>"& id &"</b> , bulan <b> '"&bulan2&"'</b>  ,tahun <b>"& tahun &"</b> telah wujud.<br>Proses tidak dapat diteruskan"
		end if
	    if bulan3 <> "" then
		msg="Rekod baru untuk id <b>"& id &"</b> , bulan <b> '"&bulan3&"'</b>  ,tahun <b>"& tahun &"</b> telah wujud.<br>Proses tidak dapat diteruskan"
		end if
		if bulan4 <> "" then
		msg="Rekod baru untuk id <b>"& id &"</b> , bulan <b> '"&bulan4&"'</b>  ,tahun <b>"& tahun &"</b> telah wujud.<br>Proses tidak dapat diteruskan"
		end if
		if bulan5 <> "" then
		msg="Rekod baru untuk id <b>"& id &"</b> , bulan <b> '"&bulan5&"'</b>  ,tahun <b>"& tahun &"</b> telah wujud.<br>Proses tidak dapat diteruskan"
		end if
		if bulan6 <> "" then
		msg="Rekod baru untuk id <b>"& id &"</b> , bulan <b> '"&bulan6&"'</b>  ,tahun <b>"& tahun &"</b> telah wujud.<br>Proses tidak dapat diteruskan"
		end if
		if bulan7 <> "" then
		msg="Rekod baru untuk id <b>"& id &"</b> , bulan <b> '"&bulan7&"'</b>  ,tahun <b>"& tahun &"</b> telah wujud.<br>Proses tidak dapat diteruskan"
		end if
		if bulan8 <> "" then
		msg="Rekod baru untuk id <b>"& id &"</b> , bulan <b> '"&bulan8&"'</b>  ,tahun <b>"& tahun &"</b> telah wujud.<br>Proses tidak dapat diteruskan"
		end if
		if bulan9 <> "" then
		msg="Rekod baru untuk id <b>"& id &"</b> , bulan <b> '"&bulan9&"'</b>  ,tahun <b>"& tahun &"</b> telah wujud.<br>Proses tidak dapat diteruskan"
		end if
		if bulan10 <> "" then
		msg="Rekod baru untuk id <b>"& id &"</b> , bulan <b> '"&bulan10&"'</b>  ,tahun <b>"& tahun &"</b> telah wujud.<br>Proses tidak dapat diteruskan"
		end if
		if bulan11 <> "" then
		msg="Rekod baru untuk id <b>"& id &"</b> , bulan <b> '"&bulan11&"'</b>  ,tahun <b>"& tahun &"</b> telah wujud.<br>Proses tidak dapat diteruskan"
		end if
		if bulan12 <> "" then
		msg="Rekod baru untuk id <b>"& id &"</b> , bulan <b> '"&bulan12&"'</b>  ,tahun <b>"& tahun &"</b> telah wujud.<br>Proses tidak dapat diteruskan"
		end if
	end if
		
end if
%>   
         

    
<form name="test0" method="post" action="sp13.asp" onSubmit="return check(this)">  
<TABLE border=1 borderColor=black cellPadding=1 cellSpacing=0 rules=all width="48%" align="center"> 
    <tr class="hd"> 
    <td align="center" class="hd" colspan="2">Daftar Kelulusan Tuntutan Melebihi Sepertiga</td>
    </tr>
    <tr>
    <td align="center" colspan="5"><font size="2"><b><a href="../sispot/sp13a.asp" target="_new">Semakan dan Kemaskini Data</a></b></font></td>
   </tr>
  
    <tr> 
    <td height="64" align="center" bgcolor="E4F3F3"> No Pekerja : </td>
    <td valign="center" bgcolor="F4FAFA">&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="id" size="6" maxlength="6">
     <!--
    <a href="javascript:;" onclick="winBRopen('cal_popup.asp?FormName=test0&FieldName=tkh_bayar_gaji&Date=<%=Date()%>&CurrentDate=<%=Date()%>
     ','popup_cal','241','206','no','no')"><img src="images/icon_pickdate.png" class="DatePicker" alt="Pilih Tarikh" /></a>	-->
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    </td>
    </tr>
    
     
    <tr>
    <td width="30%" valign="middle" align="center" bgcolor="E4F3F3">Pilih Bulan :</td>
    <td width="70%" align="left"  bgcolor="F4FAFA">
    <li style="margin-top: 20px">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="all" id="all" /><label for='all'><b>Pilih Semua</b></label>
    <ul>
    <li>&nbsp;&nbsp;&nbsp;<input type="checkbox" name="bulan1" id="1" value="1" /> <label for="1">Jan</label>&nbsp;&nbsp;&nbsp;&nbsp;
    <input type="checkbox" name="bulan2" id="2" value="2" /> <label for="2">Feb</label>&nbsp;&nbsp;&nbsp;&nbsp;
    <input type="checkbox" name="bulan3" id="3" value="3" /> <label for="3">Mac</label>&nbsp;&nbsp;&nbsp;
    <input type="checkbox" name="bulan4" id="4" value="4" /> <label for="4">April</label>&nbsp;&nbsp;
    <input type="checkbox" name="bulan5" id="5" value="5" /> <label for="5">Mei</label>&nbsp;&nbsp;
    <input type="checkbox" name="bulan6" id="6" value="6" /> <label for="6">Jun</label></li>
    <li>&nbsp;&nbsp;&nbsp;<input type="checkbox" name="bulan7" id="7" value="7" /> <label for="7">Julai</label>&nbsp;&nbsp;&nbsp;
    <input type="checkbox" name="bulan8" id="8" value="8" /> <label for="8">Ogos</label>&nbsp;&nbsp;
    <input type="checkbox" name="bulan9" id="9" value="9" /> <label for="9">Sept</label>&nbsp;&nbsp;
    <input type="checkbox" name="bulan10" id="10" value="10" /> <label for="10">Okt</label>&nbsp;&nbsp;&nbsp;
    <input type="checkbox" name="bulan11" id="11" value="11" /> <label for="11">Nov</label>&nbsp;&nbsp;
    <input type="checkbox" name="bulan12" id="12" value="12" /> <label for="12">Dis</label></li>
    </ul>
    </li><br>
    </td>
    </tr>
    
     
    <tr valign="middle"> 
    <td align="center" bgcolor="E4F3F3">Tahun :</td>
    <td valign="middle" bgcolor="F4FAFA"><br>&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="tahun" size="4" maxlength="4">&nbsp;&nbsp;
    <!--
    <a href="javascript:;" onclick="winBRopen('cal_popup.asp?FormName=test0&FieldName=tkh_bayar_gaji&Date=<%=Date()%>&CurrentDate=<%=Date()%>
     ','popup_cal','241','206','no','no')"><img src="images/icon_pickdate.png" class="DatePicker" alt="Pilih Tarikh" /></a>	-->  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    </tr>
    </table>
    <br><p align="center"><input type="submit" name="proses" value="Simpan"></p>
    </form>
    <p align="center"><font color="#CC0000"><b><%=msg%></b></font></p>  
    
                
<% 
end sub 
%>


<%
sub papar1()
id=request("id")
bulan=request("bulan")
tahun=request("tahun")


    'check jabatan by id login
    pekd = request.cookies("gnop")
    a =     "select jabatan from majlis.pengguna where no_pekerja = '"& pekd &"' "
    Set rsa = objConn.Execute(a)


    'set bulan/tahun - arahan kerja 
    q4="select distinct to_char(tkh_ot,'mm/yyyy') as tkh_ot"
    q4 = q4 & " from payroll.jadual_ot order by tkh_ot"  
    set oq4 = objConn.Execute(q4)
    tkh_ot2 = oq4 ("tkh_ot")

    
	'data pengguna login
    idd = "select a.nama, a.lokasi, b.keterangan , to_char(sysdate,'dd/mm/yyyy') tkh  from payroll.paymas a, iabs.jabatan b "
    idd = idd & " where a.no_pekerja = '"&id&"' and b.kod = a.lokasi "
    set idd2 = objConn.Execute(idd)
    lokasi = idd2 ("lokasi")
	ket = idd2 ("keterangan")
	
	
   'papar data id yang sedang dikemaskini sahaja ---> nadia (24032016)
   q1="select a.rowid ,a.id as id , a.tahun as tahun, a.bulan as bulan from payroll.selenggara_kerja a , payroll.paymas b where a.id=b.no_pekerja "
   q1 = q1 & " and b.lokasi = '"&lokasi&"' and a.id = '"&id&"' order by a.tahun , a.bulan , a.id  asc"
   Set rs1 = objConn.Execute(q1)
  ' response.write q1
   if not rs1.bof and rs1.eof then
   id = rs1 ("id")	
   end if
%>
<br>

     
<TABLE border=1 borderColor=black cellPadding=0 cellSpacing=0 rules=all class="hd">
  <TBODY>    
    <tr>
  	<td align="center" class="hd" colspan="5"><B>SENARAI KELULUSAN TUNTUTAN MELEBIHI SEPERTIGA<br>
      JABATAN <%=ket%> (<%=lokasi%>)</B></td>
    <td width="0%"></td>
  </tr>
  
  <tr>
  	<td width="2%"   align="center" class="hd"><B>Bil</B></td>
    <td width="25%"  align="center" class="hd"><B>Nama Pekerja</B></td>
    <td width="22%"  align="center" class="hd"><B>Jabatan</B></td>
    <td width="24%"  align="center" class="hd"><B>Bulan</B></td>
    <td width="27%"  align="center" class="hd"><B>Tahun</B></td>   
  </tr>
  </tbody> 
</table>


<%  
bil=0
Do While Not rs1.eof 
bil=bil + 1
id = rs1 ("id")


mm = "select nama , lokasi from payroll.paymas where no_pekerja='"& id & "' "
Set rsmm = objConn.Execute(mm)
nama = rsmm ("nama")
lokasi = rsmm ("lokasi")


nn = "select keterangan from iabs.jabatan where kod='"& lokasi & "' "
Set rsnn = objConn.Execute (nn)
ket = rsnn ("keterangan")

%>


<TABLE border=1 borderColor=black cellPadding=0 cellSpacing=0 rules=all class="hd">
  <TBODY> 
  <tr>
  	<td width="2%"  align="center"><%=bil%></td>
    <td width="25%" align="left">&nbsp;<%=rs1("id")%>-<%=rsmm("nama")%></td>
    <td width="22%" align="left">&nbsp;(<%=rsmm("lokasi")%>) - <%=ket%></td>
    <td width="24%" align="center"><%=rs1("bulan")%></td>
    <td width="27%" align="center"><%=rs1("tahun")%></td>
  </tr>
  
  <% 
  rs1.movenext
  loop 
  %>
  
</tbody> 
</table>
<br><br>
<% end sub %>
</body>
</html>