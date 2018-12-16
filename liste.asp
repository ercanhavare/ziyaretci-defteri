<%@language="javascript"%>
<h2><center>Ziyaretçi Defteri</h2>
<table border=1 align="center">
<tr>
<th>Ad</th>
<th>Soyad</th>
<th>E-Posta</th>
<th>Mesaj</th>
</tr>
<%
fso=Server.CreateObject("Scripting.FileSystemObject")
dosya=fso.OpenTextFile(Server.Mappath("/uye.txt"),1)
do
{
ad=dosya.ReadLine(); soyad=dosya.ReadLine()
eposta=dosya.ReadLine(); mesaj=dosya.ReadLine()
%>
<tr>
<td><%=ad%></td>
<td><%=soyad%></td>
<td><%=eposta%></td>
<td><%=mesaj%></td>
</tr>
<%
}
while(dosya.AtEndOfStream==false)
dosya.Close()
%>
</table><p><a href="index.asp">Ana Sayfa...</a>
</center>