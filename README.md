# gsColorearNET
Biblioteca de clases para colorear código y convertirlo en formato HTML (para usar en sitios WEB)
<br> 
El lenguaje de código usado es Visual Basic .NET.<br>
<br>
<br>
La carpeta <b>actual</b> tiene el código del proyecto actual con las modificaciones.<br>
Lo que hay en la carpeta <b>source</b> son copias de las distintas versiones creadas. <b>Esta carpeta no la seguiré actualizando</b>.<br>
<br>

IMPORTANTE:
===========
Esta biblioteca está codificada para usar en .NET Standard 2.0 y debe seguir siendo así, con idea de que sea compatiuble con las versiones de .NET Framework 4.6.1 a 4.8.
<br>
Puedes usar este código en tus proyectos sin ninguna restricción, así como la DLL una vez compilada, la cual puedes descargarla o instalarla desde el paquete de NuGet:
https://www.nuget.org/packages/gsColorearNET/<br>
<br>

Uso de esta DLL y de las versiones anteriores:
==============================================

<br> 
Actualmente esta biblioteca (o sus variantes anteriores para .NET Framework y .NET Core 3.1) las utilizo personalmente (algunas están distribuidas en la red).<br>
En estas aplicaciones si no indico el nombre de la DLL de colorear será gsColorear para .NET Framework, en los casos que utilizo gsColorearNET es instalándola desde 
el paquete de <a href="https://www.nuget.org/packages/gsColorearNET/">NuGet</a>.<br>
gsColorearCodigo v1.0.3.21 para .NET Framework 2.0<br>
gsColorearCodigo v1.0.7.1 para .NET Framework 4.7.2<br>
gsColorearCodigo v1.0.8.6 para .NET Framework 4.8 (utiliza gsColorearNET) (1)<br>
gsEditor 2008 v1.0.7.0 para .NET Framework 4.7.2<br>
gsEditor 2008 v1.0.7.2 para .NET Framework 4.7.2 (utiliza gsColorearNET)<br>
Compilar y ejecutar VB v1.0.0.22 para .NET Framework 4.7.2 (utiliza gsColorearNET)<br>
<br> 
<br> 
Y en estas otras que actualmente estoy depurando o solo para uso personal:<br>
<br>Si no indico lo contrario, utiliza gsColorearNET.<br>
Compilar y ejecutar (.NET 5.0) v1.0.0.0 para .NET 5.0 Preview 8 (utiliza gsColorearCore)<br>
Compilar NETCore WinF v1.0.0.4 para .NET 5.0 Preview 8<br>
gsColorearCodigoNET v1.1.0.0 para .NET 5.0 Preview 8 (utilidad convertida para .NET 5.0 a partir de gsColorearCodigo v1.0.8.4)<br>
<br>
<br>
(1) Esta utilidad está publicada en GitHub: <a href="https://github.com/elGuille-info/gsColorearCodigo">gsColorearCodigo</a>.
<br>
<br> 
Guillermo<br>
<br>
<br>
<h2>Actualizaciones</h2>
v1.0.0.8 del 17 de septiemrbe de 2020<br>
Corregido un BUG que dejaba algún \f0 al final.<br>
<br>
v1.0.0.7 del 16 de septiembre de 2020<br>
Cambio la función vb.Split para que no quite las líneas vacías si no se indica expresamente.<br>
Corregido BUG que quitaba todas las líneas en blanco.<br>
<br>
<br>
Actualizado el 17 de septimbre de 2020 a las 13:17 GMT+2
