Aquí estarán los diferentes paquetes de NuGet de la versión Release de gsColorearNET.<br>
<br>
El paquete gsColorearNET.1.0.0.nupkg no es operativo ya que tiene un fallo en el código que hace que falle siempre.<br>
El código en realidad funciona bien, al menos si se usa el código como proyecto en Visual Studio, pero falla al instarlo como paquete NuGet<br>
Y es que utiliza un directorio con ficheros que contienen las instrucciones de los lenguajes a colorear.<br>
Pero se ve que al instalarlo y usarlo, no encuentra esa carpeta y producía un error.<br>
<br>
El paquete gsColorearNET.1.0.0.1.nupkg es operativo y corresponde a la revisión 1.0.0.2<br>
En este paquete he eliminado el directorio con los ficheros de lenguajes y en su lugar uso una colección con las palabras de cada uno de los lenguajes soportados.<br>
<br>
El paquete gsColorearNET.1.0.0.2.nupkg es operativo y corresponde a la revisión 1.0.0.4<br>
Este paquete tiene mejoras en el código aparte de otras asignaciones en las propiedades del paquete.<br>

