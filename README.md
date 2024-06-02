# Script Menu CAU
Se trata de un script PS1 de PowerShell que utiliza el embedido de Windows Forms (.NET Framework).

Ha sido realizado "just for learn" para ayudar en el trabajo diario a compañeros de un CAU (centro atencion al usuario) de soporte telefónico/web de una empresa. 

<img src="https://github.com/dbash13/menuCau/blob/main/images/launchScreen.png?raw=true" alt="Pantalla principal" style="width:350px;"/>

### Cuenta con las siguientes funcionalidades:
+ Busqueda de un usuario en servidores de directorio activo especificados por GivenName (usuario) o por DNI (EmployeeID).
+ Abrir aplicación de CmRcViewer (SCCM Remote Control) con usuario administrador.
+ Abrir aplicación Herramientas de directorio activo (DSA) con usuario administrador.
+ Cambio de contraseña de usuarios.
+ Desbloqueo de cuenta.
+ Detalle de puestos virtuales (si estos son agregados por grupo de seguridad en DA).
+ Guardado de logs de acciones realizadas para mantener trazabilidad de los errores.
+ Aviso visual en caso de que un usuario tenga agregado un valor en concreto en el campo "Company" de DA.

### Dispone de los siguientes ajustes visuales:
+ Avisos visuales en caso de que el usuario filtrado tenga la contraseña bloqueada, la contraseña expirada o que el usuario se encuentre deshabilitado.
+ Boton de ajustes para mantener la ventana del script siempre encima del resto de aplicaciones.

Dado que el script ha sido realizado teniendo en cuenta aspectos y configuraciones muy especificas de esta empresa, se han eliminado ciertos ajustes y dejado en blanco otros para su edición.

