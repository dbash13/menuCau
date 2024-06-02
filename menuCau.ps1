<#
    ********************** MENU CAU **********************
    Script realizado por dbash13
    ******************************************************
    Version actual v1.0.0
#>


#Titulo ventana
$windowTitle = "Menu CAU v1.0.0"

#Vaciar variables globales previas
$global:dataOK = $null
$global:username = $null
$global:vdi = $null
$global:avd = $null
$global:displayName = $null
$global:enabled = $null
$global:groups = $null
$global:checkCompany = $null
$global:userType = $null


<#
    Definir variable de nombre de usuario.
    - En este script se tienen en cuenta que la cuenta de administrador del usuario,
        tiene en su inicio el prefijo "ADM-".
    Por ejemplo, usuario que lanza el script es DBASH y el usuario con privilegios en
    directorio activo es "ADM-DBASH"
#>
$global:user = "ADM-" + $env:USERNAME #Usuario ADM


# Cargar los ensamblados de Windows Forms, Drawing y DA
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module ActiveDirectory

# Crear la ventana principal
$form = New-Object System.Windows.Forms.Form
    $form.Size = New-Object System.Drawing.Size(560, 600)
    $form.Font = New-Object System.Drawing.Font("Arial", 8)
    $form.StartPosition = "CenterScreen" #Lanzar la ventana centrada en la pantalla
    $form.MaximizeBox = $false
    $form.FormBorderStyle = "Fixed3D" #Desactivar que el usuario modifique el tamaño



# Crear el menu principal
$menu = New-Object System.Windows.Forms.MenuStrip
    $menu.Font = New-Object System.Drawing.Font("Arial", 8)

# Crear elementos de menu
$toolsMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Herramientas")
$userMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Buscar DisplayName (F4)")
$updateUser = New-Object System.Windows.Forms.ToolStripMenuItem("Actualizar (F5)")
$settingsMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Ajustes")
$separator1 = New-Object System.Windows.Forms.ToolStripSeparator #separador 1
$separator2 = New-Object System.Windows.Forms.ToolStripSeparator #separador 2

################
#Submenus de Herramientas
################
#Texto login ADM segun el estado (se llama en la funcion de Check-Condition)
$loginAdm = New-Object System.Windows.Forms.ToolStripMenuItem
$launchCmrc = New-Object System.Windows.Forms.ToolStripMenuItem("Abrir CMRC (F2)")
$launchDsa = New-Object System.Windows.Forms.ToolStripMenuItem("Abrir Active Directory (F3)")
$exitItem = New-Object System.Windows.Forms.ToolStripMenuItem("Salir")

# Agregar submenus a 'Herramientas'
$toolsMenu.DropDownItems.Add($loginAdm)
$toolsMenu.DropDownItems.Add($separator1)
$toolsMenu.DropDownItems.Add($launchCmrc)
$toolsMenu.DropDownItems.Add($launchDsa)
$toolsMenu.DropDownItems.Add($separator2)
$toolsMenu.DropDownItems.Add($exitItem)
#End

################
#Submenu ajustes
################
$alwaysOnTop = New-Object System.Windows.Forms.ToolStripMenuItem("Siempre encima") #Mantener ventana siempre encima
    $alwaysOnTop.CheckOnClick = $true
$showLogs = New-Object System.Windows.Forms.ToolStripMenuItem("Abrir logs") #Mostrar carpeta de logs

#ToDo, Actualmente sin configurar.    
#$showConsole = New-Object System.Windows.Forms.ToolStripMenuItem("Mostrar consola") 
#   $alwaysOnTop.CheckOnClick = $true

#Argregar al submenu de ajustes
$settingsMenu.DropDownItems.Add($alwaysOnTop)
$settingsMenu.DropDownItems.Add($showLogs)
#End

# Agregar elementos de menu al menu principal
$menu.Items.Add($toolsMenu)
$menu.Items.Add($userMenu)
$menu.Items.Add($updateUser)
$menu.Items.Add($settingsMenu)

# Agregar el menu a la ventana
$form.MainMenuStrip = $menu
$form.Controls.Add($menu)

#################################################################################
#            KeyBindings
#################################################################################

# F1: Lanzar login pulsando F1 únicamente si no se ha iniciado sesion
$form.Add_KeyDown({
    param($sender, $e)
    if (($e.KeyCode -eq [System.Windows.Forms.Keys]::F1) -and (!$global:credentials))
    {$loginAdmSB.Invoke()}
})

# F2: Lanzar CMRC VIewer
$form.Add_KeyDown({
    param($sender, $e)
    if (($e.KeyCode -eq [System.Windows.Forms.Keys]::F2) -and ($global:credentials))
    {$launchCmrcSB.Invoke()}
})

# F3: Lanzar DSA
$form.Add_KeyDown({
    param($sender, $e)
    if (($e.KeyCode -eq [System.Windows.Forms.Keys]::F3) -and ($global:credentials))
    {$launchDsaSB.Invoke()}
})

# F4: Lanzar ventana de busqueda de usuario
$form.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::F4)
    {$userMenuSB.Invoke()}
})


# F5: Actualizar datos pulsando F5
$form.Add_KeyDown({
    param($sender, $e)
    if (($e.KeyCode -eq [System.Windows.Forms.Keys]::F5) -and ($global:displayName))
        {$updateUserSB.Invoke()}
})

# F7: Cambiar contraseña de usuario
$form.Add_KeyDown({
    param($sender, $e)
    if (($e.KeyCode -eq [System.Windows.Forms.Keys]::F7) -and ($global:displayName) -and ($global:credentials))
        {$buttonChangePasswordSB.Invoke()}
})


# F8: Desbloquear usuario
$form.Add_KeyDown({
    param($sender, $e)
    if (($e.KeyCode -eq [System.Windows.Forms.Keys]::F8) -and ($global:displayName) -and ($global:credentials))
        {$buttonUnlockUserSB.Invoke()}
})
#################################################################################
#End
#################################################################################


#################################################################################
#            Textos y cuadros de texto en form
#################################################################################

############################
#Seccion de info basica del suaurio
############################
#Mostrar texto principal
$labelUserInfo = New-Object System.Windows.Forms.Label
    $labelUserInfo.Text = "Informacion del usuario: " + ""
    $labelUserInfo.Location = New-Object System.Drawing.Point(10, 28)
    $labelUserInfo.Size = New-Object System.Drawing.Size(280,17)
    $labelUserInfo.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($labelUserInfo)


#Mostrar nombre
$labelName = New-Object System.Windows.Forms.Label
    $labelName.Text = "Nombre:"
    $labelName.Location = New-Object System.Drawing.Point(10, 50)
    $labelName.Size = New-Object System.Drawing.Size(58,15)
$form.Controls.Add($labelName)

$textBoxName = New-Object System.Windows.Forms.TextBox
    $textBoxName.Location = New-Object System.Drawing.Point(68, 47)
    $textBoxName.Size = New-Object System.Drawing.Size(210, 20)
    $textBoxName.ReadOnly = $true
$form.Controls.Add($textBoxName)

#Mostrar valor employeeID (En muchas empresas, el DNI)
$labelEmployeeID = New-Object System.Windows.Forms.Label
    $labelEmployeeID.Text = "DNI:"
    $labelEmployeeID.Location = New-Object System.Drawing.Point(322, 50)
    $labelEmployeeID.Size = New-Object System.Drawing.Size(34,15)
$form.Controls.Add($labelEmployeeID)

$textBoxEmployeeID = New-Object System.Windows.Forms.TextBox
    $textBoxEmployeeID.Location = New-Object System.Drawing.Point(360, 47)
    $textBoxEmployeeID.Size = New-Object System.Drawing.Size(140, 20)
    $textBoxEmployeeID.ReadOnly = $true
$form.Controls.Add($textBoxEmployeeID)

#Mostrar valor email
$labelEmail = New-Object System.Windows.Forms.Label
    $labelEmail.Text = "Email:"
    $labelEmail.Location = New-Object System.Drawing.Point(10, 90)
    $labelEmail.Size = New-Object System.Drawing.Size(49,15)
$form.Controls.Add($labelEmail)

$textBoxEmail = New-Object System.Windows.Forms.TextBox
    $textBoxEmail.Location = New-Object System.Drawing.Point(68, 87)
    $textBoxEmail.Size = New-Object System.Drawing.Size(210, 20)
    $textBoxEmail.ReadOnly = $true
$form.Controls.Add($textBoxEmail)

#Mostrar valor title
$labelTitle = New-Object System.Windows.Forms.Label
    $labelTitle.Text = "Cargo:"
    $labelTitle.Location = New-Object System.Drawing.Point(10, 130)
    $labelTitle.Size = New-Object System.Drawing.Size(49,20)
$form.Controls.Add($labelTitle)

$textBoxTitle = New-Object System.Windows.Forms.TextBox
    $textBoxTitle.Location = New-Object System.Drawing.Point(68, 128)
    $textBoxTitle.Size = New-Object System.Drawing.Size(210, 20)
    $textBoxTitle.ReadOnly = $true
$form.Controls.Add($textBoxTitle)

#Mostrar departamento
$labelDepartment = New-Object System.Windows.Forms.Label
    $labelDepartment.Text = "Dpto:"
    $labelDepartment.Location = New-Object System.Drawing.Point(10, 170)
    $labelDepartment.Size = New-Object System.Drawing.Size(49,20)
$form.Controls.Add($labelDepartment)

$textBoxDepartment = New-Object System.Windows.Forms.TextBox
    $textBoxDepartment.Location = New-Object System.Drawing.Point(68, 168)
    $textBoxDepartment.Size = New-Object System.Drawing.Size(210, 20)
    $textBoxDepartment.ReadOnly = $true
$form.Controls.Add($textBoxDepartment)

#Mostrar localizacion
$labelOffice = New-Object System.Windows.Forms.Label
    $labelOffice.Text = "Ubicacion:"
    $labelOffice.Location = New-Object System.Drawing.Point(285, 90)
    $labelOffice.Size = New-Object System.Drawing.Size(73,20)
$form.Controls.Add($labelOffice)

$textBoxOffice = New-Object System.Windows.Forms.TextBox
    $textBoxOffice.Location = New-Object System.Drawing.Point(360, 88)
    $textBoxOffice.Size = New-Object System.Drawing.Size(140, 20)
    $textBoxOffice.ReadOnly = $true
$form.Controls.Add($textBoxOffice)

#Mostrar fecha creacion cuenta
$labelCreated = New-Object System.Windows.Forms.Label
    $labelCreated.Text = "F.Creacion:"
    $labelCreated.Location = New-Object System.Drawing.Point(281, 130)
    $labelCreated.Size = New-Object System.Drawing.Size(78,20)
$form.Controls.Add($labelCreated)

$textBoxCreated = New-Object System.Windows.Forms.TextBox
    $textBoxCreated.Location = New-Object System.Drawing.Point(360, 128)
    $textBoxCreated.Size = New-Object System.Drawing.Size(140, 20)
    $textBoxCreated.ReadOnly = $true
$form.Controls.Add($textBoxCreated)

#Mostrar telefono oficina
$labelTelOficina = New-Object System.Windows.Forms.Label
    $labelTelOficina.Text = "Tel.Oficina:"
    $labelTelOficina.Location = New-Object System.Drawing.Point(281, 170)
    $labelTelOficina.Size = New-Object System.Drawing.Size(78,20)
$form.Controls.Add($labelTelOficina)

$textBoxTelOficina = New-Object System.Windows.Forms.TextBox
    $textBoxTelOficina.Location = New-Object System.Drawing.Point(360, 168)
    $textBoxTelOficina.Size = New-Object System.Drawing.Size(140, 20)
    $textBoxTelOficina.ReadOnly = $true
$form.Controls.Add($textBoxTelOficina)


############################
#Seccion de checks de usuario
############################
#Mostrar texto
$labelUserCheck = New-Object System.Windows.Forms.Label
    $labelUserCheck.Text = "Comprobaciones adicionales:"
    $labelUserCheck.Location = New-Object System.Drawing.Point(10, 200)
    $labelUserCheck.Size = New-Object System.Drawing.Size(230,15)
    $labelUserCheck.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($labelUserCheck)

#Mostrar si pertenece al colectivo de empleados
$labelEmpleado = New-Object System.Windows.Forms.Label
    $labelEmpleado.Text = "Pertenece a GRUPOXX:"
    $labelEmpleado.Location = New-Object System.Drawing.Point(10, 220)
    $labelEmpleado.Size = New-Object System.Drawing.Size(80,25)
    $labelEmpleado.Font = New-Object System.Drawing.Font("Arial",6)
$form.Controls.Add($labelEmpleado)

$textBoxCheckGroup1 = New-Object System.Windows.Forms.TextBox
    $textBoxCheckGroup1.Location = New-Object System.Drawing.Point(13, 248)
    $textBoxCheckGroup1.Size = New-Object System.Drawing.Size(60, 20)
    $textBoxCheckGroup1.ReadOnly = $true
$form.Controls.Add($textBoxCheckGroup1)


#Mostrar si pertenece a algun grupo de seguridad de VDI
$labelVDI = New-Object System.Windows.Forms.Label
    $labelVDI.Text = "VDI:"
    $labelVDI.Location = New-Object System.Drawing.Point(95, 226)
    $labelVDI.Size = New-Object System.Drawing.Size(80,20)
$form.Controls.Add($labelVDI)

$textBoxVDI = New-Object System.Windows.Forms.TextBox
    $textBoxVDI.Location = New-Object System.Drawing.Point(97, 248)
    $textBoxVDI.Size = New-Object System.Drawing.Size(60, 20)
    $textBoxVDI.ReadOnly = $true
$form.Controls.Add($textBoxVDI)



#Mostrar si pertenece a algun grupo de seguridad que concede acceso a puesto AVD
$labelAVD = New-Object System.Windows.Forms.Label
    $labelAVD.Text = "AVD:"
    $labelAVD.Location = New-Object System.Drawing.Point(180, 226)
    $labelAVD.Size = New-Object System.Drawing.Size(38,20)
$form.Controls.Add($labelAVD)

$textBoxAVD = New-Object System.Windows.Forms.TextBox
    $textBoxAVD.Location = New-Object System.Drawing.Point(182, 248)
    $textBoxAVD.Size = New-Object System.Drawing.Size(60, 20)
    $textBoxAVD.ReadOnly = $true
$form.Controls.Add($textBoxAVD)

#Mostrar si la cuenta esta activada
$labelEnabled = New-Object System.Windows.Forms.Label
    $labelEnabled.Text = "Activada:"
    $labelEnabled.Location = New-Object System.Drawing.Point(264, 226)
    $labelEnabled.Size = New-Object System.Drawing.Size(100,20)
$form.Controls.Add($labelEnabled)

$textBoxEnabled = New-Object System.Windows.Forms.TextBox
    $textBoxEnabled.Location = New-Object System.Drawing.Point(266, 248)
    $textBoxEnabled.Size = New-Object System.Drawing.Size(60, 20)
    $textBoxEnabled.ReadOnly = $true
$form.Controls.Add($textBoxEnabled)

#Mostrar compañia
$labelCompany = New-Object System.Windows.Forms.Label
    $labelCompany.Text = "Compañia:"
    $labelCompany.Location = New-Object System.Drawing.Point(10, 288)
    $labelCompany.Size = New-Object System.Drawing.Size(74,20)
$form.Controls.Add($labelCompany)

$textBoxCompany = New-Object System.Windows.Forms.TextBox
    $textBoxCompany.Location = New-Object System.Drawing.Point(97, 286)
    $textBoxCompany.Size = New-Object System.Drawing.Size(230, 20)
    $textBoxCompany.ReadOnly = $true
$form.Controls.Add($textBoxCompany)

#Mostrar permisos de acceso sobre IVANTI
$labelDetallePulse = New-Object System.Windows.Forms.Label
    $labelDetallePulse.Text = "Permisos de VPN:"
    $labelDetallePulse.Location = New-Object System.Drawing.Point(10, 320)
    $labelDetallePulse.Size = New-Object System.Drawing.Size(158,15)
    $labelDetallePulse.Font = New-Object System.Drawing.Font("Arial",7, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($labelDetallePulse)

$textBoxVpn = New-Object System.Windows.Forms.TextBox
    $textBoxVpn.Location = New-Object System.Drawing.Point(13, 340)
    $textBoxVpn.Size = New-Object System.Drawing.Size(143, 20)
    $textBoxVpn.ReadOnly = $true
$form.Controls.Add($textBoxVpn)




############################
#Seccion de estado de acceso
############################
#Mostrar texto de seccion
$labelAdditionalInfo = New-Object System.Windows.Forms.Label
    $labelAdditionalInfo.Text = "Estado de acceso del usuario:"
    $labelAdditionalInfo.Location = New-Object System.Drawing.Point(10, 370)
    $labelAdditionalInfo.Size = New-Object System.Drawing.Size(230,15)
    $labelAdditionalInfo.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($labelAdditionalInfo)

#Cuenta bloqueada
$labelAccountLocked = New-Object System.Windows.Forms.Label
    $labelAccountLocked.Text = "Cuenta bloqueada:"
    $labelAccountLocked.Location = New-Object System.Drawing.Point(10, 390)
    $labelAccountLocked.Size = New-Object System.Drawing.Size(77,40)
$form.Controls.Add($labelAccountLocked)

$textBoxAccountLocked = New-Object System.Windows.Forms.TextBox
    $textBoxAccountLocked.Location = New-Object System.Drawing.Point(90, 392)
    $textBoxAccountLocked.Size = New-Object System.Drawing.Size(60, 20)
    $textBoxAccountLocked.ReadOnly = $true
$form.Controls.Add($textBoxAccountLocked)

#Numero de intentos fallidos
$labelFailedLogons = New-Object System.Windows.Forms.Label
    $labelFailedLogons.Text = "Intentos fallidos:"
    $labelFailedLogons.Location = New-Object System.Drawing.Point(180, 390)
    $labelFailedLogons.Size = New-Object System.Drawing.Size(75,35)
$form.Controls.Add($labelFailedLogons)

$textBoxFailedLogons = New-Object System.Windows.Forms.TextBox
    $textBoxFailedLogons.Location = New-Object System.Drawing.Point(255, 392)
    $textBoxFailedLogons.Size = New-Object System.Drawing.Size(60, 20)
    $textBoxFailedLogons.ReadOnly = $true
$form.Controls.Add($textBoxFailedLogons)

#Fecha de ultimo intento fallido en DC registrado
$labelLastFailedLogon = New-Object System.Windows.Forms.Label
    $labelLastFailedLogon.Text = "Ultimo intento:"
    $labelLastFailedLogon.Location = New-Object System.Drawing.Point(335, 390)
    $labelLastFailedLogon.Size = New-Object System.Drawing.Size(63,30)
$form.Controls.Add($labelLastFailedLogon)

$textBoxLastFailedLogon = New-Object System.Windows.Forms.TextBox
    $textBoxLastFailedLogon.Location = New-Object System.Drawing.Point(405, 392)
    $textBoxLastFailedLogon.Size = New-Object System.Drawing.Size(90, 20)
    $textBoxLastFailedLogon.ReadOnly = $true
$form.Controls.Add($textBoxLastFailedLogon)

#Contraseña expirada
#Cuenta bloqueada
$labelPasswordExpired = New-Object System.Windows.Forms.Label
    $labelPasswordExpired.Text = "Contraseña expirada:"
    $labelPasswordExpired.Location = New-Object System.Drawing.Point(10, 430)
    $labelPasswordExpired.Size = New-Object System.Drawing.Size(80,40)
$form.Controls.Add($labelPasswordExpired)

$textBoxPasswordExpired = New-Object System.Windows.Forms.TextBox
    $textBoxPasswordExpired.Location = New-Object System.Drawing.Point(90, 432)
    $textBoxPasswordExpired.Size = New-Object System.Drawing.Size(60, 20)
    $textBoxPasswordExpired.ReadOnly = $true
$form.Controls.Add($textBoxPasswordExpired)




#################################################################################
#           Botones inferiores
#################################################################################
# Boton para cambiar la contraseña
$buttonChangePassword = New-Object System.Windows.Forms.Button
    $buttonChangePassword.Text = "Restablecer contraseña (F7)"
    $buttonChangePassword.Location = New-Object System.Drawing.Point(10,505)
    $buttonChangePassword.Size = New-Object System.Drawing.Size(115, 38)
$form.Controls.Add($buttonChangePassword)

# Boton desbloquear el usuario
$buttonUnlockUser = New-Object System.Windows.Forms.Button
    $buttonUnlockUser.Text = "Desbloquear (F8)"
    $buttonUnlockUser.Location = New-Object System.Drawing.Point(130,505)
    $buttonUnlockUser.Size = New-Object System.Drawing.Size(94, 38)
$form.Controls.Add($buttonUnlockUser)

# Boton para mostrar detalle de puestos virtuales asignados al usuario
$buttonInfoVDesktop = New-Object System.Windows.Forms.Button
    $buttonInfoVDesktop.Text = "P. Virtuales Asignados"
    $buttonInfoVDesktop.Location = New-Object System.Drawing.Point(230,505)
    $buttonInfoVDesktop.Size = New-Object System.Drawing.Size(94, 38)
$form.Controls.Add($buttonInfoVDesktop)



#################################################################################
#End
#################################################################################




#################################################################################
#            Funciones PWS
#################################################################################

<#
    Funcion de logging.
    La funcion sera llamada cada vez que sea necesario registrar un evento, ya sea
    para guardar trazabilidad de los errores como para mostrar en la consola un 
    historial de acciones realizadas.
    - Se guardara toda la informacion en cada llamada a un archivo TXT en appdata
    - El nombre del archivo contiene la fecha actual y se guardara en la carpeta logs
    - En cada llamada ademas de guardar el evento, lo escribira en la consola.
#>
#Eliminar logs viejos (mantener solo 5)
function Del-OldLogs {
    $logDirectory = (Get-LogPath | Split-Path)
    $files = Get-ChildItem -Path $logDirectory -Filter "log_*.txt" | Sort-Object LastWriteTime -Descending
    if ($files.Count -gt 5) {
        $files | Select-Object -Skip 5 | Remove-Item
    }
}

# Funcion para obtener la ruta del log
function Get-LogPath {
    $appDataLocalPath = $env:appdata
    #Guardar ruta de appdata
    $logPath = Join-Path -Path $appDataLocalPath -ChildPath "menuCau\logs"
    
    #Si no extiste la ruta , crearla
    if (-not (Test-Path -Path $logPath)) {
        New-Item -Path $logPath -ItemType Directory | Out-Null
    }
    $date = Get-Date -Format "yyyy-MM-dd" #Establecer formato de fecha para agregar a nombre
    return Join-Path -Path $logPath -ChildPath "log_$date.txt"
}
 
# Funcion para escribir mensajes en el log y en la consola.
function Write-LogMessage {
    param (
        [string]$message
    )
    Del-OldLogs
    $logFile = Get-LogPath
    $time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$time - $message"
    $logEntry | Out-File -FilePath $logFile -Append -Encoding default
    Write-Host $logEntry
}
#End




#Funcion principal de comprobar condiciones
function Check-Conditions{

    #ADM: Actualizar texto de submenu y de titulo de ventana
    if(!$global:credentials){
        $loginAdm.Text = "Iniciar sesion con ADM (F1)"
        $form.Text = $windowTitle + "  (Read-Only)" 
        $launchCmrc.Enabled = $false
        $launchDsa.Enabled = $false   
    }else{
        $loginAdm.Text = "Sesion iniciada como ADM-" + "$env:username"
        $form.Text = $windowTitle + "  (ADM)" 
        $launchCmrc.Enabled = $true
        $launchDsa.Enabled = $true   
    }
    
    #Botones inferiores ADM: Desactivar si NO ADM, cuenta desactivada o es usuario CM
    if(!$dataOK -or !$global:credentials -or !$global:enabled -or $global:checkCompany){
        $buttonChangePassword.Enabled = $false
        $buttonUnlockUser.Enabled = $false
    }else{
        $buttonChangePassword.Enabled = $true
        $buttonUnlockUser.Enabled = $true
       }

    #Botones principales: Deshabilitar si no hay usuario buscado
    if(!$global:displayName){
        $updateUser.Enabled = $false
    }else{
        $updateUser.Enabled = $true
    }

    #Botones inferiores: Ocultar o deshabilitar si no hay DisplayName
    if(($global:avd -or $global:vdi) -and ($global:dataOK)){
        $buttonInfoVDesktop.Enabled = $true
    }else{
        $buttonInfoVDesktop.Enabled = $false
    }

    #Texto: Mostrar DisplayName buscado
    $labelUserInfo.Text = "Informacion del usuario: " + "$global:displayName"

    #
    #Cambiar colores de texbox segun estados
    #

    #Si cuenta activada
    if($global:dataOK.Enabled){
        $textBoxEnabled.BackColor = "LightGreen"
    }elseIf(!$global:dataOK.Enabled -and $global:displayName){
        $textBoxEnabled.BackColor = "Red"
        Write-LogMessage "Atencion: El usuario se encuentra deshabilitado"
    }else{
        $textBoxEnabled.BackColor = ""
        }

    #Si cuenta bloqueada
    if($global:dataOK.LockedOut){
        $textBoxAccountLocked.BackColor = "#ee3b3b"
        Write-LogMessage "Atencion: El usuario tiene la cuenta bloqueada."
    }elseIf(!$global:dataOK.LockedOut -and $global:displayName){
        $textBoxAccountLocked.BackColor = "LightGreen"
    }else{
        $textBoxAccountLocked.BackColor = ""
        }

    #Si contraseña expirada
    if($global:dataOK.PasswordExpired){
        $textBoxPasswordExpired.BackColor = "#ee3b3b"
        Write-LogMessage "Atencion: El usuario tiene la contraseña expirada."
    }elseIf(!$global:dataOK.PasswordExpired -and $global:displayName){
        $textBoxPasswordExpired.BackColor = "LightGreen"
    }else{
        $textBoxAccountLocked.BackColor = ""
    }


    #Si cuenta pertenece a un grupo en concreto en DA, avisar de forma visual
    if($global:groups.name -contains "Insertar-Grupo-Aqui"){
        $textBoxCheckGroup1.BackColor = "#FFFFA07A"
    }else{
        $textBoxCheckGroup1.BackColor = ""
        }

    
    #Si usuario CM
    if($global:checkCompany -and $dataOK){
        $textBoxCompany.BackColor = "#ee3b3b"
        Write-LogMessage "Atencion: Se ha buscado un usuario de compañia no autorizada."
    }else{
        $textBoxCompany.BackColor = ""
    }
}



<#
    Guardar datos de usuario a variables configuradas.
    Se utiliza esta funcion principal para devolver todos los parametros y textos necesarios
    que deben de ser mostrados o utilizados a posterior en el script.
    Dicha funcion, llama a las posteriores y utiliza sus datos para condiciones.
    - Se le debe de pasar obligatoriamente el
#>
function Get-userData {  
    param(
        [Parameter(Mandatory, Position = 1)]
        [PSObject] $username
    )
    
    #Vaciar variables de todo el script en caso de que contenga valores previos
    $global:dataOK = $null
    $global:enabled = $null
    $global:vdi = $null
    $global:avd = $null
    $DisplayName = $null
    $enabled = $null
    $checkES = $null
    $company = $null
    $vdi = $null
    $avd = $null
    $checkCompany = $null


    #########################
    #Definiciones a variables
    #########################
    ##Listado de compañias que el CAU no gestiona 
    $listadoBlackListCompany = "Compañia excluida 1", "Compañia excluida 2", "Compañia excluida 3"
    #Listado grupos ES que el CAU si gestiona
    $listadoGruposEs = "Compañia permitida 1", "Compañia permitida 2", "Compañia permitida 3"
            

    <#
        Patrones de grupos de seguridad habituales de puestos virtuales.
        - Se crean variables objetos para AVD y VDI
        - Se crean variables objeto para que grupos en concreto excluir de ese patron posterirmente
            en el foreach.
    #>    
    #VDI
    $groupsVdi = @("*PATRON1*", "*PATRON2*", "*PATRON3*")
    $groupsExcVdi = @("*PATRONExcluir1")
    #AVD
    $groupsAvd = @("*PATRON1*")
    $groupsExcAvd = @("*PATRONExcluir1*")


    <#
        Crear variables objeto para grupos habituales de VPN.
    #>
    $groupListVPN1 = @("GrupoAccesoVPN1")
    $groupListVPN2 = @("GrupoAccesoVPN2")


    <#
        Llamar a funcion que lanza una ventana solicitando un usuario a buscar.
        - Dicho usuario, puede ser filtrado por su nombre (DisplayName) o por DNI (EmployeeID)
        - Posteriormente, se guardan a variable independiente para evitar llamar la funcion
            varias veces dentro de esta otra funcion.
    #>    
    $data = Get-UserValues $username
    $dataOK = $data
            
    <#
        Se llama a funcion para filtrar en servidores ES y EXT el dato introducido en el paso anterior.
    #>
    $DisplayName = $dataOK | Select-Object -ExpandProperty Name


    <#
        Se busca en DA los grupos del usuario filtrado y se almacena en una variable objeto.
        - El uso de esta funcion, nos proporciona todos los datos acerca del grupo asignado, ya
            que en caso contrario, con Get-ADUser solo obtenemos MemberOf y no contiene mas que 
            el nombre.
    #>
    #Inicializar variables objeto
    $groups = @() 
    $search_groups = @()
    #Utilizar ADPrincipalGroupMembership para obtener
    try{
    $search_groups = Get-userGroups $DisplayName 
    }
    catch{
        $search_groups = $null
    }
    $groups = $search_groups

            
    #Comprobar si la cuenta esta activada y guardar en booleano
    $enabled = $dataOK | Select-Object -ExpandProperty Enabled


    <#
        Comprobar si el usuario es una compañia permitida 
        - Se utiliza un patron definido al inicio de la funcion.
    #>
    $company = $dataOK.Company #Leer compañia de usuario a variable
    $checkES = $listadoGruposEs -contains $company 
    $checkCompany = $listadoBlackListCompany -contains $company
    
    #Registrar en log coincidencia
    if($checkCompany){Write-LogMessage "Atencion, se ha buscado un usuario de compañia no permitida"} 
    

    <#
        Comprobar si el user dispone de puestos virtuales teniendo en cuenta los grupos de 
        seguridad habituales utilizando un patron definido al inicio de la funcion.
        - Se utiliza un patron para cada tipo de puesto definido al inicio de la gestion.
    #>
    #VDI
    $vdi = @()
    foreach ($pattern in $groupsVdi) {
        $vdi += ($groups.Name | Where-Object { $_ -like $pattern }) | Where-Object {$_ -NotContains $groupsExcVdi}
    }
    #AVD
    $avd = @()
    foreach ($pattern in $groupsAvd) {
        $avd += ($groups.Name | Where-Object { $_ -like $pattern }) | Where-Object {$_ -NotContains $groupsExcAvd}
    }
    
    
    <#
        Check VPN.
        Se revisara si en los grupos del usuario buscado se encuentra alguna coincidencia
        de los grupos habituales..
        - En el inicio de la funcion, se definen los grupos habituales dependiendo de si
            se trata del tipo de acceso 1 o tipo de acceso 2.
        - Posteriormente, se crea un bucle if para utilizar las condicionales y asi establecer
            el texto a mostrar en el script y que sea facil de analizar rapidamente.
    #>
    #Se comprueba array con los grupos habituales de Grupo 1
    foreach ($groupVPN in $groupListVPN1){
        $groupVpn1 += ($groups.Name | Where-Object { $_ -like $groupVPN })
    }

    #A continuacion, se comprueba para el array de grupos habituales de grupo 2
    foreach ($groupVPN in $groupListVPN2){
        $groupVpn2 += ($groups.Name | Where-Object { $_ -like $groupVPN })
    }
    
    #Utilizamos condicionales para estalecer el texto.
    if($groupVpn1 -and !$groupVpn2){
        $permisosVpn = "Si, acceso a VPN1"
    }elseif($groupVpn2 -and !$groupVpn1){
        $permisosVpn = "Si, acceso a VPN2"
    }elseif($groupVpn1 -and $groupVpn2){
        $permisosVpn = "Si, acceso a VPN 1 y 2"
    }else{
        $permisosVpn = "No"
    }
    
    
    <#
        Seccion de la funcion para guardar el texto a mostrar 
        y que es necesario almacenar previamente.
    #>
    #Texto en caso de que el usuario pertenezca a los valores "Company" excluidos
    $textBlockedCompanys = "Escribe aqui un texto a mostrar en ventana"
           
            
    #Devolver resultados, en formulario se pasaran a variables globales.
    return $DisplayName, $dataOK, $groups, $enabled, $checkES, $checkCompany, $vdi, $avd, $textBlockedCompanys, $permisosVpn

#End
}





<#
    Funcion para filtrar info de DisplayName o DNI introducido.
    En ella, se busca inicialmente en servidor serv1.empresa.com y en caso de error, se
    realiza la busqueda en servidor 2.
    - Se registrara en el log los eventos.
    - Se debe de llamar la funcion pasando un valor que puede ser un DisplayName (usuario)
        o un DNI.
#>
function Get-UserValues{
    param(
        [Parameter(Mandatory, Position = 1)]
        [PSObject] $username
                    )

        #Vaciar caches
        $data = $null
        $name = $null
        
        try{
            Write-LogMessage "Buscando datos en servidor 1"
            $data =  Get-ADUser -Server serv1.empresa.com -Filter {EmployeeID -eq $username -or Name -eq $username} -Properties *
            $global:userType = 1
        }catch{
            Write-LogMessage "Error al buscar el usuario en servidor serv1. $_"
        }
        if(!$data){
            Write-LogMessage "Buscando datos en servidor 2"
            $data = Get-ADUser -Server serv2.empresa.com -Filter {EmployeeID -eq $username -or Name -eq $username} -Properties *
            $global:userType = 2
            }
        return $data
#End
}




<#
    Obtener grupos de un usuario.
    Dado que de la llamada previa ya tenemos el dato que indica si el usuario es de
    servidor 1 o 2, lo utilizamos para reducir busquedas innecesarias y optimizar
    el tiempo que carga de datos a variable.
    - Es necesario llamarla pasandole un valor que debe de ser un DisplayName (usuario).
    - Se obtienen todos los datos a una variable de tipo objeto
#>
function Get-userGroups {
    param(
        [Parameter(Mandatory, Position = 1)]
        [PSObject] $DisplayName
    )

    if($global:userType -eq 1){
        Write-LogMessage "Obteniendo listado de grupos DA en servidor 1"
        $search_groups = Get-ADPrincipalGroupMembership -Server serv1.empresa.com $DisplayName
        return $search_groups 
    }
    if($global:userType -eq 2){
        Write-LogMessage "Obteniendo listado de grupos DA en servidor 2"
        $search_groups = Get-ADPrincipalGroupMembership -Server serv2.empresa.com $DisplayName
        return $search_groups
    }
    Write-LogMessage "Error al obtener grupos del usuario"
    return  $null
#End                                                                                        
}

             
<#
    Crear el cuadro de dialogo para ingresar el nombre del usuario a buscar.
    - Es necesario una funcion para utilizar el bucle while
#>
function Show-UserInputDialog {
    #Crear ventana formulario
    $inputForm = New-Object System.Windows.Forms.Form
        $inputForm.Text = "Busqueda"
        $inputForm.Size = New-Object System.Drawing.Size(320, 135)
        $inputForm.StartPosition = "CenterParent"
        $inputForm.FormBorderStyle = "Fixed3D"
        $inputForm.MaximizeBox = $false
        $inputForm.MinimizeBox = $false

    #Agregar texbox para guardar dato introducido posteriormente a var $username
    $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point(60, 18)
        $textBox.Size = New-Object System.Drawing.Size(180, 20)
        $inputForm.Controls.Add($textBox)

    $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "Buscar"
        $okButton.Location = New-Object System.Drawing.Point(110, 60)
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $inputForm.AcceptButton = $okButton
    $inputForm.Controls.Add($okButton)

    #SI se introduce un texto, devolver dato
    if ($inputForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    } else {
        #En caso de no introducir ningun dato, devolver variable vacia
        Write-LogMessage "Error, no has buscado ningun dato."
        return $null
    }
}


<#
    Funcion para llamar cada vez que sea necesario actualizar los texbox.
    Se debe de llamar tanto despues de buscar un DisplayName como al realizar cualquier accion
    sobre un usuario. Por ejemplo, puede ser necesario cada vez que se le cambie la contraseña al usuario
    para asi confirmar que se ha marcado como "contraseña expirada" o mismamente confirmar 
    si el cambio se ha realizado correctamente.
#>
function Update-CheckBox{
    try {
        $global:displayName, $global:dataOK, $global:groups, $global:enabled, $global:checkES, $global:checkCompany, $global:vdi, $global:avd, $global:textBlockedCompanys, $global:permisosVpn = Get-userData $username
        
        if ($global:displayName) {
            #Definir variable global
            $global:username = $username
            #Definir texbox nombre para mostrar
            $textBoxName.Text = $global:dataOK.DisplayName
            #Definir texbox DNI
            $textBoxEmployeeID.Text = $global:dataOK.EmployeeID
            #Definir texbox mail
            $textBoxEmail.Text = $global:dataOK.EmailAddress
            #Definir texbox Cargo
            $textBoxTitle.Text = $global:dataOK.Title
            #Definir texbox dpto
            $textBoxDepartment.Text = $global:dataOK.Department
            #Definir texbox de localizacion oficina
            $textBoxOffice.Text = $global:dataOK.Office
            #Definir texbox de f.creacion
            $textBoxCreated.Text = $global:dataOK.WhenCreated
            #Definir texbox de telefono
            $textBoxTelOficina.Text = $global:dataOK.OfficePhone
            
            #
            ## Seccion de "Comprobaciones adicionales"
            #
            #Check empleados
            $textBoxCheckGroup1.Text = if($global:groups.name -contains "GROUPXX"){"SI"}else{"NO"} 
            #Definir texbox de estado activacion cuenta
            $textBoxEnabled.Text = if($global:dataOK.Enabled){"SI"}else{"NO"}
            #Puestos virtuales ajuste temporal
            $textBoxAVD.Text = if(($global:avd).Count -gt 0){"SI"}else{"NO"}
            $textBoxVDI.Text = if(($global:vdi).Count -gt 0){"SI"}else{"NO"}
            #Definir texbox de compañia
            $textBoxCompany.Text = $global:dataOK.Company
            
            #Definir texbox de permisos sobre IVANTI (Pulse)
            $textBoxVpn.Text = $permisosVpn



            #
            # Seccion de bloqueos
            #
            #Definir textbox estado bloqueo
            $textBoxAccountLocked.Text = if($global:dataOK.LockedOut){"SI"}else{"NO"}
            #Definir texbox de numero de intentos fallidos
            $textBoxFailedLogons.Text = $global:dataOK.BadLogonCount
            #Definir texbox de ultimo intento fallido
            $textBoxLastFailedLogon.Text = $global:dataOK.LastBadPasswordAttempt
            #Definir texbox de contraseña expirada
            $textBoxPasswordExpired.Text = if($global:dataOK.PasswordExpired){"SI"}else{"NO"}

            #Test
            $global:userUPN = $global:dataOK.UserPrincipalName #Usos futuros
            $global:displayName = $global:dataOK.Name #Actualmente en uso para guardar nombre del usuario buscado


            #Llamar a funcion de condicion para comprobaciones adicionales
            Check-Conditions

            #Lanzar mensaje en pantalla en caso de que sea usuario CM
            if($global:checkCompany){
                Show-PopUpUserCm
            }

        } else {
                [System.Windows.Forms.MessageBox]::Show("DisplayName no encontrado",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
                Write-LogMessage "Error, usuario $username no encontrado"
            }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ha ocurrido un error.`n$_", #Mostrar el error con salto de linea. ¿C#?
        "Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error)
        Write-LogMessage "Ha ocurrido un error $_" #Mostrar el error con salto de linea 
    }
}


<#
    Funcion para solicitar la nueva contraseña del usuario.
#>
function Get-UserPassword {
    $newPassword = $null
    $passwordForm = New-Object System.Windows.Forms.Form
    $passwordForm.Text = "Restablecer contraseña"
    $passwordForm.Size = New-Object System.Drawing.Size(300, 150)
    $passwordForm.StartPosition = "CenterParent" #Lanzarla siempre centrada a la ventana principal
    $passwordForm.FormBorderStyle = "FixedDialog" #Desactivar modificacion de tamaño
    $passwordForm.MaximizeBox = $false
    $passwordForm.MinimizeBox = $false

    $labelPassword = New-Object System.Windows.Forms.Label
    $labelPassword.Text = "Nueva contraseña:"
    $labelPassword.Location = New-Object System.Drawing.Point(8, 15)
    $labelPassword.Size = New-Object System.Drawing.Size(90, 40)
    $passwordForm.Controls.Add($labelPassword)

    $textBoxPassword = New-Object System.Windows.Forms.TextBox
    $textBoxPassword.Location = New-Object System.Drawing.Point(120, 18)
    $textBoxPassword.Size = New-Object System.Drawing.Size(150, 20)
    $textBoxPassword.PasswordChar = '*'
    $passwordForm.Controls.Add($textBoxPassword)

    $okButtonPassword = New-Object System.Windows.Forms.Button
    $okButtonPassword.Text = "OK"
    $okButtonPassword.Location = New-Object System.Drawing.Point(100, 60)
    $okButtonPassword.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $passwordForm.AcceptButton = $okButtonPassword
    $passwordForm.Controls.Add($okButtonPassword)

    if ($passwordForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $newPassword = $textBoxPassword.Text
    }
    return $newPassword
#End
}


<#
    Mostrar ventana emergente en caso de que se haya buscado un usuario CM.
    - Actualmente esta funcion es llamada por la funcion de check conditions.
#>
function Show-PopUpUserCm{        
    #MostrarPopUp avisando
    [System.Windows.Forms.MessageBox]::Show("$global:textBlockedCompanys",
        "Atención",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Warning)
    }


<#
    Funcion para mostrar una lista de puestos virtuales de un usuario
    en una nueva ventana llamada desde el forms principal.
    - Pendiente ajustar correctamente para lograr su funcionalidad.
#>
function Show-ListVirtualDesktops {
    #Crear ventana 
    $formList = New-Object System.Windows.Forms.Form
    $formList.Text = "Detalle de puestos virtuales"
    $formList.Size = New-Object System.Drawing.Size(300,300)
    $formList.StartPosition = "CenterParent"
    $formList.Controls.Add($listBox)
    $formList.FormBorderStyle = "FixedDialog"
    $formList.MaximizeBox = $false
    $formList.MinimizeBox = $false

    #Crear tabla
    $pvTable = New-Object System.Windows.Forms.DataGridView
    $pvTable.Location = New-Object System.Drawing.Point(10,10)
    $pvTable.Size = New-Object System.Drawing.Size(255,200)
    $pvTable.AllowUserToAddRows = $false

    #Agregar columnas
    $pvTable.Columns.Add("1","2")
    $pvTable.Columns.Add("1","2")
    
    #Hacer split de saltos de linea
    $listAvd = $global:avd -split '\r?\n' | Where-Object {$_ }

    #Convertir variable AVD a objeto para procesarse correctamente
    $objectsAvd = foreach ($line in $listAvd){
        $valueAvd = ($line -split ",")[0] -replace "CN=", ""
        $objectAvd = [PSCustomObject]@{
            Nombre = $valueAvd
        }
      #Devolver resultado
      $objectAvd
    }

    $objectsAvd | format-table *


    $global:avd | format-table *
    foreach ($pvAvd in $global:avd){
        #$pvAvd = ($_ -split ",")[0] -replace "CN=", ""
        #$pvTable.Rows.Add($pvAvd.CN)
        
    }

    #$arrayAvd = @($global:avd | ForEach-Object {$_.Name})
    if($global:avd){
        $global:avd.ForEach({
            $pvName = ($_ -split ",")[0] -replace "CN=", ""
            Write-Host "$pvName"
        
            })
    
    }

    #Mostrar formulario en pantalla
    $formList.ShowDialog()

}

#################################################################################
#            Fin funciones
#################################################################################



#################################################################################
#            Funcionalidades y ScriptBlocks
#################################################################################

<#
    En esta seccion se definiran los 
#>



<#
    Funcionalidades menu principal "Herramientas"
#>

#Definicion de boton de Login con ADM y scriptblock
$loginAdmSB = {
    $global:crendentials = $null
    $global:user = "ADM-" + $env:USERNAME
    Write-LogMessage "Lanzando login..."
    $global:credentials = $host.ui.PromptForCredential("Solicitud de DisplayName ADM", "Por favor, introduce tu contraseña de ADM.", "$user", "DES000")
            
    #Comprobar que las credenciales sean validas y en caso negativo, vaciar credencial guardada para evitar bloqueos en cuenta
    try{
       $displayName = Get-ADUser -Server "serv1.empresa.com" -Identity $user -Credential $global:credentials -Properties GivenName | Select-Object -ExpandProperty GivenName
               
        if($displayName){
            Write-LogMessage "Inicio de sesion correcto $displayName"
        }
        

    }catch{
       Write-LogMessage "Error, contraseña incorrecta o falta de permisos."
       [System.Windows.Forms.MessageBox]::Show("Error, la contraseña no es correcta o no dispones de permisos.",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
       $global:credentials = $null
      }
    Check-Conditions
#End
}
$loginAdm.Add_Click($loginAdmSB)


<#
    # Lanzar CMRC Viewer.
    - Es necesario disponerlo en la ruta correcta, ya que en ese scipt
        se tienen en cuenta que esta en la caperta CmRcViewer de la raiz.
    - Muy importante respetar la ubicacion de la liberia .DLL en la carpeta 00000409
    - Dado que en este caso se requiere de otros privilegios para el acceso remoto,
        utiliza las credenciales de administrador previamente cacheadas.
#>
$launchCmrcSB = {
    try{
        Write-LogMessage "Lanzando CMRC..."
        Start-Process -WorkingDirectory $env:systemroot -FilePath 'C:\CmRcViewer\CmRcViewer.exe' -Credential $global:credentials
    }catch{
        Write-LogMessage "Error al lanzar aplicación. Por favor, comprueba que dispongas de la aplicacion en la ruta correcta. - C:\CmRcViewer\CmRcViewer.exe"
    }
    Write-LogMessage "Lanzado CMRC correctamente."   
}
$launchCmrc.Add_Click($launchCmrcSB)


<#
    Lanzar DSA.msc (Active directory)
    - Aplicacion proporcionada por MS si instalas el modulo de active directory en el equipo.
    - En este caso, se utiliza login de credenciales administrador cacheadas en variable.
#>
$launchDsaSB = {
    try{
        Write-LogMessage "Lanzando DSA (ActiveDirectory)..."
        Start-Process -WorkingDirectory $env:systemroot mmc dsa.msc -Credential $credentials #Iniciar el proceso DSA
    }catch{
        Write-LogMessage "Error al lanzar aplicacion"
        break
    }
    Write-LogMessage "Lanzada APP correctamente."
}
$launchDsa.Add_Click($launchDsaSB)


# Funcionalidad para el elemento 'Exit'
$exitItem.Add_Click({
    $form.Close()
})


<#
    FIN Funcionalidades menu principal "Herramientas
#>



# Funcionalidad para buscar y mostrar la informacion del usuario en Active Directory

$userMenuSB = {
    $username = Show-UserInputDialog
    if ($username) {
        Write-LogMessage "Se ha buscado el dato $username"
        # Buscar la informacion del usuario en Active Directory
        Update-CheckBox
        #Actualizar condiciones (botones)
        Check-Conditions
        }

#End
}
$userMenu.Add_Click($userMenuSB)


#Actualizar datos del usuario buscado pulsando boton (ver tambien keybinding)
$updateUserSB = {

    if ($global:displayName) {
        # Buscar la informacion del usuario en Active Directory
        try {Update-CheckBox
        #Actualizar condiciones (botones)
        Check-Conditions
        }catch{
            Write-LogMessage "Error al actualizar datos de usuario"
            break
        }
    }
    Write-LogMessage "Datos de usuario $global:displayName actualizados"
#End        
}
$updateUser.Add_Click($updateUserSB)

#Ventana siempre encima "Always On Top"
$alwaysOnTopSB = {
    if(!$form.TopMost){
        $form.TopMost = $true
    }else{
        $form.TopMost = $false
    }

}
$alwaysOnTop.Add_Click($alwaysOnTopSB)


#Boton "Abrir logs" de submenu Ajustes
$showLogsSB = {
    #Guardar ruta appdata de variable del sistema
    $appDataLocalPath = $env:appdata
    #Guardar ruta completa de appdata
    $logPath = Join-Path -Path $appDataLocalPath -ChildPath "menuCAU\logs"
    
    #Si no extiste la ruta , crearla
    if (-not (Test-Path -Path $logPath)) {
        New-Item -Path $logPath -ItemType Directory | Out-Null
    }
    
    Start-Process -FilePath $logPath
    
}
$showLogs.Add_Click($showLogsSB)


<#
    Funcionalidades botones inferiores
#>

# Funcionalidad para cambiar la contraseña del usuario
$buttonChangePasswordSB = {
    if ($global:displayName) {
        while ($true) {
            $newPassword = Get-UserPassword
            if (-not $newPassword) {
                break
            }

            try {
                # Intentar accion en servidor 1
                Set-ADAccountPassword -Server serv1.empresa.com -Identity $global:displayName -Credential $global:credentials -NewPassword (ConvertTo-SecureString -AsPlainText $newPassword -Force) -Reset
                Set-ADUser -Server serv1.empresa.com -Identity $global:displayName -Credential $global:credentials -ChangePasswordAtLogon $true
                [System.Windows.Forms.MessageBox]::Show("Contraseña cambiada correctamente.",
                    "Correcto",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information)
                    Write-LogMessage "Contraseña modificada correctamente al usuario $global:displayName"
                    Update-CheckBox
                break
            } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                try {
                    # Intentar accion en servidor 2
                    Set-ADAccountPassword -Server serv2.empresa.com -Identity $global:displayName -Credential $global:credentials -NewPassword (ConvertTo-SecureString -AsPlainText $newPassword -Force) -Reset
                    Set-ADUser -Server serv2.empresa.com -Identity $global:displayName -Credential $global:credentials -ChangePasswordAtLogon $true
                    [System.Windows.Forms.MessageBox]::Show("Contraseña cambiada correctamente.",
                        "Correcto",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information)
                        Write-LogMessage "Contraseña modificada correctamente al usuario $global:displayName"
                        Update-CheckBox
                    break
                } catch [Microsoft.ActiveDirectory.Management.ADPasswordComplexityException] {
                    [System.Windows.Forms.MessageBox]::Show("La contraseña no cumple con los requisitos de complejidad. Por favor, introduzca una nueva contraseña.",
                        "Error",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error)
                        Write-LogMessage "Error. La contraseña no cumple con los requisitos de complejidad para el usuario $global:displayName"
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Ha ocurrido un error. $_",
                        "Error",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error)
                        Write-LogMessage "Ha ocurrido un error. $_"
                    break
                }
            } catch [Microsoft.ActiveDirectory.Management.ADPasswordComplexityException] {
                [System.Windows.Forms.MessageBox]::Show("La contraseña no cumple con los requisitos de complejidad. Por favor, introduzca una nueva contraseña.",
                    "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error)
                    Write-LogMessage "Error al establecer contraseña, no cumple los requisitos de complejidad para el usuario $global:displayName"
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Ha ocurrido un error. $_",
                    "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error)
                break
            }
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No has seleccionado un usuario.",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    return
#End
}
$buttonChangePassword.Add_Click($buttonChangePasswordSB)




#Boton de desbloqueo de usuario
$buttonUnlockUserSB = {
    try {
        # Intentar acción en servidor 1
        Write-LogMessage "Intentando realizar desbloqueo del usuario $global:displayName en servidor serv1.empresa.com..."
        Unlock-ADAccount -Server serv1.empresa.com -Credential $global:credentials -Identity $global:displayName
        [System.Windows.Forms.MessageBox]::Show("Usuario desbloqueado correctamente en servidor serv1.empresa.com.",
            "Correcto",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information)
        Write-LogMessage "Desbloqueo correcto del usuario $global:displayName en servidor serv1.empresa.com"
        Update-CheckBox
    } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        try {
            # Intentar acción en servidor 1
            Write-LogMessage "Usuario $global:displayName no encontrado en servidor serv1.empresa.com, intentando en servidor serv2.empresa.com..."
            Unlock-ADAccount -Server serv2.empresa.com -Credential $global:credentials -Identity $global:displayName
            [System.Windows.Forms.MessageBox]::Show("Usuario desbloqueado correctamente en servidor serv2.empresa.com.",
                "Correcto",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information)
            Write-LogMessage "Desbloqueo correcto del usuario $global:displayName en servidor serv2.empresa.com"
            Update-CheckBox
        } catch {
            # En caso de error de algún tipo en el servidor 2, mostrar motivo
            [System.Windows.Forms.MessageBox]::Show("Ha ocurrido un error en servidor serv2.empresa.com: $_",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error)
            Write-LogMessage "Error en servidor serv2.empresa.com: $_"
        }
    } catch {
        # En caso de error de algún tipo en el servidor 1, mostrar motivo
        [System.Windows.Forms.MessageBox]::Show("Ha ocurrido un error en servidor serv1.empresa.com: $_",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        Write-LogMessage "Error en servidor serv1.empresa.com: $_"
    }
}
$buttonUnlockUser.Add_Click($buttonUnlockUserSB)


# Funcionalidad para listar puestos virtuales del usuario
$buttonInfoVDesktop.Add_Click({
    Show-ListVirtualDesktops
})



#################################################################################
#                Fin funcionalidades
#################################################################################




#################################################################################
#                Ajustes y configuraciones previo StartUp
#################################################################################

#Llamar a funcion de comprobar condiciones
Check-Conditions

#Descomentar para depuracion de errores
#pause

#Limpiar shell para mantener logs limpios
Clear-Host

#Guardar en log script lanzado
Write-LogMessage "Inicio del Script $windowTitle"

#Escuchar keybindings
$form.KeyPreview = $true


#################################################################################
#                Mostrar form
#################################################################################

$form.Add_Shown({ $form.Activate() })
[void] $form.ShowDialog()
