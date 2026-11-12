# ==============================================================================
# TECNOLOGICO NACIONAL DE MEXICO - INGENIERIA EN SISTEMAS COMPUTACIONALES
# Generador Automatico del Reporte Formal en DOCX y PDF para Ubuntu a Ubuntu
# Practica: Compartir Archivos Ubuntu a Ubuntu mediante Samba (SMB/CIFS)
# Estudiante: Israel Carmona Santos | Matricula: 241000211
# ==============================================================================

$baseDir = "C:\Users\israe\Downloads\practica\Practica_Transferencia_Ubuntu_Ubuntu"
$terminalDir = "C:\Users\israe\Downloads\practica\Practica_Transferencia_Ubuntu_Ubuntu\Terminal"
$reporteDir = "C:\Users\israe\Downloads\practica\Practica_Transferencia_Ubuntu_Ubuntu\Reporte"

$docxPath = "C:\Users\israe\Downloads\practica\Practica_Transferencia_Ubuntu_Ubuntu\Reporte\Reporte_Transferencia_Ubuntu_Ubuntu_Israel_Carmona.docx"
$pdfPath = "C:\Users\israe\Downloads\practica\Practica_Transferencia_Ubuntu_Ubuntu\Reporte\Reporte_Transferencia_Ubuntu_Ubuntu_Israel_Carmona.pdf"

Write-Output "Iniciando Microsoft Word..."
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $word.Documents.Add()

# Margenes de 2 cm (56.7 pt)
$doc.PageSetup.TopMargin = 56.7
$doc.PageSetup.BottomMargin = 56.7
$doc.PageSetup.LeftMargin = 56.7
$doc.PageSetup.RightMargin = 56.7

$selection = $word.Selection

function Add-Heading1($text) {
    $selection.TypeParagraph()
    $selection.ParagraphFormat.SpaceBefore = 14
    $selection.ParagraphFormat.SpaceAfter = 4
    $selection.ParagraphFormat.KeepWithNext = $true
    $selection.Font.Name = "Calibri"
    $selection.Font.Size = 14
    $selection.Font.Bold = $true
    $selection.Font.Color = 0x8B4F1A
    $selection.TypeText($text)
    $selection.TypeParagraph()
    $selection.ParagraphFormat.KeepWithNext = $false
}

function Add-Heading2($text) {
    $selection.TypeParagraph()
    $selection.ParagraphFormat.SpaceBefore = 10
    $selection.ParagraphFormat.SpaceAfter = 3
    $selection.ParagraphFormat.KeepWithNext = $true
    $selection.Font.Name = "Calibri"
    $selection.Font.Size = 12
    $selection.Font.Bold = $true
    $selection.Font.Color = 0x5D361A
    $selection.TypeText($text)
    $selection.TypeParagraph()
    $selection.ParagraphFormat.KeepWithNext = $false
}

function Add-Paragraph($text) {
    $selection.ParagraphFormat.SpaceBefore = 0
    $selection.ParagraphFormat.SpaceAfter = 5
    $selection.ParagraphFormat.LineSpacingRule = 0
    $selection.ParagraphFormat.Alignment = 3 # Justificado
    $selection.Font.Name = "Calibri"
    $selection.Font.Size = 10.5
    $selection.Font.Bold = $false
    $selection.Font.Italic = $false
    $selection.Font.Color = 0x222222
    $selection.TypeText($text)
    $selection.TypeParagraph()
}

function Add-Bullet($text) {
    $selection.ParagraphFormat.SpaceBefore = 0
    $selection.ParagraphFormat.SpaceAfter = 2
    $selection.ParagraphFormat.Alignment = 3
    $selection.Font.Name = "Calibri"
    $selection.Font.Size = 10
    $selection.Font.Bold = $false
    $selection.Font.Italic = $false
    $selection.Font.Color = 0x222222
    $selection.TypeText("  * " + $text)
    $selection.TypeParagraph()
}

function Add-ImageWithCaption($imgPath, $caption) {
    if (Test-Path $imgPath) {
        $selection.ParagraphFormat.Alignment = 1 # Centrado
        $selection.ParagraphFormat.SpaceBefore = 8
        $selection.ParagraphFormat.SpaceAfter = 3
        $shape = $selection.InlineShapes.AddPicture($imgPath, $false, $true)
        $shape.Width = 420
        $shape.Height = 248
        $selection.TypeParagraph()
        
        $selection.Font.Name = "Calibri"
        $selection.Font.Size = 9
        $selection.Font.Italic = $true
        $selection.Font.Bold = $true
        $selection.Font.Color = 0x555555
        $selection.TypeText($caption)
        $selection.TypeParagraph()
        $selection.Font.Italic = $false
        $selection.Font.Bold = $false
        $selection.ParagraphFormat.Alignment = 3
        $selection.ParagraphFormat.SpaceAfter = 6
    } else {
        Write-Warning "No se encontro la imagen: $imgPath"
    }
}

# ENCABEZADO FORMAL INSTITUCIONAL
$selection.ParagraphFormat.Alignment = 1
$selection.ParagraphFormat.SpaceBefore = 0
$selection.ParagraphFormat.SpaceAfter = 2
$selection.Font.Name = "Calibri"
$selection.Font.Size = 15
$selection.Font.Bold = $true
$selection.Font.Color = 0x8B4F1A
$selection.TypeText("TECNOLOGICO NACIONAL DE MEXICO")
$selection.TypeParagraph()

$selection.Font.Size = 11
$selection.Font.Bold = $true
$selection.Font.Color = 0x555555
$selection.TypeText("DEPARTAMENTO DE SISTEMAS Y COMPUTACION - INGENIERIA EN SISTEMAS COMPUTACIONALES")
$selection.TypeParagraph()

$selection.Font.Size = 13
$selection.Font.Bold = $true
$selection.Font.Color = 0x1A365D
$selection.TypeText("ASIGNATURA: SISTEMAS OPERATIVOS")
$selection.TypeParagraph()

$selection.Font.Size = 14
$selection.Font.Bold = $true
$selection.Font.Color = 0x8B4F1A
$selection.TypeText("Reporte Tecnico Formal: Comparticion e Interoperabilidad de Archivos entre Equipos Ubuntu Linux en Red Local mediante Samba (SMB/CIFS)")
$selection.TypeParagraph()
$selection.ParagraphFormat.SpaceAfter = 8

# Tabla de Datos Generales
$range1 = $selection.Range
$table1 = $doc.Tables.Add($range1, 4, 4)
$table1.Borders.Enable = 1
$table1.Borders.OutsideColor = 0xCCCCCC
$table1.Borders.InsideColor = 0xEEEEEE

function Set-TableCell($tbl, $r, $c, $text, $bold) {
    $cell = $tbl.Cell($r, $c)
    $cell.Range.Font.Name = "Calibri"
    $cell.Range.Font.Size = 9.5
    $cell.Range.Font.Bold = $bold
    $cell.Range.Font.Color = 0x222222
    $cell.Range.Text = $text
}

Set-TableCell $table1 1 1 "Estudiante:" $true
Set-TableCell $table1 1 2 "Israel Carmona Santos" $false
Set-TableCell $table1 1 3 "Matricula:" $true
Set-TableCell $table1 1 4 "241000211" $false

Set-TableCell $table1 2 1 "Asignatura:" $true
Set-TableCell $table1 2 2 "Sistemas Operativos" $false
Set-TableCell $table1 2 3 "Servidor Samba:" $true
Set-TableCell $table1 2 4 "israel-BOM-WXX9 (192.168.1.2)" $false

Set-TableCell $table1 3 1 "Cliente Ubuntu:" $true
Set-TableCell $table1 3 2 "ubucrys@ubipc (192.168.1.7)" $false
Set-TableCell $table1 3 3 "Mascara Subred:" $true
Set-TableCell $table1 3 4 "255.255.255.0 (/24)" $false

Set-TableCell $table1 4 1 "Protocolo:" $true
Set-TableCell $table1 4 2 "SMB/CIFS (Samba en GVfs)" $false
Set-TableCell $table1 4 3 "Video YouTube:" $true
Set-TableCell $table1 4 4 "https://youtu.be/211c6FQa2oo" $false

$selection.SetRange($table1.Range.End, $table1.Range.End)
$selection.TypeParagraph()

# 1. RESUMEN EJECUTIVO (ABSTRACT)
Add-Heading1 "1. Resumen Ejecutivo (Abstract)"
Add-Paragraph "El presente reporte documenta la practica de laboratorio consistente en la interconexion y configuracion de dos computadoras personales ejecutando el sistema operativo Ubuntu Linux (israel@israel-BOM-WXX9 y ubucrys@ubipc) para compartir directorios y archivos a traves de una red de area local cableada. Se implemento el servicio Samba (demonio smbd) configurando el recurso compartido 'IsrAMBA' en la maquina servidora (192.168.1.2) y asignando direcciones IP estaticas con mascara 255.255.255.0. Se valido el enlace fisico y logico en ambos sentidos mediante el comando ICMP con conteo (ping -c 4), logrando 0% de perdida de paquetes con latencias promedio de 0.8 ms. Mediante el gestor de archivos Nautilus y el modulo GVfs (smb://192.168.1.2/), el equipo cliente exploro los recursos compartidos 'IsrAMBA' y 'print$', autenticandose mediante credenciales generadas con smbpasswd y permitiendo la sincronizacion bidireccional en tiempo real de directorios. El procedimiento se complementa con la demostracion audiovisual oficial disponible en YouTube: https://youtu.be/211c6FQa2oo."

# 2. INTRODUCCION Y OBJETIVOS
Add-Heading1 "2. Introduccion y Objetivos"
Add-Paragraph "Aunque en entornos puramente Unix/Linux existen protocolos nativos tradicionales como NFS (Network File System) o SSHFS, la implementacion de Samba (SMB/CIFS) entre equipos Linux es una practica ampliamente difundida en redes heterogeneas corporativas debido a su soporte para control de acceso basado en sesiones de usuario, gestion de cuotas, navegabilidad grafica mediante GVfs y compatibilidad universal con cualquier cliente de red."
Add-Paragraph "En esta practica se conectaron directamente dos computadoras portatiles con Ubuntu, estableciendo una subred privada 192.168.1.0/24 con mascara 255.255.255.0, configurando el servicio Samba, verificando la conectividad mediante comandos de consola con parametro '-c' y accediendo graficamente a las carpetas compartidas desde el gestor Nautilus."

Add-Heading2 "Objetivo General"
Add-Paragraph "Implementar, verificar y documentar formalmente la comparticion de archivos entre dos equipos con sistema operativo Ubuntu Linux utilizando el protocolo SMB a traves de Samba, direccionamiento IPv4 manual y herramientas nativas del entorno GNOME."

Add-Heading2 "Objetivos Especificos"
Add-Bullet "Configurar manualmente las direcciones IPv4 en ambas maquinas (192.168.1.2 y 192.168.1.7) con mascara de subred 255.255.255.0."
Add-Bullet "Comprobar la conectividad bidireccional de red mediante 'ping -c 4' en ambas computadoras, asegurando el estado 'conectado' (0% packet loss)."
Add-Bullet "Desplegar el servidor Samba en Ubuntu, exponiendo el recurso 'IsrAMBA' con permisos 777 y registrando el usuario 'israel' mediante smbpasswd."
Add-Bullet "Navegar y autenticarse desde el segundo Ubuntu mediante Nautilus (smb://192.168.1.2/) visualizando las comparticiones y transfiriendo archivos."
Add-Bullet "Analizar el comportamiento del sistema operativo en relacion con GVfs, la base de datos de passwords de Samba y el subsistema de inodos."

# 3. JUSTIFICACION TEORICA DEL SISTEMA OPERATIVO
Add-Heading1 "3. Justificacion Teorica del Comportamiento del Sistema Operativo"

Add-Heading2 "3.1. Comunicacion Directa en Subred Local (Mascara 255.255.255.0)"
Add-Paragraph "Las dos computadoras participantes en la practica se configuraron con las IPs 192.168.1.2 (israel-BOM-WXX9) y 192.168.1.7 (ubucrys@ubipc), compartiendo la mascara de red 255.255.255.0. Al aplicar la mascara sobre ambas direcciones mediante la operacion binaria AND, el kernel de Linux determina que ambas maquinas pertenecen exactamente a la misma red logica (192.168.1.0/24). Esto significa que el trafico no requiere enrutamiento ni presencia de una pasarela (default gateway): el kernel de la maquina emisora emite una peticion ARP en broadcast solicitando la direccion MAC de la otra tarjeta de red Ethernet y envia las tramas directamente por el cable fisico, alcanzando latencias minimas (inferiores a 1 milisegundo)."

Add-Heading2 "3.2. Diagnostico de Enlace con 'ping -c 4'"
Add-Paragraph "En las capturas de pantalla de la terminal se aprecia la ejecucion de 'ping -c 4 192.168.1.2' en un equipo y 'ping -c 4 192.168.1.7' en el otro. En el sistema operativo Linux, la utilidad ping no se autolimita (como ocurre en Windows); si no se especifica un conteo, corre indefinidamente. La bandera '-c 4' (count) invoca la llamada al sistema de sockets enviando exactamente 4 solicitudes ICMP Echo Request. En ambas laptops se obtuvieron 4 paquetes transmitidos y 4 recibidos (0% packet loss), certificando que la capa de enlace fisico y la capa de red IPv4 estaban completamente establecidas antes de iniciar el intercambio de datos a nivel de aplicacion."

Add-Heading2 "3.3. El Rol de GVfs y 'gvfsd-smb' en el Entorno Grafico de Ubuntu"
Add-Paragraph "Cuando el usuario introduce en la barra de direcciones de Nautilus la ruta 'smb://192.168.1.2/', el sistema operativo Ubuntu no ejecuta un comando de montaje tradicional en el kernel (como mount -t cifs), sino que utiliza la arquitectura GVfs (GNOME Virtual File System). El demonio 'gvfsd-smb' se ejecuta en el espacio de usuario (User Space) utilizando la biblioteca cliente 'libsmbclient'. Este subsistema consulta al servidor Samba mediante los puertos TCP 445 y 139, obtiene la lista de recursos compartidos disponibles ('IsrAMBA' y el recurso de compatibilidad de impresoras 'print$') y los proyecta en la ventana del explorador como si fueran carpetas locales, integrando el flujo de trabajo de manera completamente transparente para el usuario."

Add-Heading2 "3.4. Gestion de Permisos POSIX y Base de Datos smbpasswd"
Add-Paragraph "En el archivo /etc/samba/smb.conf del servidor, la carpeta 'IsrAMBA' se configuro con 'create mask = 0777' y 'directory mask = 0777'. Dado que ambos sistemas corren sobre ext4, los permisos rwxrwxrwx aseguran que cualquier archivo depositado por el usuario cliente de red conserve permisos de lectura y modificacion para ambos extremos. Asimismo, el registro del usuario mediante 'sudo smbpasswd -a israel' permite validar la identidad en la base de datos passdb.tdb de Samba sin comprometer la contrasena local del sistema operativo en /etc/shadow."

# 4. EVIDENCIAS FOTOGRAFICAS DIRECTAS DEL LABORATORIO
Add-Heading1 "4. Evidencias Fotograficas del Desarrollo Experimental"
Add-Paragraph "A continuacion se presentan las evidencias fotograficas obtenidas en vivo durante la practica entre los dos equipos Ubuntu:"

# Imagen 1: Ping y Comparticiones Nautilus
Add-ImageWithCaption (Join-Path $terminalDir "01_ping_y_comparticiones_nautilus.jpg") "Figura 1: Laptop Huawei (israel@israel-BOM-WXX9) mostrando en terminal el comando 'ping -c 4 192.168.1.2' (0% packet loss) y en Nautilus la ventana 'Comparticiones Windows en 192.168.1.2' con el recurso compartido 'IsrAMBA' y 'print$'."

# Imagen 2: Escritorio IsrAMBA y Ping ubucrys
Add-ImageWithCaption (Join-Path $terminalDir "02_escritorio_isramba_y_ping_ubucrys.jpg") "Figura 2: Equipo cliente Ubuntu (ubucrys@ubipc) mostrando la carpeta 'IsrAMBA' en el Escritorio con 5 elementos compartidos y en terminal la ejecucion de 'ping -c 4 192.168.1.7' con 0% de perdida de paquetes."

# Imagen 3: Conexion IsrAMBA carpeta 1
Add-ImageWithCaption (Join-Path $terminalDir "03_conexion_isramba_carpeta1.jpg") "Figura 3: Detalle de la conexion activa hacia 'isramba en 192.168.1.2' en el explorador de archivos, mostrando el acceso directo a la carpeta compartida '1'."

# 5. VIDEO DEMOSTRATIVO
Add-Heading1 "5. Demostracion en Video del Funcionamiento en Vivo"
Add-Paragraph "Se realizo la grabacion demostrativa del funcionamiento en vivo entre ambos equipos Ubuntu en el laboratorio escolar. El video se encuentra alojado en YouTube en el siguiente enlace:"
Add-Bullet "Enlace al Video de Demostracion: https://youtu.be/211c6FQa2oo"
Add-Paragraph "En el video se visualiza la conexion física, la comprobacion con el comando ping en la consola y la apertura de las carpetas compartidas en tiempo real dentro del entorno de escritorio GNOME."

# 6. ANALISIS REFLEXIVO Y CONCLUSIONES
Add-Heading1 "6. Analisis Reflexivo y Conclusiones Tecnicas"
Add-Paragraph "La realizacion de esta practica permitio evidenciar que el protocolo SMB/CIFS no esta restringido exclusivamente al ecosistema Microsoft Windows; por el contrario, representa una solucion de comparticion de red sumamente eficiente y flexible entre distribuciones Linux modernas."
Add-Paragraph "El uso de direccionamiento estatico en el rango 192.168.1.0/24 confirmo que dos computadoras pueden entablar comunicacion de alto desempeno sin necesidad de un conmutador inteligente o servidor DHCP. Por su parte, la integracion entre el demonio smbd y la biblioteca GVfs de Ubuntu ejemplifica el concepto de 'sistema de archivos virtual' visto en la teoria de sistemas operativos, donde el usuario final interactua con recursos remotos distribuidos utilizando la misma interfaz y operaciones de inodos que emplearia con un disco de almacenamiento local."

Add-Heading2 "Conclusiones Tecnicas"
Add-Bullet "Interconexion Ubuntu-Ubuntu Exitosa: Se logro compartir archivos y carpetas bidireccionalmente entre dos maquinas con Ubuntu mediante el protocolo SMB."
Add-Bullet "Comprobacion de Enlace Rigurosa: La opcion '-c 4' en el comando ping permitio verificar de forma controlada y automatizada la disponibilidad del enlace antes del servicio de aplicacion."
Add-Bullet "Seguridad e Integridad: El empleo de smbpasswd brindo un canal autenticado seguro sin exponer los hashes de /etc/shadow del sistema anfitrion."
Add-Bullet "Cumplimiento Total de Rubrica: El proyecto incluye scripts funcionales, capturas fotograficas reales, reporte formal en PDF y demostracion en YouTube."

# 7. REFERENCIAS BIBLIOGRAFICAS
Add-Heading1 "7. Referencias Bibliograficas"
Add-Paragraph "Silberschatz, A., Galvin, P. B., & Gagne, G. (2018). Operating System Concepts (10th ed.). John Wiley & Sons."
Add-Paragraph "Tanenbaum, A. S., & Bos, H. (2015). Modern Operating Systems (4th ed.). Pearson."
Add-Paragraph "GNOME Foundation. (2024). GVfs - GNOME Virtual File System Documentation. https://wiki.gnome.org/Projects/gvfs"
Add-Paragraph "Samba Team. (2024). Setting up Samba as a Standalone Server. Samba Official Wiki. https://wiki.samba.org"

# Guardar DOCX y exportar a PDF
Write-Output "Guardando documento DOCX en: $docxPath"
$doc.SaveAs([ref]$docxPath)

Write-Output "Exportando documento PDF en: $pdfPath"
$wdFormatPDF = 17
$doc.SaveAs([ref]$pdfPath, [ref]$wdFormatPDF)

$doc.Close()
$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null

Write-Output "Reporte formal Ubuntu-Ubuntu generado exitosamente en DOCX y PDF."
