# Set-Signatures

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Version](https://img.shields.io/badge/Version-2.0-orange.svg)](https://github.com/pablolgamarra/Set-Signatures/releases)

> Script de Powershell para configurar firmas de correo electrónico de usuarios a nivel de Tenant en Microsoft Office 365

## Características

**Set-Signatures** es un script de PowerShell que permite a los administradores de Microsoft 365 y Exchange Online configurar firmas de correo electrónico de forma centralizada para todos los usuarios de una organización.

### Funciones

- Configuración masiva de firmas para múltiples usuarios
- Configuración individual por usuario específico
- Plantillas HTML personalizables
- Integración con Microsoft Graph para obtener datos de usuarios
- Configuración automática en Outlook Web App (OWA)
- Sistema de logging detallado
- Estadísticas de ejecución
- Manejo de errores

## 🔧 Requisitos

### Requisitos del sistema

- **PowerShell 5.1** o (Recomendable [**Powershell 7**](https://learn.microsoft.com/es-es/powershell/scripting/whats-new/migrating-from-windows-powershell-51-to-powershell-7?view=powershell-7.5))
- **Windows 10/11** o **Windows Server 2016+**
- Conexión a Internet

### Módulos de PowerShell requeridos

El script utiliza diferentes modulos de Powershell, que verificará e instalará automáticamente en caso de que no se encuentren instalados en el sistema:

- `ExchangeOnlineManagement`
- `Microsoft.Graph`
- `Az.Accounts`

### Permisos necesarios

- **Microsoft Graph**: `User.Read.All`
- **Exchange Online**: Permisos de administrador para configurar buzones

## 📦 Instalación

### Opción 1: Descarga directa (Recomendado)

```powershell
# Descargar la última versión
Invoke-WebRequest -Uri "https://github.com/pablolgamarra/Set-Signatures/releases/latest/download/Set-Signatures.ps1" -OutFile "Set-Signatures.ps1"
```

### Opción 2: Clonar repositorio

```powershell
git clone https://github.com/pablolgamarra/Set-Signatures.git
cd Set-Signatures
```

## 🚀 Uso

### Sintaxis básica

```powershell
.\Set-Signatures.ps1 -BusinessName <string> -TemplatePath <string> [-UserMail <string>] [-SetOWA] [-LogPath <string>]
```

### Parámetros

| Parámetro | Tipo | Requerido | Descripción |
|-----------|------|-----------|-------------|
| `BusinessName` | String | ✅ Sí | Nombre de la organización para nombrar las firmas |
| `TemplatePath` | String | ✅ Sí | Ruta al archivo HTM de plantilla |
| `UserMail` | String | ❌ No | Email del usuario específico. Si se omite, procesa todos los usuarios |
| `SetOWA` | Switch | ❌ No | Habilita la configuración de firmas en OWA (default: `$false`) |
| `LogPath` | String | ❌ No | Ruta para guardar logs. Si se omite, solo muestra en consola |

### Ejemplos de uso

#### Ejemplo 1: Configurar firmas para todos los usuarios

```powershell
.\Set-Signatures.ps1 `
    -BusinessName "Acme Corp" `
    -TemplatePath "C:\Templates\signature.html" `
    -SetOWA
```

#### Ejemplo 2: Configurar firma para un usuario específico

```powershell
.\Set-Signatures.ps1 `
    -BusinessName "Acme Corp" `
    -TemplatePath "C:\Templates\signature.html" `
    -UserMail "usuario@acme.com" `
    -SetOWA
```

#### Ejemplo 3: Generar firmas con logging

```powershell
.\Set-Signatures.ps1 `
    -BusinessName "Acme Corp" `
    -TemplatePath "C:\Templates\signature.html" `
    -SetOWA `
    -LogPath "C:\Logs"
```

#### Ejemplo 4: Solo generar vista previa (sin aplicar)

```powershell
.\Set-Signatures.ps1 `
    -BusinessName "Acme Corp" `
    -TemplatePath "C:\Templates\signature.html"
```

## 📝 Plantilla HTML

### Estructura de la plantilla

La plantilla HTML debe contener los siguientes placeholders:

| Placeholder | Descripción | Ejemplo |
|-------------|-------------|---------|
| `{0}` | Nombre completo del usuario | Juan Pérez |
| `{1}` | Título del trabajo | Gerente Dpto. TI |
| `{2}` | Teléfono de negocio | +595 21 123456 |
| `{3}` | Email (para href) | <juan.perez@empresa.com> |
| `{4}` | Email (para mostrar) | <juan.perez@empresa.com> |

### Ejemplo de plantilla básica

```html
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Email Signature</title>
</head>
<body style="font-family: Arial, sans-serif; font-size: 10pt;">
    <table style="border-collapse: collapse;">
        <tr>
            <td>
                <strong style="font-size: 14pt; color: #1a1a1a;">{0}</strong><br>
                <span style="color: #666666;">{1}</span><br>
                <span style="color: #666666;">Tel: {2}</span><br>
                <a href="mailto:{3}" style="color: #0066cc;">{4}</a>
            </td>
        </tr>
    </table>
</body>
</html>
```

### 🎨 Plantilla de ejemplo incluida

El repositorio incluye una plantilla profesional en `templates/signature-example.html` que se puede usar como punto de partida, aunque con que el HTML personalizado de tu firma esté bien configurado, debería funcionar.

## 📊 Salida del script

### Mensajes de log

El script proporciona mensajes codificados por colores:

- 🟢 **SUCCESS**: Operaciones exitosas
- 🔵 **INFO**: Información general
- 🟡 **WARNING**: Advertencias
- 🔴 **ERROR**: Errores

### Resumen de ejecución

Al finalizar, el script muestra un resumen:

```powershell
========================================
         EXECUTION SUMMARY
========================================
Total Users Processed:    50
Successful:               48
Failed:                   2

Failed Users:
  - usuario1@empresa.com
  - usuario2@empresa.com
========================================
```

## 🔐 Autenticación

### Primera ejecución

La primera vez que ejecutes el script:

1. Se abrirá una ventana de autenticación de Microsoft
2. Inicia sesión con una cuenta de administrador
3. Acepta los permisos solicitados

### Ejecución automatizada

Para scripts automatizados, considera usar:

- **Service Principal** con certificado
- **Managed Identity** en Azure

## ⚠️ Consideraciones importantes

### Imágenes en plantillas

Las imágenes deben usar URLs absolutas (no rutas locales):

```html
<!-- ❌ NO funcionará -->
<img src="../images/logo.png">

<!-- ✅ SÍ funcionará -->
<img src="https://www.empresa.com/assets/logo.png">
```

### Límites de Microsoft 365

- Las firmas HTML tienen un límite de **30 KB**
- Las imágenes externas pueden ser bloqueadas por políticas de seguridad
- Algunos clientes de correo pueden no renderizar CSS complejo

### Roaming Signatures

El script configura automáticamente `PostponeRoamingSignaturesUntilLater` en el tenant para evitar conflictos con firmas locales.

## 🐛 Solución de problemas

### Error: "No se puede conectar a Microsoft Graph"

**Solución**: Verifica que tengas permisos de administrador y que el módulo `Microsoft.Graph` esté actualizado:

```powershell
Update-Module Microsoft.Graph -Force
```

### Error: "The HTML template does not contain the required placeholders"

**Solución**: Asegúrate de que tu plantilla HTML contenga los placeholders `{0}` a `{4}`.

### Error: "No users found with the specified criteria"

**Solución**:

- Verifica que el email del usuario sea correcto
- Confirma que el usuario tenga un buzón de Exchange Online
- Verifica que tengas permisos para leer usuarios

## 📚 Recursos adicionales

- [Documentación de Exchange Online PowerShell](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell)
- [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/overview)
- [Firmas de correo en Microsoft 365](https://docs.microsoft.com/en-us/microsoft-365/admin/setup/create-signatures-and-disclaimers)

## 🤝 Contribuir

Las contribuciones son bienvenidas. Por favor:

1. Fork el proyecto
2. Crea una rama para tu feature (`git checkout -b feature/AmazingFeature`)
3. Commit tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push a la rama (`git push origin feature/AmazingFeature`)
5. Abre un Pull Request

## 📄 Licencia

Este proyecto está bajo la Licencia MIT - ver el archivo [LICENSE](LICENSE) para más detalles.

## ✍️ Autor

**Pablo Gamarra**

- GitHub: [@pablolgamarra](https://github.com/pablolgamarra)
- LinkedIn: [Pablo Gamarra](https://www.linkedin.com/in/pablolgamarra)

## 🙏 Agradecimientos

- Comunidad de PowerShell
- Microsoft Exchange Team
- Todos los contribuidores

## 📋 Changelog

### [2.0] - 2024-12-02

#### Added

- Sistema de logging con niveles (INFO, WARNING, ERROR, SUCCESS)
- Estadísticas detalladas de ejecución
- Soporte para múltiples usuarios
- Validación de plantillas HTML
- Manejo robusto de errores por usuario
- Barras de progreso

#### Changed

- Mejorado el manejo de conexiones
- Optimizado el procesamiento de arrays
- Actualizada la documentación

#### Fixed

- Problema con objetos individuales vs arrays
- Error al procesar un solo usuario
- Manejo de teléfonos de negocio vacíos

### [1.3] - 2024-04-13

- Versión inicial funcional

---

<p align="center">
  Hecho con ❤️ por <a href="https://github.com/pablolgamarra">Pablo Gamarra</a>
</p>
