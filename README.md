# VCard a Google Contacts

Conversor de archivos .vcf (vCard) a CSV compatible con Google Contacts. Procesa y limpia tus contactos automáticamente, eliminando duplicados y normalizando datos.

## Características

- Convierte archivos vCard (.vcf) a formato CSV compatible con Google Contacts
- Elimina contactos duplicados automáticamente
- Normaliza nombres y datos de contacto
- Unifica múltiples entradas del mismo contacto
- Maneja múltiples teléfonos y correos electrónicos por contacto
- Compatible con direcciones y notas personalizadas
- Decodifica valores codificados en QUOTED-PRINTABLE
- Procesa caracteres especiales correctamente
- Genera estadísticas detalladas del proceso

## Requisitos

- Python 3.7 o superior
- No requiere dependencias externas (solo librerías estándar)

## Instalación

1. Descarga el script `vcf_to_google_contacts.py`
2. Coloca tu archivo .vcf en el mismo directorio (opcional)

## Uso

### Opción 1: Con argumentos de línea de comandos

```bash
python3 vcf_to_google_contacts.py entrada.vcf salida.csv
```

### Opción 2: Modo interactivo (valores por defecto)

```bash
python3 vcf_to_google_contacts.py
```

Por defecto busca `vcards_20260115_140358.vcf` como entrada y genera `salida.csv`

## Parámetros

- `entrada.vcf`: Archivo vCard a procesar (obligatorio)
- `salida.csv`: Nombre del archivo CSV de salida (obligatorio)

## Proceso de conversión

El script ejecuta 3 pasos principales:

1. **Parseo**: Lee y extrae todos los contactos del archivo .vcf
2. **Fusión**: Identifica y unifica contactos duplicados por nombre o teléfono
3. **Generación**: Crea el archivo CSV en formato Google Contacts

## Campos procesados

El conversor extrae y procesa los siguientes campos:

- Nombre completo (FN)
- Nombre y apellido estructurado (N)
- Teléfono (TEL)
- Correo electrónico (EMAIL)
- Organización (ORG)
- Dirección (ADR)
- Notas (NOTE)
- Foto (PHOTO)

## Salida

El archivo CSV generado incluye todas las columnas estándar de Google Contacts:

- Nombre, Apellido, Nombre del medio
- Teléfono (hasta 3 números)
- Correo electrónico
- Organización
- Dirección
- Notas
- Foto y otros campos personalizados

## Importar en Google Contacts

1. Accede a [Google Contacts](https://contacts.google.com)
2. Haz clic en "Importar"
3. Selecciona el archivo CSV generado
4. Confirma la importación

## Estadísticas

Al finalizar, el script muestra:

- Total de contactos procesados
- Número de duplicados eliminados
- Contactos únicos finales
- Total de teléfonos
- Total de correos electrónicos

## Ejemplo de ejecución

```
======================================================================
Conversor de vCard (.vcf) a Google Contacts CSV
======================================================================

Archivo de entrada: contactos.vcf
Archivo de salida: contactos_google.csv

Paso 1/3: Parseando archivo .vcf...
[OK] 150 contactos extraidos

Paso 2/3: Fusionando contactos duplicados...
[OK] 12 duplicados fusionados
[OK] 138 contactos unicos

Paso 3/3: Generando CSV compatible con Google Contacts...
CSV generado exitosamente: contactos_google.csv

======================================================================
ESTADISTICAS:
   - Contactos procesados: 150
   - Duplicados fusionados: 12
   - Contactos en CSV final: 138
   - Total telefonos: 182
   - Total emails: 145

Proceso completado exitosamente!
Puedes importar contactos_google.csv directamente en Google Contacts
======================================================================
```

## Notas técnicas

- El script ignora errores de codificación para mejorar compatibilidad
- Soporta líneas continuadas (RFC 5545)
- Detecta y decodifica automáticamente valores QUOTED-PRINTABLE
- Normaliza teléfonos eliminando caracteres especiales
- Agrupa contactos por nombre y teléfono para detectar duplicados
- Basado en especificación RFC 6350 (vCard 4.0)
- Compatible con vCard versiones 3.0 y 4.0

## Compatibilidad

### Plataformas soportadas

- Google Contacts (recomendado)
- Microsoft Outlook (después de importar CSV)
- Apple Contacts (macOS, iOS)
- Thunderbird
- Otras aplicaciones que acepten importación CSV

### Campos soportados

El script mantiene compatible los siguientes campos estándar:
- Nombre y apellido
- Múltiples teléfonos (hasta 3 en export)
- Múltiples correos (1 en export, extensible)
- Dirección principal
- Organización y cargo
- Notas y comentarios
- Foto de perfil

## Limitaciones conocidas

- Solo exporta 1 email por contacto (se puede extender modificando el código)
- Solo exporta 3 teléfonos por contacto (se puede extender)
- Solo exporta 1 dirección por contacto
- Campos personalizados del VCF se pierden (excepto los reconocidos)
- No mantiene fotos en formato PHOTO del CSV
- La detección de duplicados usa lógica simple (nombre y teléfono)
- No procesa por lotes (puede ser lento con archivos > 10MB)
- Sin soporte para múltiples vCards en una sola línea

## FAQ

### ¿Puedo procesar múltiples archivos .vcf?

No directamente. Como workaround, puedes:
1. Unir los archivos VCF manualmente: `cat archivo1.vcf archivo2.vcf > combinado.vcf`
2. Procesar el archivo combinado

### ¿Cómo exporto más de 3 teléfonos por contacto?

Edita la línea en `GoogleContactsCSV.generate()`:
```python
for i, phone in enumerate(contact.get('phones', [])[:3], 1):
```
Cambia `[:3]` por `[:5]` (o el número que necesites)

### ¿Por qué algunos caracteres se ven mal?

El archivo se genera con encoding UTF-8. Al importar en Excel o Google Sheets:
- Excel: Selecciona "Texto" como opciones de importación
- Google: Usa "Importar en una nueva tabla" en Sheets

### ¿Se pierden contactos sin nombre o teléfono?

Sí. El script filtra contactos vacíos en la línea:
```python
if contact['fn'] or contact['phones']:
    return contact
```
Esto asegura que se procesan solo contactos con datos válidos.

### ¿Cómo modifico la lógica de duplicados?

Edita la clase `ContactMerger.merge_duplicates()`. Por ejemplo, para detectar duplicados solo por email:
```python
for indices in email_index.values():
    if len(indices) > 1:
        # crear grupo
```

### ¿Qué hacer si el archivo .vcf no se procesa?

1. Verifica que el archivo está en UTF-8
2. Comprueba que tiene extensión .vcf
3. Asegúrate de que la ruta es correcta
4. Intenta con otro programa (como Evolution) para validar el VCF

### ¿Es seguro para datos sensibles?

Sí. El script:
- Solo procesa archivos locales
- No envía datos a internet
- No requiere conexión externa
- No almacena datos en caché

### ¿Puedo importar el CSV después nuevamente?

No recomendado, ya que reintroducirá duplicados. Si necesitas hacer cambios:
1. Harlos en Google Contacts
2. Exportar nuevamente si es necesario

## Extensibilidad

### Agregar nuevos campos

Para agregar un campo, por ejemplo "Website":

1. En `VCardParser._parse_vcard()`, agrega:
```python
elif field == 'URL':
    contact['website'] = decoded_value
```

2. En `GoogleContactsCSV.__init__()`, agrega a headers:
```python
'Website'
```

3. En `GoogleContactsCSV.generate()`, mapea el campo:
```python
'Website': contact.get('website', '')
```

### Cambiar estrategia de duplicados

Edita `ContactMerger.merge_duplicates()` para usar otros criterios como email o dirección.

### Exportar a otros formatos

Crea una nueva clase similar a `GoogleContactsCSV` que genere JSON, XML, etc.

## Solución de problemas

**Caracteres especiales se ven mal:**
- El archivo se genera con encoding UTF-8. Abre en Excel con la opción correcta.

**Contactos no se importan:**
- Verifica que el CSV tenga todas las columnas requeridas por Google
- Comprueba que los emails tengan el formato correcto

**Archivo .vcf no se procesa:**
- Asegúrate de que el archivo está en el mismo directorio
- Verifica que el archivo no está corrupto

**Error "FileNotFoundError":**
- Comprueba que el archivo existe y la ruta es correcta
- En Windows, usa rutas con barras inversas o entrecomilladoras

**Estado de acceso denegado:**
- En Linux/Mac: verifica permisos con `ls -l archivo.vcf`
- Intenta ejecutar sin sudo primero

## Autor

Héctor Reyes Armenta

## Licencia

MIT License - Libre para usar, modificar y distribuir
