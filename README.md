# excel-vba-cancelaciones

# Automatización en Excel con VBA

Macros para procesar datos, realizar cruces y aplicar reglas de negocio en Excel para una Base de Cancelaciones de seguros

## Archivos utilizados

- Archivo principal: donde se ejecuta la macro
- BaseAgentes: información de agentes (consultas)
- BaseCancelaciones: validaciones y cruces

## Funcionalidades

- Inserción automática de columnas
- Cálculos (ej. conversión a miles)
- Cruces de información
- Uso de VLOOKUP
- Clasificación de datos
- Conversión de fórmulas a valores

## Uso

1. Abrir el archivo principal
2. Ajustar rutas en los códigos:

```vba
rutaBaseExterna = "C:\Ruta\BaseCancelaciones.xlsx"
rutaAgentes = "C:\Ruta\BaseAgentes.xlsx"
