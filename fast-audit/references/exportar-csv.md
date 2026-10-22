# Cómo exportar las oportunidades de cada CRM

Dale a la persona solo la sección de su CRM. Los nombres de los menús cambian un poco según la versión y el idioma: si algo no coincide, que busque el equivalente.

Para cualquier CRM vale lo mismo:

- Una fila por oportunidad (negocio, trato, deal).
- Mejor abiertas, ganadas y perdidas juntas. Las cerradas sirven para la tasa de cierre y los motivos de pérdida.
- Los últimos 12 meses bastan.
- Exportar el mismo día que se audita, o decir la fecha de la exportación.
- Correos y teléfonos de los contactos no hacen falta: si salen, se pueden quitar.

Columnas que conviene añadir a la vista antes de exportar: etapa, estado, propietario o comercial, importe, fecha de creación, fecha de la última actividad, fecha de la próxima actividad, fecha de cierre prevista, fecha del último cambio de etapa y motivo de pérdida.

## HubSpot

1. Ve a CRM > Negocios y cambia a la vista de tabla (no la de tablero).
2. Quita los filtros que dejen fuera negocios cerrados, o elige el pipeline que quieras revisar.
3. En "Editar columnas", añade: Fecha de creación, Fecha de la última actividad, Fecha de la próxima actividad, Fecha de cierre, Motivo de pérdida del negocio cerrado, Propietario del negocio, Etapa del negocio y Cantidad.
4. Pulsa "Exportar", elige formato CSV y exporta las propiedades de la vista.
5. HubSpot manda un correo con el enlace de descarga.

## Pipedrive

1. Ve a Tratos y cambia a la vista de lista.
2. En el filtro, elige todos los tratos (abiertos, ganados y perdidos).
3. Con el icono de columnas, añade: Fecha de la última actividad, Fecha de la próxima actividad, Fecha de cierre prevista, Último cambio de etapa, Razón de pérdida y Estado.
4. Pulsa los tres puntos junto al filtro y elige "Exportar resultados del filtro". Formato CSV.
5. Solo sale lo que tu usuario puede ver. Para el equipo entero, que exporte alguien con permisos de administrador.

## Zoho CRM

1. Ve al módulo Tratos con la vista de todos los tratos.
2. En el menú de acciones, elige Exportar y formato CSV. También se puede desde Configuración > Administración de datos > Exportar.
3. Si la vista lo permite, incluye: Fase, Propietario de trato, Importe, Fecha de cierre, Hora de creación, Hora de la última actividad y Motivo de pérdida.
4. La descarga llega en un .zip: hay que descomprimirlo y subir el CSV.

## Salesforce

1. Ve a Informes > Nuevo informe y elige Oportunidades.
2. Filtro: todas las oportunidades, creadas en los últimos 12 meses.
3. Columnas: Nombre de la oportunidad, Propietario de la oportunidad, Fase, Importe, Fecha de cierre, Fecha de creación, Última actividad, Duración de la fase y el motivo de pérdida si lo tenéis como campo.
4. Ejecuta el informe y pulsa Exportar. Elige "Solo detalles", formato CSV delimitado por comas y codificación UTF-8.

## Otro CRM

Cualquier exportación con una fila por oportunidad y las columnas de arriba sirve, aunque se llamen distinto. Si el script no reconoce alguna, se asigna a mano con `--map`.
