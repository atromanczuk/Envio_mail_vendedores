# Titulo del proyecto

Breve descripción del proyecto

## Datos de origen

- Cliente: *codigo y nombre del cliente destinatario del proyecto. Si es un producto usar el cliente 379-Inteligencia-y-Tecnologia-SA*
- Contrato: *codigo de contrato que da origen al proyecto*
- Solicitud: *Solicitud que dio origen al proyecto*

## Actualizaciones

### Actualizaciones con solicitud

Lista de solicitudes que impactan en la solución

- dd/mm/aaaa *solicitud*

### Actualizaciones sin solicitud

Lista de solicitudes que impactan en la solución

- dd/mm/aaaa *breve descripcion o link a changelog*

## Dependencias

*Librerias, scripts u otros artefactos que son necesarios para ejecutar la solución. Deben ser indicados con numero de versión.*

## Artefactos

*lista de funciones, store procedure, tablas, jobs y demas objetos sql que se implementaran*
*tambien deben incluirse ejecutables necesarios para el deploy y la implementación*

## Implementación

*Instrucciones para implementar en el cliente la solución utilizando los artefactos desplegados en el release*

## Desarrollo

### Deploy

*Instrucciones para que los desarrolladores generen un release del proyecto*

### Contenido

*Descripción del contenido, estructura de archivos, dependencias internas y otras consideraciones para editar y manipular el codigo*

/
|--doc/
|--gco/
|--sql/
|--test/
|--.gitignore
|--CHANGELOG.md
|--README.md

#### /doc/

Almacena archivos de documentación, imagenes descriptivas, snippets. Generalmente es requerido cuando la solución es muy grande o cuando el instructivo de implementación requiere ayuda visual.

#### /gco/

Contiene los archivos exportados por el Generador de Consultas listos para importar en el cliente. Cada archivo almacena la consulta y algunos aspectos de la configuración.

#### /sql/

Contiene los scripts sql que dan forma a la solución.

#### /test/

Contiene scripts para testear la solucion o con casos harcodeados.

### Testing

*Consideraciones para testear y depurar la solución*
