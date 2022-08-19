DECLARE @CODIGO_TAREA_PROGRAMADA INT = <<Ingresar número de tarea programada>>		--Codigo de la tarea programada que desea configurar

--FILTROS PARAMETRIZABLES DE LA TAREA PROGRAMADA
DECLARE @COMPROBANTES_DIVISION INT = 1			--Codigo de división

--Testing
--SELECT * FROM dmTPES_PARM WHERE PARM_CODIGO_TPRG = 19

-- PARAMETROS PARA CONFIGURAR ASPECTOS DEL DESTINATARIO
-- La tarea programada acepta el uso de tags para auto completar los campos del destinatario con informacion que se genera durante la construccion del mensaje
-- Los tags son cadenas de texto que se reemplazan por un valor que puede cambiar de destinatario a destinatario, o de mensaje a mensaje.
-- Cada aspecto de la configuracion del destinatario soporta un conjunto determinado de tags
-- TAGS DISPONIBLES PARA...
--	DESTINATARIOS_PARA
--	DESTINATARIOS_CC
DECLARE @ASUNTO VARCHAR(150) = 'ALTA DE PEDIDOS'
DECLARE @DESTINATARIOS_PARA VARCHAR(3000) = '{tag::mail-vendedor}'
DECLARE @DESTINATARIOS_CC VARCHAR(3000) = ''
DECLARE @DESTINATARIOS_CCO VARCHAR(3000) = ''
DECLARE @ANEXAR_CUERPO NVARCHAR(MAX) = '{tag::cuerpo-envio-comprobantes-facturas}'

--CONFIGURACION DEL REMITENTE
--El remitente puede provenir de varios origenes, dependiendo de la clase de remitente a utilizar, hay variables que no son requeridas.
--Existen tres clases de remitentes posibles:
--	INSTITUCIONAL, la cuenta remitente que se utilizara sera una cuenta institucional.
--	USUARIO PLATAFORMA, se utilizara como remitente la cuenta asignada a un usuario de plataforma
--	POR DOMINIO, se utilizara la cuenta configurada en el Dominio de Email dentro de los Parametros Generales De Mensajeria del Modulo Mensajeria de Plataforma
--
--En cualquiera de los tres casos, el sistema de mensajeria requiere que se indique un usuario de plataforma vinculado a los mensajes
--El usuario de plataforma que quedara ligado a los mensajes se puede indicar a mano (completando la variable @USUARIO_PLATAFORMA),
--o se puede vincular con el usuario responsable de la tarea programada (configurando en 1 la varaible @USAR_USUARIO_RESPONSABLE_DE_TPRG)
--Considerando estos elementos, la configuracion del remitente puede tomar diferentes variantes:
--  Usar cuenta institucional y usuario de plataforma DIFERENTE al responsable de la tarea programda
--		@USAR_USUARIO_RESPONSABLE_DE_TPRG = 0
--		@USUARIO_PLATAFORMA = 'codigo del usuario'
--		@CLASE_DE_REMITENTE = 'INSTITUCIONAL'
--		@REMITENTE = 'cuenta de correo del remitente'
--  Usar cuenta institucional y usuario de plataforma IGUAL al responsable de la tarea programda
--		@USAR_USUARIO_RESPONSABLE_DE_TPRG = 1
--		@USUARIO_PLATAFORMA = ''
--		@CLASE_DE_REMITENTE = 'INSTITUCIONAL'
--		@REMITENTE = 'cuenta de correo del remitente'
--  Usar cuenta de un usuario de plataforma y usuario de plataforma DIFERENTE al responsable de la tarea programda
--		@USAR_USUARIO_RESPONSABLE_DE_TPRG = 0
--		@USUARIO_PLATAFORMA = 'codigo del usuario'
--		@CLASE_DE_REMITENTE = 'USUARIO PLATAFORMA'
--		@REMITENTE = ''
--  Usar cuenta de un usuario de plataforma y usuario de plataforma IGUAL al responsable de la tarea programda
--		@USAR_USUARIO_RESPONSABLE_DE_TPRG = 1
--		@USUARIO_PLATAFORMA = ''
--		@CLASE_DE_REMITENTE = 'USUARIO PLATAFORMA'
--		@REMITENTE = ''
--  Usar cuenta de dominio y usuario de plataforma DIFERENTE al responsable de la tarea programda
--		@USAR_USUARIO_RESPONSABLE_DE_TPRG = 0
--		@USUARIO_PLATAFORMA = 'codigo del usuario'
--		@CLASE_DE_REMITENTE = 'POR DOMINIO'
--		@REMITENTE = 'cuenta de correo del dominio elegido'
--  Usar cuenta de dominio y usuario de plataforma IGUAL al responsable de la tarea programda
--		@USAR_USUARIO_RESPONSABLE_DE_TPRG = 1
--		@USUARIO_PLATAFORMA = ''
--		@CLASE_DE_REMITENTE = 'POR DOMINIO'
--		@REMITENTE = 'cuenta de correo del dominio elegido'

DECLARE @USAR_USUARIO_RESPONSABLE_DE_TPRG SMALLINT = 0	--Indica si el usuario de plataforma del mensaje es el mismo que el responsable de la tarea programada
DECLARE @USUARIO_PLATAFORMA VARCHAR(8) = 'INTEC'	--Usuario de plataforma vinculado a los mensajes. Es obligatorio, a menos que se indique usar usuario responsable de TPRG
DECLARE @CLASE_DE_REMITENTE VARCHAR(20)= 'POR DOMINIO'		--Origen del remitente. Acepta 3 clases, 'INSTITUCIONAL', 'USUARIO PLATAFORMA', 'POR DOMINIO'
--Testing:
--		INTEC.DAM@outlook.com
DECLARE @REMITENTE VARCHAR(3000) = ''


--PARAMETROS DE LA PLANTILLA
--Consisten en valores que modifican el contenido o algunos aspectos de la visualizacion de la plantilla
DECLARE @INTRODUCCION NVARCHAR(MAX) = 'Se cargaron los siguientes pedidos'
DECLARE @FIRMA NVARCHAR(MAX) = 'Saludos.'
DECLARE @URL_IMAGEN_MENSAJE_ENCABEZADO NVARCHAR(MAX) = ''
DECLARE @URL_IMAGEN_MENSAJE_PIE NVARCHAR(MAX) = ''
DECLARE @TIPOGRAFIA NVARCHAR(MAX) = 'Cambria, Cochin Georgia, Times, serif'
DECLARE @TIPOGRAFIA_ALTURA NVARCHAR(MAX) = '20px'

--COLUMNAS DE LA GRILLA PRINCIPAL
--La tarea programada genera una grilla con los registros a informar.
--En esta seccion es posible definir la estrucutra que tendra esa grilla dandole forma a cada columna
--La estructura basica que tienen que tener las columnas es
--
--<Columna>
--	<Encabezado>Titulo de la columna</Encabezado>
--	<Contenido>{data::algun campo valido para la tpes}</Contenido>
--	<Formato_Contenido>algun formato valido para la tpes</Formato_Contenido>
--	<Estilo_Encabezado>codigo CSS</Estilo_Encabezado>
--	<Estilo_Contenido>codigo CSS</Estilo_Contenido>
--</Columna>
--
--La columna se configura modificando el contenido de las etiquetas. Cada etiqueta modifica un aspecto de la columna:
--Encabezado,	indica el nombre que tendra la columna en el encabezado de la grilla
--Contenido,	indica el valor que tendra cada celda de la columna. Este valor por lo general depende de algun dato procesado por la tarea progrmada estandar
--				Por lo tanto, cada tarea programada estandar tiene un conjunto de contenidos validos para mostrar
--				Mas abajo se aclaran los contenidos permitidos para esta tarea en particular
--Formato_Contenido,	Es opcional. Se utiliza para modificar la apariencia del contenido a mostrar. Por ejemplo: una fecha podria mostrarse con varios formatos 01/01/2021, 01-01-2021, 2021-01-01
--						De forma similar a Contenido, el formato dependera de la tarea programada estandar.
--Estilo_Encabezado		Contiene un conjunto de atributos CSS que modifican el aspecto de la celda que contiene el encabezado de la columna en la grilla.
--Estilo_Contenido		Contiene un conjunto de atributos CSS que modifican el aspecto de las celdas de contenido que forman parte de la columna.
--
--El orden en que las columnas se declaran sera el mismo que se tomara en cuenta para la construccion de la grilla. Siendo la primer columna definida, la columna de la izquieda en la tabla.
--
--CONTENIDOS Y FORMATOS DISPONIBLES DE COLUMNAS PARA LA TPES
--
--{data::comprobante-emision}	Fecha de emision del comprobante.
--		Formato: 103			Muestra la fecha con el formato dd/mm/yyyy
--{data::vencimiento}			Fecha de vencimiento del slado.
--		Formato: 103			Muestra la fecha con el formato dd/mm/yyyy
--{data::comprobante}			Muestra el codigo completo del comprobante. Por defecto el formato es corto.
--		Formato: corto			dd-ss-ttt-nnn
--		Formato: largo			000d-000s-ttt-00000000n
--{data::division}				Muestra el numero de division del comprobante. Por defecto el formato es corto
--		Formato: corto			dd
--		Formato: largo			000d
--{data::sucursal}				Muestra el numero de sucursal del comprobante. Por defecto el formato es corto
--		Formato: corto			ss
--		Formato: largo			000s
--{data::tipo}					Muestra el codigo de tipo de comprobante
--{data::numero}				Muestra el numero de comprobante. Por defecto el formato es corto
--		Formato: corto			nnnn
--		Formato: largo			00000000n
--{data::saldo-origen}			Muestra saldo en moneda origen en formato decimal
--		Formato: moneda			ssss.00
--{data::saldo-local}			Muestra saldo en moneda local en formato decimal
--		Formato: moneda			ssss.00
--{data::comprobante-moneda}	muestra el codigo del tipo de moneda
--{data::comprobante-importe-origen}	Muestra el importe total en moneda origen del comprobante en formato decimal
--		Formato: moneda			ssss.00

DECLARE @COLUMNAS XML = '
<Columna>
	<Encabezado>
		Cod. Cliente
	</Encabezado>
	<Contenido>
		{data::cod-cliente}
	</Contenido>
	<Formato_Contenido>
		
	</Formato_Contenido>
	<Estilo_Encabezado>
		width:				100px;
		text-align:			center;
		background:			black;
		text-align:			center;
		font-weight:		bold;
		color:				rgb(255, 255, 255);
		background:			rgb(121, 183, 41);
	</Estilo_Encabezado>
	<Estilo_Contenido>
		padding-top:		0.4em;
		padding-bottom:		0.3em;
		padding-left: 0.5em;
		padding-right: 0.5em;
		border-top-color:	rgb(121, 183, 41);
		border-top-style:	solid;
		border-top-width:	1px;
		color:				rgb(0, 0, 0);
		background-color:	rgb(255, 255, 255);
		text-align:			center;
	</Estilo_Contenido>
</Columna>
<Columna>
	<Encabezado>
		Nombre cliente
	</Encabezado>
	<Contenido>
		{data::nombre-cliente}
	</Contenido>
	<Formato_Contenido>
		
	</Formato_Contenido>
	<Estilo_Encabezado>
		width:				100px;
		text-align:			center;
		background:			black;
		text-align:			center;
		font-weight:		bold;
		color:				rgb(255, 255, 255);
		background:			rgb(121, 183, 41);
	</Estilo_Encabezado>
	<Estilo_Contenido>
		padding-top:		0.4em;
		padding-bottom:		0.3em;
		padding-left: 0.5em;
		padding-right: 0.5em;
		border-top-color:	rgb(121, 183, 41);
		border-top-style:	solid;
		border-top-width:	1px;
		color:				rgb(0, 0, 0);
		background-color:	rgb(255, 255, 255);
		text-align:			center;
	</Estilo_Contenido>
</Columna>
<Columna>
	<Encabezado>
		Fecha de emisión
	</Encabezado>
	<Contenido>
		{data::fecha-emi}
	</Contenido>
	<Formato_Contenido>
		103
	</Formato_Contenido>
	<Estilo_Encabezado>
		width:				100px;
		text-align:			center;
		background:			black;
		text-align:			center;
		font-weight:		bold;
		color:				rgb(255, 255, 255);
		background:			rgb(121, 183, 41);
	</Estilo_Encabezado>
	<Estilo_Contenido>
		padding-top:		0.4em;
		padding-bottom:		0.3em;
		padding-left: 0.5em;
		padding-right: 0.5em;
		border-top-color:	rgb(121, 183, 41);
		border-top-style:	solid;
		border-top-width:	1px;
		color:				rgb(0, 0, 0);
		background-color:	rgb(255, 255, 255);
		text-align:			left;
	</Estilo_Contenido>
</Columna>
<Columna>
	<Encabezado>
		Tipo nota de pedido
	</Encabezado>
	<Contenido>
		{data::tipo-np}
	</Contenido>
	<Formato_Contenido>		
	</Formato_Contenido>
	<Estilo_Encabezado>
		width:				100px;
		text-align:			center;
		background:			black;
		text-align:			center;
		font-weight:		bold;
		color:				rgb(255, 255, 255);
		background:			rgb(121, 183, 41);
	</Estilo_Encabezado>
	<Estilo_Contenido>
		padding-top:		0.4em;
		padding-bottom:		0.3em;
		padding-left: 0.5em;
		padding-right: 0.5em;
		border-top-color:	rgb(121, 183, 41);
		border-top-style:	solid;
		border-top-width:	1px;
		color:				rgb(0, 0, 0);
		background-color:	rgb(255, 255, 255);
		text-align:			center;
	</Estilo_Contenido>
</Columna>
<Columna>
	<Encabezado>
		Nro. nota de pedido
	</Encabezado>
	<Contenido>
		{data::nro-np}
	</Contenido>
	<Formato_Contenido>		
	</Formato_Contenido>
	<Estilo_Encabezado>
		width:				100px;
		text-align:			center;
		background:			black;
		text-align:			center;
		font-weight:		bold;
		color:				rgb(255, 255, 255);
		background:			rgb(121, 183, 41);
	</Estilo_Encabezado>
	<Estilo_Contenido>
		padding-top:		0.4em;
		padding-bottom:		0.3em;
		padding-left: 0.5em;
		padding-right: 0.5em;
		border-top-color:	rgb(121, 183, 41);
		border-top-style:	solid;
		border-top-width:	1px;
		color:				rgb(0, 0, 0);
		background-color:	rgb(255, 255, 255);
		text-align:			center;
	</Estilo_Contenido>
</Columna>
<Columna>
    <Encabezado>
		Importe total sin impuestos
	</Encabezado>
	<Contenido>
		$ {data::imp-tot-sim}
	</Contenido>
	<Formato_Contenido>
	</Formato_Contenido>
	<Estilo_Encabezado>
		width:				100px;
		text-align:			center;
		background:			black;
		text-align:			center;
		font-weight:		bold;
		color:				rgb(255, 255, 255);
		background:			rgb(121, 183, 41);
	</Estilo_Encabezado>
	<Estilo_Contenido>
		padding-top:		0.4em;
		padding-bottom:		0.3em;
		padding-left: 0.5em;
		padding-right: 0.5em;
		border-top-color:	rgb(121, 183, 41);
		border-top-style:	solid;
		border-top-width:	1px;
		color:				rgb(0, 0, 0);
		background-color:	rgb(255, 255, 255);
		text-align:			center;
	</Estilo_Contenido>
</Columna>
'

--ESTILOS CSS DE LA PLANTILLA
--Por lo general es suficiente configurar los parametros de la plantilla para alterar el aspecto visual del mensaje
--Sin embargo, en ocasiones es necesario realizar cambios mas profundos para manipular aspectos de la apariencia del mensaje que no es posible
--realizar por medio de los parametros de la plantilla
--En este caso, los estilos de la plantilla permiten editar de manera mas concreta y especifica el aspecto por medio del estandar CSS
--Tenga en cuenta que los clientes de correo electronico tienen capacidades limitadas y que es posible que algunas propiedades css no funcionen como lo espera.
DECLARE @ESTILO_MENSAJE NVARCHAR(MAX) = '
    font-family: {style-value::tipografia};
	font-size: {style-value::tipografia-size};
	line-height: 1.5em;
	margin-bottom: 150px;
	margin-top: 50px;
	width: 900px;
	margin-left: auto;
	margin-right: auto;
'
DECLARE @ESTILO_ENCABEZADO NVARCHAR(MAX) ='
	width: 100%;
	display: block;
	text-align: right;
'
DECLARE @ESTILO_IMAGEN_ENCABEZADO NVARCHAR(MAX) = ''
DECLARE @ESTILO_PIE NVARCHAR(MAX) = '
	width: 100%;
	display: block;
	text-align: center;
'
DECLARE @ESTILO_IMAGEN_PIE NVARCHAR(MAX) = '
	width: 100%;
'

DECLARE @ESTILO_FIRMA NVARCHAR(MAX) = ''

--PLANTILLAS
--Son los fragmentos HTML que definen la estructura del mensaje
--En esta plantilla preparan los parametros, campos de datos, estilos y otras plantillas que deben incluirse en el mensaje
--Modificar la estructura impacta significativamente en el formato que toma el mensaje y en la posibilidad de que los clientes de correo puedan interpretar el mensaje correctamente
DECLARE @PLANTILLA_PRINCIPAL AS NVARCHAR(MAX) = '
	</p>
    <div style="{style::mensaje}">

		{template::head}

		{template::cuerpo-anexo}

        <p>{template-value::introduccion}</p>

        {template::table}

        <p style="{style::firma}">{template-value::firma}</p>

		{template::foot}
    </div>
	<p>
'

DECLARE @PLANTILLA_ENCABEZADO AS NVARCHAR(MAX) = '
    <div style="{style::encabezado}width: 100%; display: block; text-align: right;">
	    <img src="{template-value::imagen-encabezado}" style="{style::imagen-encabezado}" />
    </div>
'
DECLARE @PLANTILLA_PIE AS NVARCHAR(MAX) = '
	<div style="{style::pie}">
	    <img src="{template-value::imagen-pie}" style="{style::imagen-pie}" />
    </div>
'


--    _   ____                     _             _           _         _ 
--   (_) |  _ \                   | |           | |         (_)       | |
--   | | | |_) |_   _  ___ _ __   | |_ _ __ __ _| |__   __ _ _  ___   | |
--   | | |  _ <| | | |/ _ \ '_ \  | __| '__/ _` | '_ \ / _` | |/ _ \  | |
--   | | | |_) | |_| |  __/ | | | | |_| | | (_| | |_) | (_| | | (_) | |_|
--   |_| |____/ \__,_|\___|_| |_|  \__|_|  \__,_|_.__/ \__,_| |\___/  (_)
--                                                         _/ |          
--                                                        |__/           
--    _  _                                  _                                                                          __   _                                  
--   | \| |  ___     __ _   _  _   ___   __| |  __ _     _ __    __ _   ___    _ __   ___   _ _     __   ___   _ _    / _| (_)  __ _   _  _   _ _   __ _   _ _ 
--   | .` | / _ \   / _` | | || | / -_) / _` | / _` |   | '  \  / _` | (_-<   | '_ \ / _ \ | '_|   / _| / _ \ | ' \  |  _| | | / _` | | || | | '_| / _` | | '_|
--   |_|\_| \___/   \__, |  \_,_| \___| \__,_| \__,_|   |_|_|_| \__,_| /__/   | .__/ \___/ |_|     \__| \___/ |_||_| |_|   |_| \__, |  \_,_| |_|   \__,_| |_|  
--                     |_|                                                    |_|                                              |___/                           

SET NOCOUNT ON
BEGIN TRY
	DECLARE @RESUME VARCHAR(MAX) = ''
	DECLARE @WARNINGS VARCHAR(MAX) = ''

	BEGIN TRANSACTION TPES_TRANSACTION

	SET @RESUME = @RESUME + CHAR(13)+CHAR(10) + '=== CONFIGURACION DE TAREA PROGRAMADA ESTANDAR ==='
	SET @RESUME = @RESUME + CHAR(13)+CHAR(10) + '== Configurando la tarea programada #' + CAST(@CODIGO_TAREA_PROGRAMADA as varchar) + ' =='

	SET @RESUME = @RESUME + CHAR(13)+CHAR(10) + '>> Limpiando configuracion'
	DELETE dmTPES_PARM WHERE PARM_CODIGO_TPRG = @CODIGO_TAREA_PROGRAMADA
	DELETE dmTPES_COLS WHERE COLS_CODIGO_TPRG = @CODIGO_TAREA_PROGRAMADA

	SET @RESUME = @RESUME + CHAR(13)+CHAR(10) + '>> Estableciendo parametros de la TPES'
	INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{parameter::divisiones}', @COMPROBANTES_DIVISION)
	
	SET @RESUME = @RESUME + CHAR(13)+CHAR(10) + '>> Estableciendo parametros del destinatario'
	INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{parameter::asunto}', @ASUNTO)

	IF(@DESTINATARIOS_PARA IS NOT NULL AND LEN(LTRIM(RTRIM(@DESTINATARIOS_PARA))) > 0)
	BEGIN
		INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{parameter::para}', @DESTINATARIOS_PARA )
	END
	ELSE
	BEGIN
		INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{parameter::para}', NULL )
	END

	IF(@DESTINATARIOS_CC IS NOT NULL AND LEN(LTRIM(RTRIM(@DESTINATARIOS_CC))) > 0)
	BEGIN
		INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{parameter::cc}', @DESTINATARIOS_CC )
	END
	ELSE
	BEGIN
		INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{parameter::cc}', NULL )
	END

	IF(@DESTINATARIOS_CCO IS NOT NULL AND LEN(LTRIM(RTRIM(@DESTINATARIOS_CCO))) > 0)
	BEGIN
		INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{parameter::cco}', @DESTINATARIOS_CCO )
	END
	ELSE
	BEGIN
		INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{parameter::cco}', NULL )
	END

	INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{parameter::cuerpo-anexo}', @ANEXAR_CUERPO)


	SET @RESUME = @RESUME + CHAR(13)+CHAR(10) + '>> Estableciendo parametros del remitente'

	DECLARE @REMITENTE_LEGAJO VARCHAR(8) = ''

	IF @USAR_USUARIO_RESPONSABLE_DE_TPRG = 1
	BEGIN;
		SET @REMITENTE_LEGAJO = (SELECT TPRG_RESPONSABLE FROM TPRG_TPRG WHERE TPRG_CODIGO_TPRG = @CODIGO_TAREA_PROGRAMADA)

		IF @REMITENTE_LEGAJO IS NULL
		BEGIN
			RAISERROR ('El usuario responsable de la tarea programada %d no esta configurado', 16, 1, @CODIGO_TAREA_PROGRAMADA);  
		END

		SET @USUARIO_PLATAFORMA = (
			SELECT
				PERS_USUARIO
			FROM TPRG_TPRG
			INNER JOIN ACCT_PERS
				ON PERS_LEGAJO = TPRG_RESPONSABLE
				AND TPRG_CODIGO_TPRG = @CODIGO_TAREA_PROGRAMADA
		)
		IF @USUARIO_PLATAFORMA IS NULL
		BEGIN
			RAISERROR ('El usuario responsable de la tarea programada %d no tiene un legajo vinculado', 16, 1, @CODIGO_TAREA_PROGRAMADA);  
		END
	END
	ELSE
	BEGIN
		IF @CLASE_DE_REMITENTE = 'USUARIO PLATAFORMA'
		BEGIN
			SET @REMITENTE_LEGAJO = (SELECT PERS_LEGAJO FROM ACCT_PERS WHERE PERS_USUARIO = @USUARIO_PLATAFORMA)
			IF @REMITENTE_LEGAJO IS NULL
			BEGIN
				RAISERROR ('El usuario de plataforma %s no tiene un legajo', 16, 2, @USUARIO_PLATAFORMA);
			END
		END
	END
	
	IF (SELECT COUNT(*) FROM SEGU_USUA WHERE USUA_USUARIO = @USUARIO_PLATAFORMA) = 0
	BEGIN
		RAISERROR ('El usuario de plataforma %s no existe', 16, 1, @USUARIO_PLATAFORMA);
	END

	DECLARE @REMITENTE_PASSWORD VARCHAR(50) = NULL

	IF @CLASE_DE_REMITENTE <> 'INSTITUCIONAL' AND @CLASE_DE_REMITENTE <> 'USUARIO PLATAFORMA' AND @CLASE_DE_REMITENTE <> 'POR DOMINIO'
	BEGIN
		RAISERROR ('La clase de remitente configurada "%s" no es una clase valida. Las clases permitidas son: "INSTITUCIONAL", "USUARIO PLATAFORMA", "POR DOMINIO"', 16, 1, @CLASE_DE_REMITENTE);
	END

	IF @CLASE_DE_REMITENTE = 'USUARIO PLATAFORMA'
	BEGIN
		SET @REMITENTE = (SELECT PERS_CUENTA_SMTP FROM ACCT_PERS WHERE PERS_LEGAJO = @REMITENTE_LEGAJO)
	END

	INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{parameter::usuario-plataforma}', @USUARIO_PLATAFORMA);
	INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{parameter::remitente-cuenta}', @REMITENTE);
	INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{parameter::remitente-clase}', @CLASE_DE_REMITENTE);

	SET @RESUME = @RESUME + CHAR(13)+CHAR(10) + '>> Configurando parametros de la plantilla'
	INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{template-value::introduccion}', @INTRODUCCION)
	INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{template-value::firma}', @FIRMA)
	INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{template-value::imagen-encabezado}', @URL_IMAGEN_MENSAJE_ENCABEZADO)
	INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{template-value::imagen-pie}', @URL_IMAGEN_MENSAJE_PIE)

	INSERT INTO dmTPES_COLS (
		COLS_CODIGO_TPRG
		,COLS_TABLA
		,COLS_ORDEN
		,COLS_ENCABEZADO
		,COLS_ESTILO_ENCABEZADO
		,COLS_ESTILO_VALOR
		,COLS_VALOR
		,COLS_FORMATO
	)
	SELECT
		@CODIGO_TAREA_PROGRAMADA
		,1
		,ROW_NUMBER() OVER (ORDER BY temp.Col)
		,REPLACE( REPLACE( temp.Col.value('Encabezado[1]', 'varchar(max)'), CHAR(9), ''), CHAR(10), '')
		,REPLACE( REPLACE( temp.Col.value('Estilo_Encabezado[1]', 'varchar(max)') , CHAR(9), ''), CHAR(10), '')
		,REPLACE( REPLACE( temp.Col.value('Estilo_Contenido[1]', 'varchar(max)') , CHAR(9), ''), CHAR(10), '')
		,REPLACE( REPLACE( temp.Col.value('Contenido[1]', 'varchar(max)') , CHAR(9), ''), CHAR(10), '')
		,REPLACE( REPLACE( temp.Col.value('Formato_Contenido[1]', 'varchar(max)') , CHAR(9), ''), CHAR(10), '')
	FROM   @COLUMNAS.nodes('//Columna') temp(Col)

	SET @RESUME = @RESUME + CHAR(13)+CHAR(10) + '>> Configurando estilos de la plantilla'

	INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{style-value::tipografia}', @TIPOGRAFIA )
	INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{style-value::tipografia-size}', @TIPOGRAFIA_ALTURA )
	
	INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{style::mensaje}', @ESTILO_MENSAJE)
	INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{style::encabezado}', @ESTILO_ENCABEZADO)
	INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{style::imagen-encabezado}', @ESTILO_IMAGEN_ENCABEZADO)
	INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{style::pie}', @ESTILO_PIE)
	INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{style::imagen-pie}', @ESTILO_IMAGEN_PIE)

	INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{style::firma}', @ESTILO_FIRMA)
	
	--aplico los valores de estilos dentro de los estilo
	EXEC dbo.dmTPES_spAplicarValoresDeEstilos @CODIGO_TAREA_PROGRAMADA


	SET @RESUME = @RESUME + CHAR(13)+CHAR(10) + '>> Configurando estructuras de la plantilla'
	INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{template::main}', @PLANTILLA_PRINCIPAL)

	IF( LEN(LTRIM(RTRIM(@URL_IMAGEN_MENSAJE_ENCABEZADO)))=0 )
	BEGIN
		SET @PLANTILLA_ENCABEZADO = NULL
	END

	IF( LEN(LTRIM(RTRIM(@URL_IMAGEN_MENSAJE_PIE)))=0 )
	BEGIN
		SET @PLANTILLA_PIE = NULL
	END

	INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{template::head}', @PLANTILLA_ENCABEZADO)
	INSERT INTO dmTPES_PARM ( PARM_CODIGO_TPRG, PARM_CLAVE, PARM_VALOR ) VALUES (@CODIGO_TAREA_PROGRAMADA, '{template::foot}', @PLANTILLA_PIE)

	--aplico los estilos dentro de las plantillas
	EXEC dbo.dmTPES_spAplicarEstilos @CODIGO_TAREA_PROGRAMADA
	--aplico los valores de plantillas dentro de las plantillas
	EXEC dbo.dmTPES_spAplicarValoresDePlantillas @CODIGO_TAREA_PROGRAMADA

	SET @RESUME = @RESUME + CHAR(13)+CHAR(10) + 'La carga de la configuracion para la TPES #' + CAST (@CODIGO_TAREA_PROGRAMADA as VARCHAR) + ' fue un exito'

	COMMIT TRANSACTION TPES_TRANSACTION

	PRINT @RESUME
	PRINT @WARNINGS

	PRINT 'FINALIZADO CORRECTAMENTE'

END TRY
BEGIN CATCH

	PRINT @RESUME
	PRINT @WARNINGS

	DECLARE @ERROR VARCHAR(4000)	= 'LN:' + RIGHT('00000'+CONVERT(VARCHAR(5), ISNULL(ERROR_LINE(), 0)), 5) + ' ' +	ERROR_MESSAGE();
	DECLARE @ERROR_SEVERETY INT = ERROR_SEVERITY()
	DECLARE @ERROR_STATE INT = ERROR_STATE()

	RAISERROR (@ERROR, @ERROR_SEVERETY, @ERROR_STATE);

	IF (@@TRANCOUNT > 0)
	BEGIN
		ROLLBACK TRANSACTION TPES_TRANSACTION
		PRINT 'Error detectado. Todos los cambios de la configuracion fueron revertidos.'
	END

	PRINT 'FINALIZADO CON ERRORES'

END CATCH
