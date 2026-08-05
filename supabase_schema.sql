-- Esquema para Supabase (Postgres). Opcional: la app las crea sola al arrancar.

CREATE TABLE cotizaciones (
	id SERIAL NOT NULL, 
	proveedor TEXT, 
	contacto TEXT, 
	email TEXT, 
	whatsapp TEXT, 
	cliente TEXT, 
	fecha_emision TEXT, 
	incoterm TEXT, 
	moneda TEXT, 
	lead_time TEXT, 
	forma_pago TEXT, 
	transporte TEXT, 
	empaque TEXT, 
	total_exw FLOAT, 
	archivo TEXT, 
	creado_en TEXT, 
	PRIMARY KEY (id)
);

CREATE TABLE lineas (
	id SERIAL NOT NULL, 
	cotizacion_id INTEGER, 
	sn TEXT, 
	descripcion TEXT, 
	categoria TEXT, 
	cantidad FLOAT, 
	unidad TEXT, 
	precio_unit FLOAT, 
	total FLOAT, 
	imagen BYTEA, 
	estado TEXT, 
	nota TEXT, 
	etapa TEXT, 
	guia TEXT, 
	contenedor TEXT, 
	tracking_num TEXT, 
	eta TEXT, 
	fecha_pedido TEXT, 
	fecha_recibido TEXT, 
	creado_en TEXT, 
	PRIMARY KEY (id), 
	FOREIGN KEY(cotizacion_id) REFERENCES cotizaciones (id) ON DELETE CASCADE
);
