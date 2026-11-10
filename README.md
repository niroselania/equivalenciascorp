# Equivalencia de códigos de barra

Aplicación web para completar códigos de barra en una planilla de pedido.

## Uso

1. Cargá la planilla de equivalencia (`.xls` o `.xlsx`) que contiene el código de barras, SKU, color y talle.
2. Subí el pedido (`.xlsx`) con las columnas `SKU`, `COLOR` o `COD.COLOR`, y `TALLE`.
3. Descargá el resultado. La aplicación agrega `CODIGO DE BARRAS` como nueva columna A y conserva el resto de la planilla.

Cuando no encuentra una coincidencia exacta por SKU, color y talle, escribe `SIN COINCIDENCIA`.

## Portainer

Usá el archivo `docker-compose.yml` para desplegar el stack. La aplicación guarda la planilla de equivalencia cargada en un volumen Docker, para reutilizarla en próximos pedidos.
