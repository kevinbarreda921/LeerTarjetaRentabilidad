using LeerTarjetasRentabilidad;

// 1. Procesar las tarjetas de rentabilidad existentes (original)
// TarjetasProcesador.Procesar();

// 2. Generar el archivo JSON de configuración con todas las hojas del Excel vacío de rentabilidad
//RentabilidadReader.GenerarConfiguracionHojas();

// 3. Procesar y extraer los días de rentabilidad según los rangos configurados en el JSON
RentabilidadDiasReader.ProcesarDias();
