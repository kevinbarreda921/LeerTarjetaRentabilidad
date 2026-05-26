using LeerTarjetasRentabilidad;

// 1. Procesar las tarjetas de rentabilidad existentes (original)
// TarjetasProcesador.Procesar();

// 2. Generar el archivo JSON de configuración con todas las hojas del Excel vacío de rentabilidad
// RentabilidadReader.GenerarConfiguracionHojas();

// 3. Procesar y extraer los días de rentabilidad (Marzo y Abril 2026) desde el Excel de rentabilidad
// RentabilidadDiasReader.ProcesarDias();

// 4. Escribir los datos de líquidos y descuentos en el Excel de rentabilidad usando EPPlus
RentabilidadExcelWriter.EscribirDatosRentabilidad();
