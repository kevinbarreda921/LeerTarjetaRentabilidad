using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeOpenXml;

namespace LeerTarjetasRentabilidad;

public class RentabilidadExcelWriter
{
    public static void EscribirDatosRentabilidad()
    {
        // Establecer el contexto de licencia de EPPlus para la versión 8+
        ExcelPackage.License.SetNonCommercialPersonal("Kevin Barreda");

        string rutaProyecto = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\"));
        string rutaArchivoExcel = Path.Combine(rutaProyecto, "ArchivosExcel", "Rentabilidad", "renabilidad vacio.xlsx");
        string rutaResultadoGrifos = Path.Combine(rutaProyecto, "ResultadoGrifos.json");
        string rutaFechasEscribir = Path.Combine(rutaProyecto, "fechasEscribir.json");

        Console.WriteLine($"\n[PROCESO] Iniciando escritura de rentabilidad con EPPlus...");
        Console.WriteLine($"Excel destino: {rutaArchivoExcel}");

        if (!File.Exists(rutaArchivoExcel))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"[ERROR] No se encontró el archivo Excel destino en: {rutaArchivoExcel}");
            Console.ResetColor();
            return;
        }

        if (!File.Exists(rutaResultadoGrifos))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"[ERROR] No se encontró el archivo de origen de datos en: {rutaResultadoGrifos}");
            Console.ResetColor();
            return;
        }

        if (!File.Exists(rutaFechasEscribir))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"[ERROR] No se encontró el archivo de mapeo de filas en: {rutaFechasEscribir}");
            Console.ResetColor();
            return;
        }

        try
        {
            // 1. Cargar y deserializar los JSONs
            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };

            string grifosJsonText = File.ReadAllText(rutaResultadoGrifos);
            var datosGrifos = JsonSerializer.Deserialize<Dictionary<string, List<GrifoData>>>(grifosJsonText, options);

            string filasJsonText = File.ReadAllText(rutaFechasEscribir);
            var mapeoFilas = JsonSerializer.Deserialize<Dictionary<string, List<FilaData>>>(filasJsonText, options);

            if (datosGrifos == null || datosGrifos.Count == 0)
            {
                Console.WriteLine("[ERROR] El archivo ResultadoGrifos.json no contiene datos válidos.");
                return;
            }

            if (mapeoFilas == null || mapeoFilas.Count == 0)
            {
                Console.WriteLine("[ERROR] El archivo fechasEscribir.json no contiene datos válidos.");
                return;
            }

            // 2. Abrir el archivo Excel con EPPlus
            var fileInfo = new FileInfo(rutaArchivoExcel);
            using (var package = new ExcelPackage(fileInfo))
            {
                int totalHojasProcesadas = 0;
                int totalFilasEscritas = 0;

                // 3. Iterar por cada hoja presente en ResultadoGrifos
                foreach (var kvp in datosGrifos)
                {
                    string sheetName = kvp.Key;
                    var listaDiasData = kvp.Value;

                    // Buscar si esta hoja está en el mapeo de filas
                    var filaKey = mapeoFilas.Keys.FirstOrDefault(k => k.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
                    if (filaKey == null)
                    {
                        continue;
                    }

                    var listaDiasFilas = mapeoFilas[filaKey];

                    // Buscar la hoja de cálculo en el archivo Excel (EPPlus)
                    var worksheet = package.Workbook.Worksheets[sheetName];
                    if (worksheet == null)
                    {
                        // Buscar de manera case-insensitive
                        worksheet = package.Workbook.Worksheets.FirstOrDefault(w => w.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
                    }

                    if (worksheet == null)
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"[ADVERTENCIA] La hoja '{sheetName}' no existe en el Excel. Saltando...");
                        Console.ResetColor();
                        continue;
                    }

                    bool hojaModificada = false;

                    // 4. Mapear y escribir cada día en su fila correspondiente
                    foreach (var grifoDia in listaDiasData)
                    {
                        // Buscar el día equivalente en fechasEscribir
                        var filaDia = listaDiasFilas.FirstOrDefault(f => f.FechaHoja != null && f.FechaHoja.Equals(grifoDia.FechaHoja, StringComparison.OrdinalIgnoreCase));
                        
                        if (filaDia != null && filaDia.Fila > 0)
                        {
                            int filaDestino = filaDia.Fila;

                            // AK = Columna 37, AN = Columna 40
                            worksheet.Cells[filaDestino, 37].Value = (double)grifoDia.DataLiquidos;
                            worksheet.Cells[filaDestino, 40].Value = (double)grifoDia.DataDescuento;

                            totalFilasEscritas++;
                            hojaModificada = true;
                        }
                    }

                    if (hojaModificada)
                    {
                        totalHojasProcesadas++;
                    }
                }

                // 5. Guardar los cambios
                if (totalFilasEscritas > 0)
                {
                    Console.WriteLine("[PROCESO] Guardando cambios en el Excel...");
                    package.Save();

                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"\n[ÉXITO] Escritura de Excel completada.");
                    Console.WriteLine($"Hojas con datos escritos: {totalHojasProcesadas}");
                    Console.WriteLine($"Total de filas/celdas modificadas: {totalFilasEscritas}");
                    Console.ResetColor();
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("\n[ADVERTENCIA] No se encontró ninguna coincidencia de hoja y fecha para escribir en el Excel.");
                    Console.ResetColor();
                }
            }
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"\n[ERROR] Ocurrió un error inesperado durante la escritura en el Excel: {ex.Message}");
            Console.ResetColor();
        }
    }
}

public class GrifoData
{
    [JsonPropertyName("fecha_hoja")]
    public string? FechaHoja { get; set; }

    [JsonPropertyName("DataLiquidos")]
    public decimal DataLiquidos { get; set; }

    [JsonPropertyName("DataDescuento")]
    public decimal DataDescuento { get; set; }
}

public class FilaData
{
    [JsonPropertyName("fecha_hoja")]
    public string? FechaHoja { get; set; }

    [JsonPropertyName("Fila")]
    public int Fila { get; set; }
}
